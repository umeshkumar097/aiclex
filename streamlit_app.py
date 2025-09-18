# streamlit_app.py
"""
Aiclex — Hall Ticket Mailer (Auto-fill recipients fixed)
- Upload Excel + ZIP of PDFs
- Map columns: Hallticket, Candidate Email, Location
- Choose recipient source: coordinator column / specific column / aggregate candidate emails
- Email extraction via regex (handles names+emails, multiple emails)
- Auto-populates per-location recipient fields when a recipient source is selected
- Create location-wise chunked ZIPs, send via SMTP with subject/body/footer templates
- Test mode and per-candidate sends supported
"""
import streamlit as st
import pandas as pd
import zipfile, os, io, tempfile, shutil, time, re
from collections import defaultdict, Counter
from datetime import datetime
import smtplib
from email.message import EmailMessage

# ---------------- Helpers ----------------
EMAIL_REGEX = re.compile(r'[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}', re.UNICODE)

def human_bytes(n):
    try:
        n = float(n)
    except:
        return ""
    for unit in ['B','KB','MB','GB','TB']:
        if n < 1024:
            return f"{n:.2f} {unit}"
        n /= 1024
    return f"{n:.2f} TB"

def looks_like_email(s):
    return bool(EMAIL_REGEX.findall(str(s))) if s else False

def extract_emails_from_text(s):
    if not s:
        return []
    return EMAIL_REGEX.findall(str(s))

def extract_zip_recursively(zip_file_like, extract_to):
    with zipfile.ZipFile(zip_file_like) as z:
        z.extractall(path=extract_to)
    for root, _, files in os.walk(extract_to):
        for f in files:
            if f.lower().endswith('.zip'):
                nested = os.path.join(root, f)
                nested_dir = os.path.join(root, f"_nested_{f}")
                os.makedirs(nested_dir, exist_ok=True)
                with open(nested, 'rb') as nf:
                    extract_zip_recursively(nf, nested_dir)

def create_chunked_zips(file_paths, out_dir, base_name, max_bytes):
    os.makedirs(out_dir, exist_ok=True)
    parts = []
    current_files = []
    part_index = 1
    for fp in file_paths:
        current_files.append(fp)
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', compression=zipfile.ZIP_DEFLATED) as z:
            for f in current_files:
                z.write(f, arcname=os.path.basename(f))
        size = buf.getbuffer().nbytes
        if size <= max_bytes:
            continue
        else:
            last = current_files.pop()
            if current_files:
                part_path = os.path.join(out_dir, f"{base_name}_part{part_index}.zip")
                with zipfile.ZipFile(part_path, 'w', compression=zipfile.ZIP_DEFLATED) as z:
                    for f in current_files:
                        z.write(f, arcname=os.path.basename(f))
                parts.append({"path": part_path, "num_files": len(current_files), "size": os.path.getsize(part_path)})
                part_index += 1
            current_files = [last]
    if current_files:
        part_path = os.path.join(out_dir, f"{base_name}_part{part_index}.zip")
        with zipfile.ZipFile(part_path, 'w', compression=zipfile.ZIP_DEFLATED) as z:
            for f in current_files:
                z.write(f, arcname=os.path.basename(f))
        parts.append({"path": part_path, "num_files": len(current_files), "size": os.path.getsize(part_path)})
    return parts

def send_email(smtp_cfg, to_email, subject, body, attachment_path=None, attachment_name=None):
    msg = EmailMessage()
    msg['From'] = smtp_cfg['sender']
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.set_content(body)
    if attachment_path:
        with open(attachment_path, 'rb') as f:
            data = f.read()
        maintype = 'application'
        subtype = 'zip' if (attachment_name or "").lower().endswith('.zip') else 'octet-stream'
        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=attachment_name or os.path.basename(attachment_path))
    if smtp_cfg.get('use_ssl', True):
        server = smtplib.SMTP_SSL(smtp_cfg['host'], smtp_cfg['port'], timeout=60)
    else:
        server = smtplib.SMTP(smtp_cfg['host'], smtp_cfg['port'], timeout=60)
        server.starttls()
    if smtp_cfg.get('password'):
        server.login(smtp_cfg['sender'], smtp_cfg['password'])
    server.send_message(msg)
    server.quit()

# ---------------- UI ----------------
st.set_page_config(page_title="Aiclex Mailer", layout="wide")
st.title("📧 Aiclex Technologies — Hall Ticket Mailer (Auto-fill)")

# Sidebar
with st.sidebar:
    st.header("SMTP Settings")
    smtp_host = st.text_input("SMTP host", value="smtp.hostinger.com")
    smtp_port = st.number_input("SMTP port", value=465)
    protocol = st.selectbox("Protocol", ["SMTPS (SSL)", "SMTP (STARTTLS)"], index=0)
    sender_email = st.text_input("Sender email", value="info@aiclex.in")
    sender_password = st.text_input("Sender password", type="password")
    st.markdown("---")
    st.header("Email Templates")
    subject_template = st.text_input("Subject template (use {location})", value="Hall Tickets — {location}")
    body_template = st.text_area("Body template (use {location} and {footer})", value="Dear Coordinator,\n\nPlease find attached the hall tickets for {location}.\n\n{footer}", height=160)
    footer_text = st.text_area("Footer text", value="Regards,\nAiclex Technologies\ninfo@aiclex.in", height=80)
    st.markdown("---")
    st.header("Options")
    delay_seconds = st.number_input("Delay between emails (s)", value=2.0, step=0.5)
    attachment_limit_mb = st.number_input("Attachment limit (MB)", value=3.0, step=0.5)
    test_mode = st.checkbox("Test mode (send all to test address)", value=True)
    test_email = st.text_input("Test recipient email", value="info@aiclex.in")
    send_candidates = st.checkbox("Also send per-candidate emails", value=False)

# 1) Upload
st.header("1) Upload Excel and ZIP")
uploaded_excel = st.file_uploader("Upload Excel (.xlsx/.csv) with Hallticket, Email, Location columns", type=["xlsx","csv"])
uploaded_zip = st.file_uploader("Upload master ZIP of PDFs (nested zips allowed)", type=["zip"])

if not uploaded_excel or not uploaded_zip:
    st.info("Please upload both Excel and ZIP to proceed.")
    st.stop()

# Read excel
try:
    if uploaded_excel.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_excel, dtype=str).fillna("")
    else:
        df = pd.read_excel(uploaded_excel, dtype=str).fillna("")
except Exception as e:
    st.error("Failed to read Excel: " + str(e))
    st.stop()

# Extract zip
tmp_zip_dir = tempfile.mkdtemp(prefix="aiclex_zip_")
try:
    bio = io.BytesIO(uploaded_zip.read())
    extract_zip_recursively(bio, tmp_zip_dir)
except Exception as e:
    st.error("Failed to extract ZIP: " + str(e))
    st.stop()

# pdf map
pdf_map = {}
for root, _, files in os.walk(tmp_zip_dir):
    for f in files:
        if f.lower().endswith(".pdf"):
            pdf_map[f] = os.path.join(root, f)
st.success(f"Loaded Excel ({len(df)} rows) and {len(pdf_map)} PDFs")

# 2) Column mapping
st.header("2) Map columns")
cols = list(df.columns)
if not cols:
    st.error("Excel has no columns.")
    st.stop()

ht_col = st.selectbox("Hallticket column", cols, index=0)
cand_email_col = st.selectbox("Candidate Email column", cols, index=1 if len(cols)>1 else 0)
loc_col = st.selectbox("Location column", cols, index=2 if len(cols)>2 else 0)

st.subheader("Preview")
preview_rows = st.number_input("Preview rows", min_value=5, max_value=2000, value=50, step=5)
st.dataframe(df.head(preview_rows), use_container_width=True)

# 3) Recipient source & extraction
st.header("3) Recipient source & auto-fill")

candidate_recipient_cols = [c for c in cols if any(k in c.lower() for k in ("recipient","coord","contact","coordinator","manager","contact_person"))]

recip_source_option = st.selectbox("Recipient source", [
    "Coordinator/Recipient column (use exact column below)",
    "Choose a specific column (pick below)",
    "Aggregate candidate emails per location (use candidate email column)"
])

chosen_coord_col = None
chosen_spec_col = None
if recip_source_option == "Coordinator/Recipient column (use exact column below)":
    if candidate_recipient_cols:
        chosen_coord_col = st.selectbox("Choose coordinator/recipient column (expected to contain emails)", ["--none--"] + candidate_recipient_cols)
        if chosen_coord_col == "--none--":
            chosen_coord_col = None
    else:
        st.info("No coordinator-like columns detected. Pick 'Choose a specific column' or use aggregate candidate emails.")
elif recip_source_option == "Choose a specific column (pick below)":
    chosen_spec_col = st.selectbox("Pick the column that contains recipient EMAILs", ["--none--"] + cols)
    if chosen_spec_col == "--none--":
        chosen_spec_col = None

# locations
df[loc_col] = df[loc_col].astype(str).str.strip()
location_values = sorted(df[loc_col].fillna("(empty)").unique())

# Function: build default_recips by extracting emails
def build_default_recips():
    defaults = {}
    if chosen_coord_col:
        for loc in location_values:
            vals = df[df[loc_col] == loc][chosen_coord_col].astype(str).tolist()
            found = []
            for v in vals:
                for e in extract_emails_from_text(v):
                    if e not in found:
                        found.append(e)
            defaults[loc] = ";".join(found)
    elif chosen_spec_col:
        for loc in location_values:
            vals = df[df[loc_col] == loc][chosen_spec_col].astype(str).tolist()
            found = []
            for v in vals:
                for e in extract_emails_from_text(v):
                    if e not in found:
                        found.append(e)
            defaults[loc] = ";".join(found)
    else:
        use_suggest = st.checkbox("Auto-suggest from candidate email column (top N)", value=True)
        top_n = st.number_input("Top N emails per location to suggest", min_value=1, max_value=10, value=3)
        for loc in location_values:
            if use_suggest:
                rows = df[df[loc_col] == loc]
                extracted = []
                for v in rows[cand_email_col].astype(str).tolist():
                    extracted += extract_emails_from_text(v)
                cnt = Counter([e for e in extracted if e])
                top = [e for e,_ in cnt.most_common(top_n)]
                defaults[loc] = ";".join(top)
            else:
                defaults[loc] = ""
    return defaults

default_recips = build_default_recips()

# DEBUG PANEL: show counts & samples
st.subheader("Auto-fill debug (extracted recipients preview)")
debug_rows = []
for loc in location_values:
    vals = default_recips.get(loc,"").split(";") if default_recips.get(loc) else []
    debug_rows.append({"Location": loc, "ExtractedCount": len(vals), "Sample": ";".join(vals[:3])})
st.dataframe(pd.DataFrame(debug_rows).head(200), use_container_width=True)

# 4) Auto-populate session_state recipients if empty (this is the fix)
if "location_recipients" not in st.session_state:
    st.session_state["location_recipients"] = {}

# If the user changed recipient source, we want to fill only empty textareas (so user edits are preserved)
# We'll set a marker in session_state when we last auto-filled so that repeated runs don't overwrite manual edits.
last_fill_key = "_last_autofill_signature"
# signature = chosen selection + chosen column name
signature = f"{recip_source_option}::{chosen_coord_col or chosen_spec_col or cand_email_col}"

if st.session_state.get(last_fill_key) != signature:
    # auto-fill where empty
    for loc in location_values:
        cur = st.session_state["location_recipients"].get(loc, "").strip()
        if not cur and default_recips.get(loc):
            st.session_state["location_recipients"][loc] = default_recips.get(loc)
    st.session_state[last_fill_key] = signature

st.header("4) Edit recipients per location (auto-populated above)")
st.markdown("Edit recipients if needed. Use semicolon `;` or comma `,` separators. Only valid emails will be used for sending.")

for loc in location_values:
    key = f"recips__{loc}"
    initial = st.session_state["location_recipients"].get(loc, default_recips.get(loc,""))
    val = st.text_area(f"Recipients for {loc} (semicolon separated)", value=initial, key=key, height=70)
    # save edited value
    st.session_state["location_recipients"][loc] = val.strip()

# quick validation
invalid_locs = []
for loc in location_values:
    raw = st.session_state["location_recipients"].get(loc,"")
    items = [x.strip() for x in re.split(r"[;,\n]+", raw) if x.strip()]
    emails = [x for x in items if looks_like_email(x)]
    if raw and not emails:
        invalid_locs.append(loc)
if invalid_locs:
    st.warning("Some locations have recipient text but no valid emails: " + ", ".join(invalid_locs))

# 5) Mapping summary (match PDFs)
st.header("5) Mapping summary")
candidate_map = defaultdict(list)
grouped_pdfs = defaultdict(list)
for _, row in df.iterrows():
    ht = str(row.get(ht_col,"")).strip()
    em = str(row.get(cand_email_col,"")).strip()
    loc = str(row.get(loc_col,"")).strip()
    matched = None
    if ht:
        candidates = [f"{ht}.pdf", f"{ht.upper()}.pdf", f"{ht.lower()}.pdf"]
        for c in candidates:
            if c in pdf_map:
                matched = pdf_map[c]; break
        if not matched:
            for fn,p in pdf_map.items():
                if ht in fn:
                    matched = p; break
    if matched:
        grouped_pdfs[loc].append(matched)
    candidate_map[em].append({"hallticket": ht, "pdf": matched, "location": loc})

summary_rows = []
for loc in location_values:
    files = [p for p in grouped_pdfs.get(loc,[]) if p]
    total_bytes = sum(os.path.getsize(p) for p in files) if files else 0
    summary_rows.append({"Location": loc, "Rows": int((df[loc_col]==loc).sum()), "MatchedPDFs": len(files), "TotalSize": human_bytes(total_bytes), "RecipientsPreview": st.session_state["location_recipients"].get(loc,"")})
st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)

# 6) Prepare & Send
st.header("6) Prepare & Send")

def validate_all_recipients():
    errs = []
    for loc in location_values:
        raw = st.session_state["location_recipients"].get(loc,"")
        items = [x.strip() for x in re.split(r"[;,\n]+", raw) if x.strip()]
        valid = [x for x in items if looks_like_email(x)]
        if raw and not valid:
            errs.append(f"{loc}: no valid emails in '{raw[:100]}'")
    return (len(errs)==0, errs)

if st.button("Prepare & Send All (use Test Mode during checks)"):
    ok, errs = validate_all_recipients()
    if not ok:
        st.error("Recipient validation failed. Fix recipient entries first:\n" + "\n".join(errs))
        st.stop()

    smtp_cfg = {"host": smtp_host, "port": int(smtp_port), "use_ssl": True if protocol.startswith("SMTPS") else False, "sender": sender_email, "password": sender_password}
    try:
        if smtp_cfg["use_ssl"]:
            srv = smtplib.SMTP_SSL(smtp_cfg["host"], smtp_cfg["port"], timeout=30)
        else:
            srv = smtplib.SMTP(smtp_cfg["host"], smtp_cfg["port"], timeout=30)
            if smtp_cfg["port"] == 587:
                srv.starttls()
        if smtp_cfg.get("password"):
            srv.login(smtp_cfg["sender"], smtp_cfg["password"])
        srv.quit()
        st.success("SMTP login OK")
    except Exception as e:
        st.error("SMTP login failed: " + str(e))
        st.stop()

    logs = []
    max_bytes = int(float(attachment_limit_mb) * 1024 * 1024)
    for loc in location_values:
        pdfs = [p for p in grouped_pdfs.get(loc,[]) if p]
        raw = st.session_state["location_recipients"].get(loc,"")
        recipients = [r.strip() for r in re.split(r"[;,\n]+", raw) if looks_like_email(r.strip())]
        if test_mode and test_email:
            recipients = [test_email]
        if not recipients:
            logs.append({"Location": loc, "Status": "Skipped", "Reason": "No recipients/invalid emails"})
            continue
        if not pdfs:
            logs.append({"Location": loc, "Status": "Skipped", "Reason": "No PDFs matched"})
            continue
        outdir = tempfile.mkdtemp(prefix=f"loc_{re.sub(r'[^a-zA-Z0-9]','_',loc)}_")
        parts = create_chunked_zips(pdfs, out_dir=outdir, base_name=re.sub(r'[^a-zA-Z0-9]','_',loc)[:40], max_bytes=max_bytes)
        if not parts:
            zpath = os.path.join(outdir, f"{re.sub(r'[^a-zA-Z0-9]','_',loc)}.zip")
            with zipfile.ZipFile(zpath,'w', compression=zipfile.ZIP_DEFLATED) as z:
                for p in pdfs:
                    z.write(p, arcname=os.path.basename(p))
            parts = [{"path": zpath, "num_files": len(pdfs), "size": os.path.getsize(zpath)}]

        for recipient in recipients:
            for idx,p in enumerate(parts, start=1):
                subj = subject_template.format(location=loc)
                body = body_template.format(location=loc, footer=footer_text)
                try:
                    send_email(smtp_cfg, recipient, subj, body, attachment_path=p["path"], attachment_name=os.path.basename(p["path"]))
                    logs.append({"Location": loc, "Recipient": recipient, "Part": idx, "Zip": os.path.basename(p["path"]), "Status": "Sent", "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
                except Exception as e:
                    logs.append({"Location": loc, "Recipient": recipient, "Part": idx, "Zip": os.path.basename(p["path"]), "Status": "Failed", "Error": str(e)})
                time.sleep(float(delay_seconds))
    st.subheader("Send logs")
    st.dataframe(pd.DataFrame(logs), use_container_width=True)
    st.success("Finished sending. Check logs and your inbox (test mode).")

# 7) Optional per-candidate sends
if send_candidates:
    st.header("7) Send per-candidate emails (optional)")
    if st.button("Send per-candidate emails now"):
        smtp_cfg = {"host": smtp_host, "port": int(smtp_port), "use_ssl": True if protocol.startswith("SMTPS") else False, "sender": sender_email, "password": sender_password}
        try:
            if smtp_cfg["use_ssl"]:
                srv = smtplib.SMTP_SSL(smtp_cfg["host"], smtp_cfg["port"], timeout=30)
            else:
                srv = smtplib.SMTP(smtp_cfg["host"], smtp_cfg["port"], timeout=30)
                if smtp_cfg["port"] == 587:
                    srv.starttls()
            if smtp_cfg.get("password"):
                srv.login(smtp_cfg["sender"], smtp_cfg["password"])
            srv.quit()
        except Exception as e:
            st.error("SMTP login failed: " + str(e))
            st.stop()

        cand_logs = []
        max_bytes = int(float(attachment_limit_mb) * 1024 * 1024)
        for cand_email, items in candidate_map.items():
            to_addr = test_email if test_mode and test_email else cand_email
            pdfs = [it["pdf"] for it in items if it.get("pdf")]
            if not pdfs:
                cand_logs.append({"Recipient": to_addr, "Status": "Skipped", "Reason": "No PDFs"})
                continue
            total_size = sum(os.path.getsize(p) for p in pdfs)
            if total_size <= max_bytes:
                tmpdir = tempfile.mkdtemp(prefix="cand_")
                zpath = os.path.join(tmpdir, f"{re.sub(r'[^a-zA-Z0-9]','_',to_addr)}.zip")
                with zipfile.ZipFile(zpath,'w', compression=zipfile.ZIP_DEFLATED) as z:
                    for p in pdfs:
                        z.write(p, arcname=os.path.basename(p))
                subj = "Your Hall Ticket(s)"
                body = f"Dear Candidate,\n\nPlease find attached your hall ticket(s).\n\n{footer_text}"
                try:
                    send_email(smtp_cfg, to_addr, subj, body, attachment_path=zpath, attachment_name=os.path.basename(zpath))
                    cand_logs.append({"Recipient": to_addr, "Zip": os.path.basename(zpath), "Status": "Sent"})
                except Exception as e:
                    cand_logs.append({"Recipient": to_addr, "Zip": os.path.basename(zpath), "Status": "Failed", "Error": str(e)})
                time.sleep(float(delay_seconds))
            else:
                tmpdir = tempfile.mkdtemp(prefix="candparts_")
                parts = create_chunked_zips(pdfs, out_dir=tmpdir, base_name=re.sub(r'[^a-zA-Z0-9]','_',to_addr)[:40], max_bytes=max_bytes)
                for idx,p in enumerate(parts, start=1):
                    subj = f"Your Hall Ticket(s) — part {idx} of {len(parts)}"
                    body = f"Dear Candidate,\n\nPlease find attached part {idx} of your hall ticket(s).\n\n{footer_text}"
                    try:
                        send_email(smtp_cfg, to_addr, subj, body, attachment_path=p["path"], attachment_name=os.path.basename(p["path"]))
                        cand_logs.append({"Recipient": to_addr, "Part": idx, "Zip": os.path.basename(p["path"]), "Status": "Sent"})
                    except Exception as e:
                        cand_logs.append({"Recipient": to_addr, "Part": idx, "Zip": os.path.basename(p["path"]), "Status": "Failed", "Error": str(e)})
                    time.sleep(float(delay_seconds))
        st.subheader("Candidate send logs")
        st.dataframe(pd.DataFrame(cand_logs), use_container_width=True)
        st.success("Candidate sends complete.")
