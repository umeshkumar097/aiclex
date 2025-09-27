# streamlit_app.py
import streamlit as st
import pandas as pd
import zipfile, os, io, tempfile, shutil, time
from collections import defaultdict, OrderedDict
from email.message import EmailMessage
import smtplib
from datetime import datetime
import re 
st.set_page_config(page_title="Aiclex Mailer — Final", layout="wide")
st.title("📧 Aiclex Hallticket Mailer — Final (Location+Emails groups + Parted ZIPs)")

# ---------------- Helpers ----------------
def human_bytes(n):
    try:
        n = float(n)
    except:
        return ""
    for unit in ["B","KB","MB","GB","TB"]:
        if n < 1024:
            return f"{n:.2f} {unit}"
        n /= 1024
    return f"{n:.2f} PB"

def extract_zip_recursively(zip_file_like, extract_to):
    """Extract zip (file-like or path) recursively (nested zips)."""
    with zipfile.ZipFile(zip_file_like) as z:
        z.extractall(path=extract_to)
    # walk to find nested zips and extract them
    for root, _, files in os.walk(extract_to):
        for f in files:
            if f.lower().endswith(".zip"):
                nested = os.path.join(root, f)
                nested_dir = os.path.join(root, f"_nested_{os.path.splitext(f)[0]}")
                os.makedirs(nested_dir, exist_ok=True)
                try:
                    with open(nested, "rb") as nf:
                        extract_zip_recursively(nf, nested_dir)
                except Exception:
                    # skip unreadable nested zip
                    continue

def create_chunked_zips(file_paths, out_dir, base_name, max_bytes):
    """
    Pack file_paths into sequential zip files each <= max_bytes.
    Returns list of zip paths.
    """
    os.makedirs(out_dir, exist_ok=True)
    parts = []
    cur_files = []
    cur_size = 0
    part_index = 1
    for fp in file_paths:
        fsz = os.path.getsize(fp)
        # If single file itself larger than max_bytes, include it alone (can't split PDF)
        if fsz > max_bytes and not cur_files:
            zpath = os.path.join(out_dir, f"{base_name}_part{part_index}.zip")
            with zipfile.ZipFile(zpath, "w", compression=zipfile.ZIP_DEFLATED) as z:
                z.write(fp, arcname=os.path.basename(fp))
            parts.append(zpath)
            part_index += 1
            continue
        # if adding this file exceeds limit, flush current
        if cur_files and (cur_size + fsz) > max_bytes:
            zpath = os.path.join(out_dir, f"{base_name}_part{part_index}.zip")
            with zipfile.ZipFile(zpath, "w", compression=zipfile.ZIP_DEFLATED) as z:
                for f in cur_files:
                    z.write(f, arcname=os.path.basename(f))
            parts.append(zpath)
            part_index += 1
            cur_files = [fp]
            cur_size = fsz
        else:
            cur_files.append(fp)
            cur_size += fsz
    # final flush
    if cur_files:
        zpath = os.path.join(out_dir, f"{base_name}_part{part_index}.zip")
        with zipfile.ZipFile(zpath, "w", compression=zipfile.ZIP_DEFLATED) as z:
            for f in cur_files:
                z.write(f, arcname=os.path.basename(f))
        parts.append(zpath)
    return parts

def send_email_smtp(smtp_host, smtp_port, use_ssl, sender_email, sender_password, recipients, subject, body, attachments, timeout=60):
    msg = EmailMessage()
    msg["From"] = sender_email
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.set_content(body)
    for ap in attachments:
        with open(ap, "rb") as af:
            data = af.read()
        msg.add_attachment(data, maintype="application", subtype="zip", filename=os.path.basename(ap))
    if use_ssl:
        server = smtplib.SMTP_SSL(smtp_host, int(smtp_port), timeout=timeout)
    else:
        server = smtplib.SMTP(smtp_host, int(smtp_port), timeout=timeout)
        server.starttls()
    if sender_password:
        server.login(sender_email, sender_password)
    server.send_message(msg)
    server.quit()

# ---------------- Sidebar / Settings ----------------
with st.sidebar:
    st.header("SMTP & App Settings")
    smtp_host = st.text_input("SMTP host", value="smtp.gmail.com")
    smtp_port = st.number_input("SMTP port", value=465)
    protocol = st.selectbox("Protocol", options=["SMTPS (SSL)", "SMTP (STARTTLS)"], index=0)
    smtp_use_ssl = True if protocol.startswith("SMTPS") else False
    sender_email = st.text_input("Sender email", value="info@cruxmanagement.com")
    sender_password = st.text_input("Sender password (app password)", value="norx wxop hvsm bvfu", type="password")
    st.markdown("---")
    st.subheader("Attachment & sending")
    per_attachment_limit_mb = st.number_input("Per-attachment limit (MB)", value=3.0, step=0.5)
    delay_seconds = st.number_input("Delay between emails (s)", value=2.0, step=0.5)
    st.markdown("---")
    st.write("Use the defaults above or update before sending.")

# ---------------- Uploads ----------------
st.header("1) Upload Excel and ZIP")
uploaded_excel = st.file_uploader("Upload Excel (.xlsx / .csv) with Hallticket, Email, Location columns", type=["xlsx","csv"])
uploaded_zip = st.file_uploader("Upload ZIP (PDFs inside; nested zips supported)", type=["zip"])

if not uploaded_excel or not uploaded_zip:
    st.info("Please upload both Excel and ZIP to continue.")
    st.stop()

# ---------------- Read Excel ----------------
try:
    if uploaded_excel.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_excel, dtype=str).fillna("")
    else:
        df = pd.read_excel(uploaded_excel, dtype=str).fillna("")
except Exception as e:
    st.error("Failed to read Excel: " + str(e))
    st.stop()

if df.empty:
    st.error("Excel appears empty.")
    st.stop()

cols = list(df.columns)
detected_ht = next((c for c in cols if "hall" in c.lower() or "ticket" in c.lower()), cols[0])
detected_email = next((c for c in cols if "email" in c.lower() or "mail" in c.lower()), cols[1] if len(cols)>1 else cols[0])
detected_loc = next((c for c in cols if "loc" in c.lower() or "center" in c.lower() or "city" in c.lower()), cols[2] if len(cols)>2 else cols[0])

ht_col = st.selectbox("Hallticket column", cols, index=cols.index(detected_ht))
email_col = st.selectbox("Email column", cols, index=cols.index(detected_email))
location_col = st.selectbox("Location column", cols, index=cols.index(detected_loc))

st.subheader("Excel preview (first 10 rows)")
st.dataframe(df[[ht_col, email_col, location_col]].head(10), use_container_width=True)

# ---------------- Extract PDFs ----------------
st.header("2) Extract PDFs from uploaded ZIP (nested supported)")
extract_root = tempfile.mkdtemp(prefix="aiclex_extract_")
with st.spinner("Extracting uploaded ZIP (this may take a moment)..."):
    try:
        b = io.BytesIO(uploaded_zip.read())
        extract_zip_recursively(b, extract_root)
    except Exception as e:
        st.error("ZIP extraction failed: " + str(e))
        st.stop()

pdf_paths = []
for root, _, files in os.walk(extract_root):
    for f in files:
        if f.lower().endswith(".pdf"):
            pdf_paths.append(os.path.join(root, f))
st.success(f"Extraction complete — found {len(pdf_paths)} PDF(s)")

if len(pdf_paths) == 0:
    st.error("No PDFs found in uploaded ZIP. Please check the archive.")
    st.stop()

# ---------------- Build mapping: group by (location + frozenset(emails_in_row)) ----------------
st.header("3) Grouping & Matching")

groups = OrderedDict()  # key: (location, frozenset_emails) -> dict with 'halltickets', 'matched_paths'
for idx, row in df.iterrows():
    ht = str(row.get(ht_col, "")).strip()
    loc = str(row.get(location_col, "")).strip()
    raw_emails = str(row.get(email_col, "")).strip()
    # split by comma/semicolon or newline/space
    emails = [e.strip().lower() for e in re.split(r"[,;\n]+", raw_emails) if e.strip()]
    if not emails:
        # treat empty email as a special group (won't be sent)
        emails = ["(no-recipient)"]
    email_key = frozenset(emails)
    key = (loc, email_key)
    if key not in groups:
        groups[key] = {"halltickets": [], "matched_paths": []}
    groups[key]["halltickets"].append(ht)

# quick lookup: pdf filename tokens -> path
pdf_lookup = {}
for p in pdf_paths:
    name = os.path.basename(p)
    # find long numeric token (hallticket) or full filename keys
    pdf_lookup[name] = p
    digits = "".join([c for c in name if c.isdigit() or c=='_'])  # rough
    if digits:
        pdf_lookup[digits] = p

# match halltickets to pdf filenames (simple contains or exact)
for key, info in groups.items():
    matched = []
    for ht in info["halltickets"]:
        found = None
        # exact filename match attempts
        candidates = [f"{ht}.pdf", f"{ht.upper()}.pdf", f"{ht.lower()}.pdf"]
        for c in candidates:
            if c in pdf_lookup:
                found = pdf_lookup[c]
                break
        if not found:
            # contains in filename
            for fname, path in pdf_lookup.items():
                if ht and ht in fname:
                    found = path
                    break
        if found:
            matched.append(found)
    # deduplicate
    groups[key]["matched_paths"] = sorted(list(dict.fromkeys(matched)))

# ---------------- Preview groups & matched counts ----------------
summary_rows = []
for (loc, email_key), info in groups.items():
    total_bytes = sum(os.path.getsize(p) for p in info["matched_paths"]) if info["matched_paths"] else 0
    summary_rows.append({
        "Location": loc or "(empty)",
        "Recipients": ", ".join(sorted(list(email_key))),
        "Tickets in Excel": len(info["halltickets"]),
        "Matched PDFs": len(info["matched_paths"]),
        "TotalSize": human_bytes(total_bytes)
    })

st.subheader("Group Summary (before creating ZIPs)")
st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)

# ---------------- Prepare ZIPs (split to per_attachment_limit) ----------------
st.header("4) Prepare ZIPs (splitting by per-attachment limit)")

if "prepared" not in st.session_state:
    st.session_state.prepared = {}

if st.button("Prepare ZIPs (create parted zips per group)"):
    st.session_state.prepared = {}
    workdir = tempfile.mkdtemp(prefix="aiclex_zips_")
    max_bytes = int(float(per_attachment_limit_mb) * 1024 * 1024)
    for (loc, email_key), info in groups.items():
        matched = info["matched_paths"]
        if not matched:
            st.warning(f"No matched PDFs for group {loc} / {', '.join(sorted(list(email_key)))} — skipping")
            continue
        safe_group_name = re.sub(r"[^A-Za-z0-9_\-]", "_", (loc or "location"))[:60] + "_" + "_".join([re.sub(r"[^A-Za-z0-9]", "_", e) for e in sorted(list(email_key))])[:120]
        out_dir = os.path.join(workdir, safe_group_name)
        os.makedirs(out_dir, exist_ok=True)
        parts = create_chunked_zips(matched, out_dir, safe_group_name, max_bytes)
        parts_info = [{"path": p, "size_bytes": os.path.getsize(p), "size_human": human_bytes(os.path.getsize(p))} for p in parts]
        st.session_state.prepared[(loc, email_key)] = {"parts": parts_info, "out_dir": out_dir, "recipients": sorted(list(email_key))}
    st.success("Prepared ZIP parts for all groups.")

# show prepared summary
if st.session_state.prepared:
    st.subheader("Prepared ZIPs Summary (per group)")
    prep_rows = []
    for (loc, email_key), pdata in st.session_state.prepared.items():
        total_size = sum(p["size_bytes"] for p in pdata["parts"])
        prep_rows.append({
            "Location": loc or "(empty)",
            "Recipients": ", ".join(pdata["recipients"]),
            "Parts": len(pdata["parts"]),
            "TotalSize": human_bytes(total_size)
        })
    st.dataframe(pd.DataFrame(prep_rows), use_container_width=True)

    # list parts and download buttons
    st.markdown("### Parts detail and downloads")
    for (loc, email_key), pdata in st.session_state.prepared.items():
        st.write(f"**Group:** {loc}  —  Recipients: {', '.join(pdata['recipients'])}  —  Parts: {len(pdata['parts'])}")
        cols = st.columns([3,1,1])
        for part in pdata["parts"]:
            c1, c2, c3 = cols
            with c1:
                st.write(os.path.basename(part["path"]))
            with c2:
                st.write(part["size_human"])
            with c3:
                with open(part["path"], "rb") as bf:
                    st.download_button(label="Download", data=bf.read(), file_name=os.path.basename(part["path"]), mime="application/zip")

# ---------------- Send prepared parts ----------------
st.header("5) Send prepared parts (one email per part; recipients = original email set)")

if st.button("Send ALL prepared parts (confirm)"):
    if not st.session_state.prepared:
        st.error("No prepared zips found. Click Prepare ZIPs first.")
    else:
        logs = []
        total_parts = sum(len(v["parts"]) for v in st.session_state.prepared.values())
        sent_count = 0
        progress = st.progress(0)
        for (loc, email_key), pdata in st.session_state.prepared.items():
            recipients = pdata["recipients"]
            parts = pdata["parts"]
            for idx, p in enumerate(parts, start=1):
                subj = f"{st.text_input('subject', value='Hall Tickets — {location}').format(location=loc)} (Part {idx}/{len(parts)})" if False else f"Hall Tickets — {loc} (Part {idx}/{len(parts)})"
                body = f"Dear Coordinator,\n\nPlease find attached the hall tickets for {loc} (Part {idx} of {len(parts)}).\n\nRegards,\nAiclex Technologies"
                try:
                    send_email_smtp(
                        smtp_host=smtp_host,
                        smtp_port=smtp_port,
                        use_ssl=smtp_use_ssl,
                        sender_email=sender_email,
                        sender_password=sender_password,
                        recipients=recipients,
                        subject=subj,
                        body=body,
                        attachments=[p["path"]]
                    )
                    logs.append({"Location": loc, "Recipients": ", ".join(recipients), "Part": idx, "Zip": os.path.basename(p["path"]), "Status": "Sent", "Time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
                except Exception as e:
                    logs.append({"Location": loc, "Recipients": ", ".join(recipients), "Part": idx, "Zip": os.path.basename(p["path"]), "Status": f"Failed: {e}"})
                sent_count += 1
                progress.progress(int(sent_count * 100 / total_parts))
                time.sleep(float(delay_seconds))
        st.subheader("Send logs")
        st.dataframe(pd.DataFrame(logs), use_container_width=True)
        st.success("Send attempts completed. Review logs above.")

# ---------------- Cleanup ----------------
st.header("6) Cleanup temporary files")
if st.button("Cleanup temporary extracted files & prepared zips"):
    removed = {"extraction_root": False, "zips_removed": 0}
    try:
        if os.path.exists(extract_root):
            shutil.rmtree(extract_root, ignore_errors=True)
            removed["extraction_root"] = True
    except Exception:
        removed["extraction_root"] = False
    if st.session_state.get("prepared"):
        try:
            for pdata in st.session_state.prepared.values():
                od = pdata.get("out_dir")
                if od and os.path.exists(od):
                    shutil.rmtree(od, ignore_errors=True)
                    removed["zips_removed"] += 1
            st.session_state.prepared = {}
        except Exception:
            pass
    st.success(f"Cleanup done. extraction_root removed: {removed['extraction_root']}. Prepared zips removed: attempted.")

st.info("Notes: Use the Prepare ZIPs step to see how many parts created per group. Verify downloads before sending final emails. If using Gmail, use an App Password (16-char) for SMTP login.")
