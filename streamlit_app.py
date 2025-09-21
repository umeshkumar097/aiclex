# streamlit_app.py
"""
Aiclex Hallticket Mailer — Final
Features:
- Login guard
- Upload Excel + nested ZIPs
- Filename-last-digit hallticket matching
- Group by Location, create one ZIP per location
- Per-location recipients from Excel (editable)
- Test mode, templates, progress UI, logs, mapping CSV download
"""
import os
import io
import re
import time
import zipfile
import tempfile
import shutil
import smtplib
from email.message import EmailMessage
from collections import defaultdict, Counter
from datetime import datetime

import streamlit as st
import pandas as pd

# ---------------- Config ----------------
st.set_page_config(page_title="Aiclex Mailer (Final)", layout="wide")
st.title("📧 Aiclex Hallticket Mailer — Final")

# ---------------- Simple Login Guard ----------------
DEFAULT_USER = "info@aiclex.in"
DEFAULT_PASS = "Aiclex@2025"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

def login_block():
    if st.session_state.authenticated:
        cols = st.columns([1, 7, 2])
        with cols[2]:
            if st.button("Logout"):
                st.session_state.authenticated = False
                st.rerun()
        return True
    st.markdown("### Login required")
    with st.form("login_form"):
        user = st.text_input("Email", value=DEFAULT_USER)
        pwd = st.text_input("Password", type="password", value=DEFAULT_PASS)
        submitted = st.form_submit_button("Login")
    if submitted:
        if user == DEFAULT_USER and pwd == DEFAULT_PASS:
            st.session_state.authenticated = True
            st.success("Login successful")
            st.rerun()
        else:
            st.error("Invalid credentials")
    return False

if not login_block():
    st.stop()

# ---------------- Helpers ----------------
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

def human_bytes(n):
    try:
        n = float(n)
    except:
        return ""
    for unit in ["B","KB","MB","GB","TB"]:
        if n < 1024:
            return f"{n:.2f} {unit}"
        n /= 1024
    return f"{n:.2f} TB"

def extract_zip_bytes_recursively(zip_bytes, out_root):
    """
    Extract zip bytes recursively (handles nested zips).
    Returns list of absolute PDF paths extracted.
    """
    extracted_pdfs = []
    # unique folder for this extraction
    base_dir = tempfile.mkdtemp(prefix="aiclex_unzip_", dir=out_root)

    def _process_zip(zf, curdir):
        for info in zf.infolist():
            if info.is_dir():
                continue
            name = info.filename
            lname = name.lower()
            try:
                data = zf.read(info)
            except Exception:
                continue
            if lname.endswith(".zip"):
                # nested zip: create nested dir and recurse
                nested_dir = os.path.join(curdir, os.path.splitext(os.path.basename(name))[0])
                os.makedirs(nested_dir, exist_ok=True)
                try:
                    with zipfile.ZipFile(io.BytesIO(data)) as nz:
                        _process_zip(nz, nested_dir)
                except Exception:
                    # fallback: write to disk then open
                    try:
                        tmpf = os.path.join(nested_dir, os.path.basename(name))
                        with open(tmpf, "wb") as wf:
                            wf.write(data)
                        with zipfile.ZipFile(tmpf) as nz:
                            _process_zip(nz, os.path.splitext(tmpf)[0])
                    except Exception:
                        continue
            elif lname.endswith(".pdf"):
                # write pdf out
                os.makedirs(curdir, exist_ok=True)
                target = os.path.join(curdir, os.path.basename(name))
                # avoid overwrite collisions
                if os.path.exists(target):
                    base, ext = os.path.splitext(os.path.basename(name))
                    target = os.path.join(curdir, f"{base}_{int(time.time()*1000)}{ext}")
                try:
                    with open(target, "wb") as outf:
                        outf.write(data)
                    extracted_pdfs.append(os.path.abspath(target))
                except Exception:
                    continue
            else:
                # ignore other file types
                continue

    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
        _process_zip(z, base_dir)

    return extracted_pdfs

def extract_hallticket_from_filename(path):
    """
    Extract last digit sequence from filename.
    e.g. 1036_17_802871022.pdf -> 802871022
    """
    base = os.path.splitext(os.path.basename(path))[0]
    digits = re.findall(r"\d+", base)
    return digits[-1] if digits else None

def create_zip_single(files, out_dir, base_name):
    os.makedirs(out_dir, exist_ok=True)
    safe = re.sub(r"[^A-Za-z0-9_\-]+", "_", base_name)[:100]
    zpath = os.path.join(out_dir, f"{safe}.zip")
    with zipfile.ZipFile(zpath, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(f, arcname=os.path.basename(f))
    return zpath

def send_email_smtp(smtp_cfg, to_addr, subject, body, attachment_paths):
    msg = EmailMessage()
    msg["From"] = smtp_cfg["sender"]
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.set_content(body)
    for ap in attachment_paths:
        with open(ap, "rb") as af:
            data = af.read()
        msg.add_attachment(data, maintype="application", subtype="zip", filename=os.path.basename(ap))
    if smtp_cfg.get("use_ssl", True):
        server = smtplib.SMTP_SSL(smtp_cfg["host"], smtp_cfg["port"], timeout=60)
    else:
        server = smtplib.SMTP(smtp_cfg["host"], smtp_cfg["port"], timeout=60)
        server.starttls()
    if smtp_cfg.get("password"):
        server.login(smtp_cfg["sender"], smtp_cfg["password"])
    server.send_message(msg)
    server.quit()

# ---------------- Sidebar / Settings ----------------
with st.sidebar:
    st.header("SMTP & Templates")
    smtp_host = st.text_input("SMTP host", value="smtp.hostinger.com")
    smtp_port = st.number_input("SMTP port", value=465)
    smtp_use_ssl = st.checkbox("Use SSL (SMTPS)", value=True)
    smtp_sender = st.text_input("Sender email", value="info@aiclex.in")
    smtp_password = st.text_input("Sender password", type="password", value="")

    st.markdown("---")
    st.subheader("Email content")
    subject_template = st.text_input("Subject (use {location})", value="Hall Tickets — {location}")
    body_template = st.text_area("Body (use {location} and {footer})", value="Dear Coordinator,\n\nPlease find attached the hall tickets for {location}.\n\n{footer}", height=140)
    footer_text = st.text_input("Footer", value="Regards,\nAiclex Technologies\ninfo@aiclex.in")

    st.markdown("---")
    st.subheader("Options")
    delay_seconds = st.number_input("Delay between sends (seconds)", value=2.0, step=0.5)
    test_mode = st.checkbox("Enable Test Mode (send all mails to Test Email)", value=True)
    test_email = st.text_input("Test Email (used when Test Mode ON)", value="info@aiclex.in")

# ---------------- File Upload ----------------
st.header("1) Upload Excel & Master ZIP")
uploaded_excel = st.file_uploader("Upload Excel (.xlsx or .csv) with Hallticket, Recipient Email, Location columns", type=["xlsx","csv"])
uploaded_zip = st.file_uploader("Upload Master ZIP (PDFs; nested zips allowed)", type=["zip"])
if not uploaded_excel or not uploaded_zip:
    st.info("Upload both Excel and ZIP to continue.")
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

# Detect columns heuristically
cols = list(df.columns)
if not cols:
    st.error("Excel contains no columns.")
    st.stop()

detected_ht_col = next((c for c in cols if "hall" in c.lower() or "ticket" in c.lower()), cols[0])
detected_email_col = next((c for c in cols if "email" in c.lower() or "mail" in c.lower()), cols[1] if len(cols) > 1 else cols[0])
detected_loc_col = next((c for c in cols if "loc" in c.lower() or "center" in c.lower() or "city" in c.lower()), cols[2] if len(cols) > 2 else cols[0])

ht_col = st.selectbox("Hallticket column", cols, index=cols.index(detected_ht_col))
email_col = st.selectbox("Recipient Email column", cols, index=cols.index(detected_email_col))
location_col = st.selectbox("Location column", cols, index=cols.index(detected_loc_col))

st.subheader("Excel preview (first 10 rows)")
st.dataframe(df[[ht_col, email_col, location_col]].head(10), use_container_width=True)

# ---------------- Extract PDFs (nested) ----------------
st.header("2) Extract PDFs from ZIP (nested supported)")
extraction_root = tempfile.mkdtemp(prefix="aiclex_all_unzips_")
with st.spinner("Extracting ZIP (this may take a moment for nested zips)..."):
    try:
        uploaded_zip_bytes = uploaded_zip.read()
        extracted_pdf_paths = extract_zip_bytes_recursively(uploaded_zip_bytes, extraction_root)
    except Exception as e:
        st.error("ZIP extraction error: " + str(e))
        extracted_pdf_paths = []

st.success(f"Extraction finished — found {len(extracted_pdf_paths)} PDFs.")
if len(extracted_pdf_paths) == 0:
    st.warning("No PDFs found inside the uploaded ZIP. Check the archive structure.")
    st.stop()

# Show sample of extracted files
st.subheader("Extracted PDFs (sample)")
sample_list = [{"basename": os.path.basename(p), "path": p, "size": human_bytes(os.path.getsize(p))} for p in extracted_pdf_paths[:200]]
st.dataframe(pd.DataFrame(sample_list), use_container_width=True)

# ---------------- Build pdf_map by hallticket extracted from filename ----------------
pdf_map = {}  # hallticket_str -> pdf_absolute_path (last wins)
for p in extracted_pdf_paths:
    ht = extract_hallticket_from_filename(p)
    if ht:
        pdf_map[ht] = p

# ---------------- Matching Excel rows to PDFs ----------------
st.header("3) Match Halltickets to PDFs & Group by Location")
mapping_rows = []
grouped_by_location = defaultdict(list)  # location -> list of (pdf_path, recipient_email)

for idx, row in df.iterrows():
    ht_val = str(row.get(ht_col, "")).strip()
    recipient = str(row.get(email_col, "")).strip()
    location = str(row.get(location_col, "")).strip()
    matched_path = pdf_map.get(ht_val)
    mapping_rows.append({
        "Sr No": idx+1,
        "Hallticket": ht_val,
        "Recipient": recipient,
        "Location": location,
        "Matched": "Yes" if matched_path else "No",
        "MatchedFile": os.path.basename(matched_path) if matched_path else ""
    })
    if matched_path:
        grouped_by_location[location].append((matched_path, recipient))

map_df = pd.DataFrame(mapping_rows)
st.subheader("Mapping Preview (first 200 rows)")
st.dataframe(map_df.head(200), use_container_width=True)

# Allow download mapping CSV
csv_buf = io.StringIO()
map_df.to_csv(csv_buf, index=False)
st.download_button("Download mapping_check.csv", data=csv_buf.getvalue(), file_name="mapping_check.csv", mime="text/csv")

# ---------------- Auto-fill & Edit Recipients per Location ----------------
st.header("4) Recipients per Location (auto-filled — edit if required)")
if "location_recipients" not in st.session_state:
    st.session_state.location_recipients = {}

all_locations = sorted(list(set(str(r).strip() for r in df[location_col].astype(str).unique())))

for loc in all_locations:
    # auto-suggest: top 3 emails from Excel rows for this location
    if not st.session_state.location_recipients.get(loc):
        rows = df[df[location_col].astype(str).str.strip() == loc]
        extracted = []
        for v in rows[email_col].astype(str).tolist():
            extracted += EMAIL_RE.findall(v)
        # unique preserve order
        seen = []
        for e in extracted:
            if e not in seen:
                seen.append(e)
        st.session_state.location_recipients[loc] = ";".join(seen[:3])
    # editable text area for each location
    st.session_state.location_recipients[loc] = st.text_area(f"Recipients for: {loc}", value=st.session_state.location_recipients[loc], key=f"recip_{loc}", height=70)

# Validate recipients
invalid_locs = []
for loc in all_locations:
    raw = st.session_state.location_recipients.get(loc, "")
    if raw:
        parts = [x.strip() for x in re.split(r"[;,\n]+", raw) if x.strip()]
        valid = [p for p in parts if EMAIL_RE.search(p)]
        if raw and not valid:
            invalid_locs.append(loc)
if invalid_locs:
    st.warning("These locations have recipient text but no valid emails: " + ", ".join(invalid_locs))

# ---------------- Prepare & Send (with Test Mode, progress, spinners, logs) ----------------
st.header("5) Prepare ZIP(s) & Send Emails")

attachment_limit_mb = st.number_input("Per-attachment size limit (MB) — use 3 for clients with 3MB limit", value=3.0, step=0.5)
attachment_limit_bytes = int(float(attachment_limit_mb) * 1024 * 1024)

def validate_all_recipients():
    errs = []
    for loc in all_locations:
        raw = st.session_state.location_recipients.get(loc, "")
        if raw:
            parts = [x.strip() for x in re.split(r"[;,\n]+", raw) if x.strip()]
            val = [p for p in parts if EMAIL_RE.search(p)]
            if not val:
                errs.append(loc)
    return errs

if st.button("Prepare & Send All (use Test Mode recommended)"):
    bad = validate_all_recipients()
    if bad:
        st.error("Fix recipient entries for: " + ", ".join(bad))
        st.stop()

    # SMTP quick test (fail early)
    smtp_cfg = {
        "host": smtp_host,
        "port": int(smtp_port),
        "use_ssl": bool(smtp_use_ssl),
        "sender": smtp_sender,
        "password": smtp_password
    }
    try:
        if smtp_cfg["use_ssl"]:
            t = smtplib.SMTP_SSL(smtp_cfg["host"], smtp_cfg["port"], timeout=30)
        else:
            t = smtplib.SMTP(smtp_cfg["host"], smtp_cfg["port"], timeout=30)
            t.starttls()
        if smtp_cfg.get("password"):
            t.login(smtp_cfg["sender"], smtp_cfg["password"])
        t.quit()
        st.success("SMTP connection / login successful (checked).")
    except Exception as e:
        st.error("SMTP test failed: " + str(e))
        st.stop()

    # Build list of (location, unique_files, recipients_list)
    tasks = []
    for loc, items in grouped_by_location.items():
        files = list(dict.fromkeys([p for p, _ in items]))  # unique file paths
        if not files:
            continue
        # recipients: from editable per-location field
        raw = st.session_state.location_recipients.get(loc, "")
        recips = [r.strip() for r in re.split(r"[;,\n]+", raw) if r.strip() and EMAIL_RE.search(r)]
        if not recips:
            continue
        tasks.append((loc, files, recips))

    if not tasks:
        st.warning("No tasks to send (no matched files or no recipients).")
        st.stop()

    total_recipients = sum(len(t[2]) for t in tasks)
    progress_bar = st.progress(0)
    job_count = 0
    logs = []

    st.info(f"Starting send: {len(tasks)} locations, {total_recipients} recipient-addresses (test_mode={test_mode}).")

    for loc, files, recips in tasks:
        # create unique outdir
        outdir = tempfile.mkdtemp(prefix="send_")
        # chunk if over limit — we'll make single zip ensuring size <= limit by naive packing; for simplicity create single zip and check size
        zip_path = create_zip_single(files, outdir, base_name=loc or "location")
        zipped_size = os.path.getsize(zip_path)
        if zipped_size > attachment_limit_bytes:
            # if zip bigger than limit, try splitting by writing per-file zips until under limit
            # simple greedy: create multiple zips with approx equal distribution
            parts = []
            cur = []
            cur_size = 0
            for f in files:
                fsz = os.path.getsize(f)
                if cur and (cur_size + fsz) > attachment_limit_bytes:
                    # flush cur to zip
                    pth = create_zip_single(cur, outdir, base_name=f"{loc}_part{len(parts)+1}")
                    parts.append(pth)
                    cur = [f]; cur_size = fsz
                else:
                    cur.append(f); cur_size += fsz
            if cur:
                pth = create_zip_single(cur, outdir, base_name=f"{loc}_part{len(parts)+1}")
                parts.append(pth)
        else:
            parts = [zip_path]

        # send to each recipient
        for rec in recips:
            send_to = test_email if test_mode and test_email else rec
            for idx, part in enumerate(parts, start=1):
                subject = subject_template.format(location=loc)
                body = body_template.format(location=loc, footer=footer_text)
                with st.spinner(f"Sending {os.path.basename(part)} → {send_to} (Location: {loc})"):
                    try:
                        send_email_smtp(smtp_cfg, send_to, subject, body, [part])
                        logs.append({
                            "Location": loc,
                            "Recipient (Excel)": rec,
                            "Sent To": send_to,
                            "Zip": os.path.basename(part),
                            "Status": "Sent",
                            "Subject": subject,
                            "Time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        })
                        st.success(f"Sent: {os.path.basename(part)} → {send_to} (Location: {loc})")
                    except Exception as e:
                        logs.append({
                            "Location": loc,
                            "Recipient (Excel)": rec,
                            "Sent To": send_to,
                            "Zip": os.path.basename(part),
                            "Status": "Failed",
                            "Error": str(e)
                        })
                        st.error(f"Failed: {os.path.basename(part)} → {send_to} (Location: {loc}) — {e}")
                job_count += 1
                # update progress bar in percent (safe handling zero)
                if total_recipients > 0:
                    progress_bar.progress(min(100, int(job_count * 100 / (total_recipients))))
                time.sleep(float(delay_seconds))

        # cleanup per-location outdir if you want (keep for a short time)
        try:
            shutil.rmtree(outdir)
        except Exception:
            pass

    st.subheader("Send logs")
    st.dataframe(pd.DataFrame(logs), use_container_width=True)
    st.success("All sending attempts complete.")

# ---------------- Cleanup temp extraction (button) ----------------
if st.button("Cleanup temporary extracted files"):
    try:
        shutil.rmtree(extraction_root)
        st.success("Temporary extraction folder removed.")
    except Exception as e:
        st.error("Cleanup failed: " + str(e))

st.info("Notes: For scanned PDFs (images), add OCR (pdf2image + pytesseract + poppler). For large volume sending consider S3 + signed links and SendGrid/SES.")
