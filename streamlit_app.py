# streamlit_app.py
"""
Aiclex Hallticket Mailer — Final (single-file)
Features:
- Login guard (default info@aiclex.in / Aiclex@2025)
- Upload Excel (.xlsx/.csv) + Master ZIP (supports single-level ZIP or nested ZIPs)
- Extract PDFs recursively and avoid collisions
- Match PDFs to hallticket numbers by extracting last digit-sequence from filename
- Group matched PDFs by Location from Excel
- Auto-suggest & editable recipients per location
- Create per-location ZIPs, split if needed to respect attachment size limit
- Test Mode (send all mails to test email) and detailed logs
- Progress UI: spinners, progress bar, live logs
- Mapping CSV download and cleanup button
Run with:
    streamlit run streamlit_app.py --server.port=8501 --server.address=0.0.0.0
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
from collections import defaultdict
from datetime import datetime

import streamlit as st
import pandas as pd

# ---------------- Streamlit config ----------------
st.set_page_config(page_title="Aiclex Mailer — Final", layout="wide")
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
            st.success("Login successful — loading app...")
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
    Robust recursive extractor:
    - Accepts zip bytes (top-level)
    - Extracts PDFs found in the archive and any nested zips
    - Writes files to unique subfolders under out_root to avoid collisions
    - Returns list of absolute PDF paths
    Works for:
      - single-level ZIP (PDFs directly inside)
      - nested ZIP(s) (ZIP inside ZIP; arbitrary nesting)
    """
    extracted_pdfs = []
    base_dir = tempfile.mkdtemp(prefix="aiclex_unzip_", dir=out_root)

    def _process_zip_bytes(zip_data, curdir):
        try:
            with zipfile.ZipFile(io.BytesIO(zip_data)) as zf:
                for info in zf.infolist():
                    if info.is_dir():
                        continue
                    name = info.filename
                    lname = name.lower()
                    try:
                        data = zf.read(info)
                    except Exception:
                        # unreadable entry: skip
                        continue
                    if lname.endswith(".zip"):
                        # nested ZIP -> recurse
                        nested_dir = os.path.join(curdir, os.path.splitext(os.path.basename(name))[0])
                        os.makedirs(nested_dir, exist_ok=True)
                        _process_zip_bytes(data, nested_dir)
                    elif lname.endswith(".pdf"):
                        os.makedirs(curdir, exist_ok=True)
                        target = os.path.join(curdir, os.path.basename(name))
                        # avoid overwrite collisions
                        if os.path.exists(target):
                            base, ext = os.path.splitext(os.path.basename(name))
                            ts = int(time.time() * 1000)
                            target = os.path.join(curdir, f"{base}_{ts}{ext}")
                        try:
                            with open(target, "wb") as outf:
                                outf.write(data)
                            extracted_pdfs.append(os.path.abspath(target))
                        except Exception:
                            continue
                    else:
                        # ignore other file types
                        continue
        except zipfile.BadZipFile:
            # corrupted zip bytes
            return
        except Exception:
            return

    # start
    _process_zip_bytes(zip_bytes, base_dir)
    return extracted_pdfs

def extract_hallticket_from_filename(path):
    """
    Extract last digit sequence from filename.
    Example: 1036_17_802871022.pdf -> 802871022
    """
    base = os.path.splitext(os.path.basename(path))[0]
    digits = re.findall(r"\d+", base)
    return digits[-1] if digits else None

def create_zip_single(files, out_dir, base_name):
    os.makedirs(out_dir, exist_ok=True)
    safe = re.sub(r"[^A-Za-z0-9_\-]+", "_", str(base_name))[:120]
    zpath = os.path.join(out_dir, f"{safe}.zip")
    with zipfile.ZipFile(zpath, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for f in files:
            # skip missing files
            if os.path.exists(f):
                z.write(f, arcname=os.path.basename(f))
    return zpath

def send_email_smtp(cfg, to_addr, subject, body, attachment_paths):
    msg = EmailMessage()
    msg["From"] = cfg["sender"]
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.set_content(body)
    for ap in attachment_paths:
        with open(ap, "rb") as af:
            data = af.read()
        msg.add_attachment(data, maintype="application", subtype="zip", filename=os.path.basename(ap))
    if cfg.get("use_ssl", True):
        server = smtplib.SMTP_SSL(cfg["host"], cfg["port"], timeout=60)
    else:
        server = smtplib.SMTP(cfg["host"], cfg["port"], timeout=60)
        server.starttls()
    if cfg.get("password"):
        server.login(cfg["sender"], cfg["password"])
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
    st.subheader("Email templates")
    subject_template = st.text_input("Subject (use {location})", value="Hall Tickets — {location}")
    body_template = st.text_area("Body (use {location} & {footer})",
                                value="Dear Coordinator,\n\nPlease find attached the hall tickets for {location}.\n\n{footer}",
                                height=140)
    footer_text = st.text_input("Footer", value="Regards,\nAiclex Technologies\ninfo@aiclex.in")

    st.markdown("---")
    st.subheader("Options")
    delay_seconds = st.number_input("Delay between sends (seconds)", value=2.0, step=0.5)
    attachment_limit_mb = st.number_input("Per-attachment limit (MB)", value=3.0, step=0.5)
    test_mode = st.checkbox("Enable Test Mode (send all to Test Email)", value=True)
    test_email = st.text_input("Test Email (used when Test Mode ON)", value="info@aiclex.in")

# ---------------- Upload area ----------------
st.header("1) Upload Excel & Master ZIP")
uploaded_excel = st.file_uploader("Upload Excel (.xlsx/.csv) with Hallticket, Recipient Email, Location columns", type=["xlsx","csv"])
uploaded_zip = st.file_uploader("Upload Master ZIP (PDFs or nested zips)", type=["zip"])

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

# detect columns
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

# ---------------- Extract PDFs ----------------
st.header("2) Extract PDFs from uploaded ZIP")
extraction_root = tempfile.mkdtemp(prefix="aiclex_all_unzips_")
with st.spinner("Extracting uploaded ZIP (supports nested zips)..."):
    try:
        uploaded_bytes = uploaded_zip.read()
        extracted_pdf_paths = extract_zip_bytes_recursively(uploaded_bytes, extraction_root)
    except Exception as e:
        st.error("ZIP extraction failed: " + str(e))
        extracted_pdf_paths = []

st.success(f"Extraction complete — found {len(extracted_pdf_paths)} PDF(s).")
if not extracted_pdf_paths:
    st.warning("No PDFs found in the uploaded ZIP. Check structure and try again.")
    st.stop()

# sample view
sample_list = [{"basename": os.path.basename(p), "path": p, "size": human_bytes(os.path.getsize(p))} for p in extracted_pdf_paths[:200]]
st.subheader("Extracted PDFs (sample)")
st.dataframe(pd.DataFrame(sample_list), use_container_width=True)

# ---------------- Build pdf_map by hallticket extracted from filename ----------------
pdf_map = {}  # hallticket_str -> pdf_path (last wins)
for p in extracted_pdf_paths:
    ht = extract_hallticket_from_filename(p)
    if ht:
        pdf_map[ht] = p

# ---------------- Mapping Excel rows to PDFs, grouping by location ----------------
st.header("3) Match Halltickets -> PDFs & Group by Location")
mapping_rows = []
grouped_by_location = defaultdict(list)

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
st.subheader("Mapping preview (first 200 rows)")
st.dataframe(map_df.head(200), use_container_width=True)

# allow mapping CSV download
csv_buf = io.StringIO()
map_df.to_csv(csv_buf, index=False)
st.download_button("Download mapping_check.csv", data=csv_buf.getvalue(), file_name="mapping_check.csv", mime="text/csv")

# ---------------- Auto-fill & edit recipients per location ----------------
st.header("4) Recipients per Location (auto-filled — edit if required)")
if "location_recipients" not in st.session_state:
    st.session_state.location_recipients = {}

all_locations = sorted(list(set(str(r).strip() for r in df[location_col].astype(str).unique())))

for loc in all_locations:
    if not st.session_state.location_recipients.get(loc):
        rows = df[df[location_col].astype(str).str.strip() == loc]
        extracted = []
        for v in rows[email_col].astype(str).tolist():
            extracted += EMAIL_RE.findall(v)
        seen = []
        for e in extracted:
            if e not in seen:
                seen.append(e)
        st.session_state.location_recipients[loc] = ";".join(seen[:3])
    st.session_state.location_recipients[loc] = st.text_area(f"Recipients for: {loc}", value=st.session_state.location_recipients[loc], key=f"recip_{loc}", height=70)

# validate recipients quickly
invalid_locs = []
for loc in all_locations:
    raw = st.session_state.location_recipients.get(loc, "")
    if raw:
        parts = [x.strip() for x in re.split(r"[;,\n]+", raw) if x.strip()]
        val = [p for p in parts if EMAIL_RE.search(p)]
        if raw and not val:
            invalid_locs.append(loc)
if invalid_locs:
    st.warning("These locations have recipient text but no valid emails: " + ", ".join(invalid_locs))

# ---------------- Prepare & Send (progress, spinners, test mode) ----------------
st.header("5) Prepare ZIP(s) & Send Emails")
attachment_limit_bytes = int(float(attachment_limit_mb) * 1024 * 1024)

def validate_all_recipients():
    errs = []
    for loc in all_locations:
        raw = st.session_state.location_recipients.get(loc, "")
        if raw:
            parts = [x.strip() for x in re.split(r"[;,\n]+", raw) if x.strip()]
            valid = [p for p in parts if EMAIL_RE.search(p)]
            if not valid:
                errs.append(loc)
    return errs

if st.button("Prepare & Send All (use Test Mode recommended)"):
    bad = validate_all_recipients()
    if bad:
        st.error("Fix recipient entries for: " + ", ".join(bad))
        st.stop()

    # SMTP quick test
    smtp_cfg = {"host": smtp_host, "port": int(smtp_port), "use_ssl": bool(smtp_use_ssl), "sender": smtp_sender, "password": smtp_password}
    try:
        if smtp_cfg["use_ssl"]:
            t = smtplib.SMTP_SSL(smtp_cfg["host"], smtp_cfg["port"], timeout=30)
        else:
            t = smtplib.SMTP(smtp_cfg["host"], smtp_cfg["port"], timeout=30)
            t.starttls()
        if smtp_cfg.get("password"):
            t.login(smtp_cfg["sender"], smtp_cfg["password"])
        t.quit()
        st.success("SMTP login successful.")
    except Exception as e:
        st.error("SMTP connection/login failed: " + str(e))
        st.stop()

    # prepare tasks
    tasks = []
    for loc, items in grouped_by_location.items():
        files = list(dict.fromkeys([p for p, _ in items]))
        if not files:
            continue
        raw = st.session_state.location_recipients.get(loc, "")
        recips = [r.strip() for r in re.split(r"[;,\n]+", raw) if r.strip() and EMAIL_RE.search(r)]
        if not recips:
            continue
        tasks.append((loc, files, recips))

    if not tasks:
        st.warning("No tasks to send (no matched files or no recipients).")
        st.stop()

    total_sends = sum(len(t[2]) for t in tasks)
    progress_bar = st.progress(0)
    sent_count = 0
    logs = []

    st.info(f"Starting sending: {len(tasks)} locations, {total_sends} recipient-addresses (test_mode={test_mode}).")

    for loc, files, recips in tasks:
        # create outdir per location
        outdir = tempfile.mkdtemp(prefix="send_")
        # create a single zip and check size
        single_zip = create_zip_single(files, outdir, base_name=loc or "location")
        if os.path.exists(single_zip) and os.path.getsize(single_zip) <= attachment_limit_bytes:
            parts = [single_zip]
        else:
            # split into multiple zips greedily
            parts = []
            cur = []
            cur_size = 0
            for f in files:
                fsz = os.path.getsize(f)
                if cur and (cur_size + fsz) > attachment_limit_bytes:
                    pth = create_zip_single(cur, outdir, base_name=f"{loc}_part{len(parts)+1}")
                    parts.append(pth)
                    cur = [f]; cur_size = fsz
                else:
                    cur.append(f); cur_size += fsz
            if cur:
                pth = create_zip_single(cur, outdir, base_name=f"{loc}_part{len(parts)+1}")
                parts.append(pth)

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
                sent_count += 1
                if total_sends > 0:
                    progress_bar.progress(min(100, int(sent_count * 100 / total_sends)))
                time.sleep(float(delay_seconds))
        # cleanup outdir
        try:
            shutil.rmtree(outdir)
        except Exception:
            pass

    st.subheader("Send logs")
    st.dataframe(pd.DataFrame(logs), use_container_width=True)
    st.success("All sending attempts complete.")

# ---------------- Cleanup extraction root ----------------
if st.button("Cleanup temporary extracted files"):
    try:
        shutil.rmtree(extraction_root)
        st.success("Temporary extraction folder removed.")
    except Exception as e:
        st.error("Cleanup failed: " + str(e))

st.info("Notes: If PDFs are scanned images (no searchable text) and you later want OCR, install poppler + tesseract and we can add pdf2image+pytesseract fallback. For high-volume emailing use SendGrid/SES and S3 signed links.")
