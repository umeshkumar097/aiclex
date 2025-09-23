# streamlit_app.py
"""
Aiclex Hallticket Mailer — Final (with 3MB part splitting, per-location zips, aggregated recipients)
"""
import os, io, re, time, zipfile, tempfile, shutil, smtplib
from email.message import EmailMessage
from collections import defaultdict
from datetime import datetime
import streamlit as st
import pandas as pd

# ---------------- Streamlit config ----------------
st.set_page_config(page_title="Aiclex Mailer — Final", layout="wide")
st.title("📩 Aiclex Technologies — Final Hall Ticket Mailer")

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

def human_bytes(n):
    try: n = float(n)
    except: return ""
    for unit in ["B","KB","MB","GB"]:
        if n < 1024: return f"{n:.2f} {unit}"
        n /= 1024
    return f"{n:.2f} TB"

def extract_zip_bytes_recursively(zip_bytes, out_root):
    """Extract PDFs from nested ZIPs"""
    extracted_pdfs = []
    extraction_root = tempfile.mkdtemp(prefix="aiclex_unzip_", dir=out_root)
    def _process_zip(data_bytes, curdir):
        try:
            with zipfile.ZipFile(io.BytesIO(data_bytes)) as zf:
                for info in zf.infolist():
                    if info.is_dir(): continue
                    lname = info.filename.lower()
                    entry_bytes = zf.read(info)
                    if lname.endswith(".zip"):
                        nested_dir = os.path.join(curdir, os.path.splitext(os.path.basename(info.filename))[0])
                        os.makedirs(nested_dir, exist_ok=True)
                        _process_zip(entry_bytes, nested_dir)
                    elif lname.endswith(".pdf"):
                        os.makedirs(curdir, exist_ok=True)
                        target = os.path.join(curdir, os.path.basename(info.filename))
                        if os.path.exists(target):
                            base, ext = os.path.splitext(os.path.basename(info.filename))
                            target = os.path.join(curdir, f"{base}_{int(time.time()*1000)}{ext}")
                        with open(target, "wb") as wf:
                            wf.write(entry_bytes)
                        extracted_pdfs.append(os.path.abspath(target))
        except Exception: return
    _process_zip(zip_bytes, extraction_root)
    return extracted_pdfs, extraction_root

def extract_hallticket_from_filename(path):
    base = os.path.splitext(os.path.basename(path))[0]
    digits = re.findall(r"\d+", base)
    return digits[-1] if digits else None

def create_split_zips(files, out_dir, base_name, max_bytes):
    """Split files into multiple zips if total > max_bytes"""
    os.makedirs(out_dir, exist_ok=True)
    zips = []
    cur, cur_size, part = [], 0, 1
    for f in files:
        size = os.path.getsize(f)
        if cur and (cur_size + size) > max_bytes:
            zpath = os.path.join(out_dir, f"{base_name}_part{part}.zip")
            with zipfile.ZipFile(zpath, "w", compression=zipfile.ZIP_DEFLATED) as z:
                for x in cur: z.write(x, arcname=os.path.basename(x))
            zips.append(zpath)
            cur, cur_size, part = [f], size, part+1
        else:
            cur.append(f); cur_size += size
    if cur:
        zpath = os.path.join(out_dir, f"{base_name}_part{part}.zip")
        with zipfile.ZipFile(zpath, "w", compression=zipfile.ZIP_DEFLATED) as z:
            for x in cur: z.write(x, arcname=os.path.basename(x))
        zips.append(zpath)
    return zips

def send_email_smtp(cfg, recipients, subject, body, attachments):
    msg = EmailMessage()
    msg["From"] = cfg["sender"]
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.set_content(body)
    for ap in attachments:
        with open(ap, "rb") as af: data = af.read()
        msg.add_attachment(data, maintype="application", subtype="zip", filename=os.path.basename(ap))
    if cfg.get("use_ssl", True):
        server = smtplib.SMTP_SSL(cfg["host"], cfg["port"], timeout=60)
    else:
        server = smtplib.SMTP(cfg["host"], cfg["port"], timeout=60)
        server.starttls()
    if cfg.get("password"): server.login(cfg["sender"], cfg["password"])
    server.send_message(msg); server.quit()

# ---------------- Sidebar ----------------
with st.sidebar:
    st.header("SMTP Settings")
    smtp_host = st.text_input("SMTP host", "smtp.hostinger.com")
    smtp_port = st.number_input("SMTP port", value=465)
    smtp_use_ssl = st.checkbox("Use SSL (SMTPS)", value=True)
    smtp_sender = st.text_input("Sender email", "info@aiclex.in")
    smtp_password = st.text_input("Sender password", type="password")
    st.markdown("---")
    subject_template = st.text_input("Subject (use {location}, {part}, {total})", "Hall Tickets — {location} (Part {part}/{total})")
    body_template = st.text_area("Body (use {location}, {part}, {total}, {footer})", 
        "Dear Coordinator,\n\nAttached are the hall tickets for {location} (Part {part} of {total}).\n\n{footer}")
    footer_text = st.text_input("Footer text", "Regards,\nAiclex Technologies")
    delay_seconds = st.number_input("Delay between sends (s)", value=2.0, step=0.5)
    attachment_limit_mb = st.number_input("Attachment limit (MB)", value=3.0, step=0.5)
    test_mode = st.checkbox("Test Mode (redirect to test email)", value=True)
    test_email = st.text_input("Test Email", "info@aiclex.in")

# ---------------- Upload ----------------
st.header("1) Upload Excel & ZIP")
uploaded_excel = st.file_uploader("Upload Excel", type=["xlsx","csv"])
uploaded_zip = st.file_uploader("Upload ZIP", type=["zip"])

if uploaded_excel and uploaded_zip:
    # Load Excel
    if uploaded_excel.name.endswith(".csv"):
        df = pd.read_csv(uploaded_excel, dtype=str).fillna("")
    else:
        df = pd.read_excel(uploaded_excel, dtype=str).fillna("")
    cols = df.columns.tolist()
    ht_col = st.selectbox("Hallticket column", cols)
    email_col = st.selectbox("Email column", cols)
    loc_col = st.selectbox("Location column", cols)

    # Extract PDFs
    pdfs, root = extract_zip_bytes_recursively(uploaded_zip.read(), "/tmp")
    st.success(f"Extracted {len(pdfs)} PDFs")
    pdf_lookup = {extract_hallticket_from_filename(p): p for p in pdfs if extract_hallticket_from_filename(p)}

    # Group by location
    grouped = defaultdict(lambda: {"files": [], "recipients": set()})
    for _, r in df.iterrows():
        ht = str(r.get(ht_col,"")).strip()
        loc = str(r.get(loc_col,"")).strip()
        raw_emails = str(r.get(email_col,"")).strip()
        if raw_emails:
            for p in re.split(r"[;, \n]+", raw_emails):
                if p.strip() and EMAIL_RE.match(p.strip()):
                    grouped[loc]["recipients"].add(p.strip())
        if ht in pdf_lookup:
            grouped[loc]["files"].append(pdf_lookup[ht])

    st.subheader("Summary")
    rows = []
    for loc, info in grouped.items():
        rows.append({"Location": loc, "Recipients": ", ".join(info["recipients"]),
                     "Files": len(info["files"]),
                     "Total Size": human_bytes(sum(os.path.getsize(f) for f in info["files"]))})
    st.dataframe(pd.DataFrame(rows))

    # Prepare ZIPs with splitting
    if st.button("Prepare ZIPs"):
        zip_dir = tempfile.mkdtemp(prefix="aiclex_zips_", dir="/tmp")
        max_bytes = int(attachment_limit_mb * 1024 * 1024)
        prepared = {}
        for loc, info in grouped.items():
            if info["files"]:
                safe_loc = re.sub(r"[^A-Za-z0-9]+", "_", loc)[:50]
                zips = create_split_zips(info["files"], zip_dir, safe_loc, max_bytes)
                prepared[loc] = zips
        st.session_state["prepared"] = prepared
        st.success("ZIPs prepared with splitting")

    if "prepared" in st.session_state:
        st.subheader("Prepared ZIPs")
        preview = []
        for loc, zips in st.session_state["prepared"].items():
            for i, zp in enumerate(zips, start=1):
                preview.append({"Location": loc, "Part": i, "Zip": os.path.basename(zp), "Size": human_bytes(os.path.getsize(zp))})
        st.dataframe(pd.DataFrame(preview))

        if st.button("📤 Send Emails"):
            smtp_cfg = {"host": smtp_host, "port": int(smtp_port), "use_ssl": smtp_use_ssl, "sender": smtp_sender, "password": smtp_password}
            logs, total = [], sum(len(z) for z in st.session_state["prepared"].values())
            done, progress = 0, st.progress(0)
            for loc, zips in st.session_state["prepared"].items():
                recips = list(grouped[loc]["recipients"])
                if test_mode: recips = [test_email]
                total_parts = len(zips)
                for i, zp in enumerate(zips, start=1):
                    subj = subject_template.format(location=loc, part=i, total=total_parts)
                    body = body_template.format(location=loc, part=i, total=total_parts, footer=footer_text)
                    try:
                        send_email_smtp(smtp_cfg, recips, subj, body, [zp])
                        logs.append({"Location": loc, "Recipients": ", ".join(recips), "Zip": os.path.basename(zp), "Status": f"Sent Part {i}/{total_parts}"})
                        st.success(f"Sent {loc} Part {i}/{total_parts} → {', '.join(recips)}")
                    except Exception as e:
                        logs.append({"Location": loc, "Recipients": ", ".join(recips), "Zip": os.path.basename(zp), "Status": f"Failed: {e}"})
                        st.error(f"Failed {loc} Part {i}/{total_parts}: {e}")
                    done += 1; progress.progress(done/total)
                    time.sleep(delay_seconds)
            st.subheader("Logs")
            st.dataframe(pd.DataFrame(logs))
