import streamlit as st
import pandas as pd
import zipfile, os, io, tempfile, shutil, time
from collections import defaultdict
from email.message import EmailMessage
import smtplib
from datetime import datetime

# ---------------- Helpers ----------------
def extract_zip_recursively(zip_file_like, extract_to):
    """Extract zip (file-like or path) recursively (nested zips)."""
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

def human_bytes(n):
    n = float(n)
    for unit in ['B','KB','MB','GB']:
        if n < 1024:
            return f"{n:.2f}{unit}"
        n /= 1024
    return f"{n:.2f}TB"

def create_chunked_zips(file_paths, out_dir, base_name, max_bytes):
    """Split into multiple zips if size > max_bytes."""
    os.makedirs(out_dir, exist_ok=True)
    parts = []
    current_files = []
    part_index = 1
    for fp in file_paths:
        current_files.append(fp)
        test_path = os.path.join(out_dir, f"__test_{part_index}.zip")
        with zipfile.ZipFile(test_path, 'w', compression=zipfile.ZIP_DEFLATED) as z:
            for f in current_files:
                z.write(f, arcname=os.path.basename(f))
        size = os.path.getsize(test_path)
        if size <= max_bytes:
            os.remove(test_path)
            continue
        else:
            last = current_files.pop()
            part_path = os.path.join(out_dir, f"{base_name}_part{part_index}.zip")
            with zipfile.ZipFile(part_path, 'w', compression=zipfile.ZIP_DEFLATED) as z:
                for f in current_files:
                    z.write(f, arcname=os.path.basename(f))
            parts.append(part_path)
            part_index += 1
            current_files = [last]
            os.remove(test_path)
    if current_files:
        part_path = os.path.join(out_dir, f"{base_name}_part{part_index}.zip")
        with zipfile.ZipFile(part_path, 'w', compression=zipfile.ZIP_DEFLATED) as z:
            for f in current_files:
                z.write(f, arcname=os.path.basename(f))
        parts.append(part_path)
    return parts

# ---------------- UI ----------------
st.title("📧 Aiclex Hallticket Mailer")
st.markdown("Upload Excel + ZIP → Match halltickets → Group by (Location+Emails) → Send clean emails")

# Sidebar SMTP settings
with st.sidebar:
    st.header("SMTP Settings")
    smtp_host = st.text_input("SMTP host", value="smtp.gmail.com")
    smtp_port = st.number_input("SMTP port", value=465)
    protocol = st.selectbox("Protocol", ["SMTPS (SSL)", "SMTP (STARTTLS)"], index=0)
    sender_email = st.text_input("Sender email", value="info@cruxmanagement.com")
    sender_password = st.text_input("Sender password", value="norx wxop hvsm bvfu", type="password")
    delay_seconds = st.number_input("Delay between emails (sec)", value=2.0)

# File upload
uploaded_excel = st.file_uploader("Upload Excel (.xlsx/.csv)", type=["xlsx","csv"])
uploaded_zip = st.file_uploader("Upload ZIP (pdfs inside, zip-in-zip supported)", type=["zip"])

if uploaded_excel and uploaded_zip:
    # Read Excel
    if uploaded_excel.name.endswith("csv"):
        df = pd.read_csv(uploaded_excel, dtype=str).fillna("")
    else:
        df = pd.read_excel(uploaded_excel, dtype=str).fillna("")
    
    cols = list(df.columns)
    ht_col = st.selectbox("Hallticket column", cols, index=0)
    email_col = st.selectbox("Email column", cols, index=1)
    location_col = st.selectbox("Location column", cols, index=2)

    # Extract zip
    temp_dir = tempfile.mkdtemp(prefix="aiclex_zip_")
    bio = io.BytesIO(uploaded_zip.read())
    extract_zip_recursively(bio, temp_dir)

    pdf_files = {}
    for root, _, files in os.walk(temp_dir):
        for f in files:
            if f.lower().endswith(".pdf"):
                pdf_files[f] = os.path.join(root, f)

    # Group by (location + frozenset(emails))
    grouped = defaultdict(list)
    for _, r in df.iterrows():
        ht = str(r[ht_col]).strip()
        loc = str(r[location_col]).strip()
        emails = [e.strip().lower() for e in str(r[email_col]).split(",") if e.strip()]
        email_key = frozenset(emails)
        key = (loc, email_key)
        grouped[key].append(ht)

    # Preview
    st.subheader("📋 Matching Preview")
    preview = []
    for (loc, email_set), hts in grouped.items():
        matched = []
        for ht in hts:
            for fn, path in pdf_files.items():
                if ht in fn:
                    matched.append(fn)
                    break
        preview.append({"Location": loc, "Emails": ", ".join(email_set), "Halltickets": len(hts), "MatchedPDFs": len(matched)})
    st.dataframe(pd.DataFrame(preview))

    # Subject/Body
    subject_template = st.text_input("Subject template", value="Hall Tickets — {location}")
    body_template = st.text_area("Body template", value="Dear Coordinator,\n\nPlease find attached the hall tickets for {location}.\n\nRegards,\nAiclex")

    # Prepare and send
    if st.button("Prepare & Send Emails"):
        max_bytes = 3 * 1024 * 1024  # 3 MB
        logs = []
        workdir = tempfile.mkdtemp(prefix="aiclex_work_")
        for (loc, email_set), hts in grouped.items():
            matched_paths = []
            for ht in hts:
                for fn, path in pdf_files.items():
                    if ht in fn:
                        matched_paths.append(path)
                        break
            if not matched_paths:
                logs.append({"Location": loc, "Emails": ", ".join(email_set), "Status": "⚠️ No PDFs"})
                continue
            out_dir = os.path.join(workdir, f"{loc}_{'_'.join(email_set)}")
            os.makedirs(out_dir, exist_ok=True)
            parts = create_chunked_zips(matched_paths, out_dir, f"{loc}", max_bytes)
            # send each part
            try:
                if protocol.startswith("SMTPS"):
                    server = smtplib.SMTP_SSL(smtp_host, int(smtp_port), timeout=60)
                else:
                    server = smtplib.SMTP(smtp_host, int(smtp_port), timeout=60)
                    server.starttls()
                server.login(sender_email, sender_password)
                for idx, part in enumerate(parts, start=1):
                    msg = EmailMessage()
                    msg['From'] = sender_email
                    msg['To'] = ", ".join(email_set)  # multiple TO from one row
                    msg['Subject'] = f"{subject_template.format(location=loc)} (Part {idx}/{len(parts)})"
                    msg.set_content(body_template.format(location=loc))
                    with open(part, "rb") as f:
                        msg.add_attachment(f.read(), maintype="application", subtype="zip", filename=os.path.basename(part))
                    server.send_message(msg)
                    logs.append({"Location": loc, "Emails": ", ".join(email_set), "Zip": os.path.basename(part), "Status": "✅ Sent", "Time": datetime.now().strftime("%H:%M:%S")})
                    time.sleep(float(delay_seconds))
                server.quit()
            except Exception as e:
                logs.append({"Location": loc, "Emails": ", ".join(email_set), "Status": f"❌ Failed {e}"})
        st.success("Done sending. Log below:")
        st.dataframe(pd.DataFrame(logs))
