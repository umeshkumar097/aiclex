# streamlit_app.py
import streamlit as st
import pandas as pd
import zipfile, os, io, tempfile, shutil, time, re
from collections import defaultdict
import smtplib
from email.message import EmailMessage
from datetime import datetime

# ---------------- Settings ----------------
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 465
SENDER_EMAIL = "info@cruxmanagement.com"
SENDER_PASS = "norx wxop hvsm bvfu"   # App password

MAX_ATTACHMENT_MB = 3

st.set_page_config(page_title="Aiclex Mailer", layout="wide")
st.title("📧 Aiclex Technologies — Hallticket Mailer (Final Version)")

# ---------------- Helpers ----------------
def extract_zip_all(uploaded_file, extract_to):
    with zipfile.ZipFile(uploaded_file, 'r') as z:
        z.extractall(extract_to)
    for root, _, files in os.walk(extract_to):
        for f in files:
            if f.endswith(".zip"):
                nested = os.path.join(root, f)
                with zipfile.ZipFile(nested, 'r') as z2:
                    z2.extractall(root)

def chunk_zip(file_list, out_dir, base_name, max_bytes):
    os.makedirs(out_dir, exist_ok=True)
    parts = []
    cur, size = [], 0
    part = 1
    for f in file_list:
        sz = os.path.getsize(f)
        if size + sz > max_bytes and cur:
            zip_path = os.path.join(out_dir, f"{base_name}_part{part}.zip")
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
                for cf in cur:
                    z.write(cf, os.path.basename(cf))
            parts.append(zip_path)
            part += 1
            cur, size = [], 0
        cur.append(f)
        size += sz
    if cur:
        zip_path = os.path.join(out_dir, f"{base_name}_part{part}.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            for cf in cur:
                z.write(cf, os.path.basename(cf))
        parts.append(zip_path)
    return parts

# ---------------- Upload ----------------
excel_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
zip_file = st.file_uploader("Upload ZIP (PDFs inside)", type=["zip"])

subject_template = st.text_input("✉️ Subject template", "Halltickets for {location} (Part {part}/{total})")
body_template = st.text_area("📄 Body template", "Dear Coordinator,\n\nPlease find attached the halltickets for {location}.\n\nRegards,\nAiclex Technologies")

delay_seconds = st.number_input("⏳ Delay between emails (seconds)", value=2.0, step=0.5)

if excel_file and zip_file:
    df = pd.read_excel(excel_file, dtype=str).fillna("")
    st.write("Preview:", df.head())

    tmpdir = tempfile.mkdtemp()
    extract_zip_all(zip_file, tmpdir)

    pdf_map = {}
    for root, _, files in os.walk(tmpdir):
        for f in files:
            if f.endswith(".pdf"):
                hall_id = re.findall(r"\d{6,}", f)
                if hall_id:
                    pdf_map[hall_id[0]] = os.path.join(root, f)

    st.success(f"Extracted {len(pdf_map)} PDFs.")

    # Process group
    groups = defaultdict(lambda: defaultdict(list))
    for _, row in df.iterrows():
        hall = str(row[0]).strip()
        emails = re.split(r"[,;\s]+", str(row[1]))
        location = str(row[2]).strip()
        matched_pdf = pdf_map.get(hall)
        if matched_pdf:
            for e in emails:
                if e:
                    groups[location][e.lower()].append(matched_pdf)

    rows = []
    all_prepared = {}
    for loc, recips in groups.items():
        for recip, files in recips.items():
            out_dir = os.path.join(tmpdir, f"{loc}_{recip}")
            parts = chunk_zip(files, out_dir, f"{loc}", int(MAX_ATTACHMENT_MB*1024*1024))
            all_prepared[(loc, recip)] = parts
            for i, p in enumerate(parts, 1):
                rows.append({
                    "Location": loc,
                    "Recipient": recip,
                    "File": os.path.basename(p),
                    "Part": f"{i}/{len(parts)}",
                    "Path": p
                })

    df_summary = pd.DataFrame(rows)
    st.subheader("📦 Prepared ZIP Parts Summary")
    st.dataframe(df_summary, width="stretch")

    for _, row in df_summary.iterrows():
        with open(row["Path"], "rb") as f:
            st.download_button(
                f"Download {row['File']} ({row['Location']} {row['Part']})",
                f.read(),
                file_name=row["File"],
                key=f"dl_{row['Location']}_{row['Recipient']}_{row['File']}_{row['Part']}_{int(time.time()*1000)}"
            )

    if st.button("🚀 Send All Emails"):
        try:
            server = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT)
            server.login(SENDER_EMAIL, SENDER_PASS)
            for (loc, recip), parts in all_prepared.items():
                for idx, p in enumerate(parts, 1):
                    msg = EmailMessage()
                    msg["From"] = SENDER_EMAIL
                    msg["To"] = recip
                    msg["Subject"] = subject_template.format(location=loc, part=idx, total=len(parts))
                    body = body_template.format(location=loc, part=idx, total=len(parts))
                    msg.set_content(body)
                    with open(p, "rb") as f:
                        msg.add_attachment(f.read(), maintype="application", subtype="zip", filename=os.path.basename(p))
                    server.send_message(msg)
                    st.write(f"✅ Sent to {recip} ({loc}, Part {idx}/{len(parts)})")
                    time.sleep(delay_seconds)
            server.quit()
            st.success("All emails sent successfully.")
        except Exception as e:
            st.error(f"Sending failed: {e}")

    if st.button("🧹 Cleanup temporary files"):
        try:
            shutil.rmtree(tmpdir)
            st.success("Temporary files deleted.")
        except:
            st.warning("Cleanup skipped.")

