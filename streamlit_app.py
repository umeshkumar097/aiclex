# streamlit_app.py
import streamlit as st
import pandas as pd
import zipfile, os, io, tempfile, shutil, time, re
from collections import defaultdict
import smtplib
from email.message import EmailMessage

# ---------------- Config ----------------
st.set_page_config(page_title="Aiclex Hallticket Mailer", layout="wide")

# ---------------- Sidebar ----------------
with st.sidebar:
    st.image("https://aiclex.in/wp-content/uploads/2024/08/aiclex-logo.png", width=180)
    st.markdown("### 📧 Email Configuration")

    smtp_host = st.text_input("SMTP Host", "smtp.gmail.com")
    smtp_port = st.number_input("SMTP Port", value=465)
    sender_email = st.text_input("Sender Email", "info@cruxmanagement.com")
    sender_pass = st.text_input("App Password", "norx wxop hvsm bvfu", type="password")

    st.markdown("### ✉️ Templates")
    subject_template = st.text_input("Subject", "Halltickets for {location} (Part {part}/{total})")
    body_template = st.text_area("Body", 
        "Dear Coordinator,\n\nPlease find attached the halltickets for {location}.\n\nRegards,\nAiclex Technologies"
    )

    st.markdown("### ⚙️ Settings")
    delay_seconds = st.number_input("Delay between emails (seconds)", value=2.0, step=0.5)
    max_mb = st.number_input("Attachment limit (MB)", value=3.0, step=0.5)

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
    parts, cur, size, part = [], [], 0, 1
    for f in file_list:
        sz = os.path.getsize(f)
        if size + sz > max_bytes and cur:
            zip_path = os.path.join(out_dir, f"{base_name}_part{part}.zip")
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
                for cf in cur:
                    z.write(cf, os.path.basename(cf))
            parts.append(zip_path)
            cur, size, part = [], 0, part+1
        cur.append(f)
        size += sz
    if cur:
        zip_path = os.path.join(out_dir, f"{base_name}_part{part}.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            for cf in cur:
                z.write(cf, os.path.basename(cf))
        parts.append(zip_path)
    return parts

# ---------------- Main Page ----------------
st.title("🎯 Aiclex Hallticket Mailer — Final Clean Version")

excel_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
zip_file = st.file_uploader("Upload ZIP (PDFs, supports nested zips)", type=["zip"])

if excel_file and zip_file:
    df = pd.read_excel(excel_file, dtype=str).fillna("")
    st.success(f"Loaded Excel with {len(df)} rows")
    st.dataframe(df.head(), width="stretch")

    tmpdir = tempfile.mkdtemp()
    extract_zip_all(zip_file, tmpdir)

    # Map PDFs by hallticket
    pdf_map = {}
    for root, _, files in os.walk(tmpdir):
        for f in files:
            if f.endswith(".pdf"):
                hall_id = re.findall(r"\d{6,}", f)
                if hall_id:
                    pdf_map[hall_id[0]] = os.path.join(root, f)

    st.info(f"Extracted {len(pdf_map)} PDFs from uploaded ZIP(s).")

    # Mapping Table
    mapping_rows = []
    for _, row in df.iterrows():
        hall = str(row.iloc[0]).strip()
        raw_emails = str(row.iloc[1]).strip()
        location = str(row.iloc[2]).strip()
        matched_file = pdf_map.get(hall, "")
        mapping_rows.append({
            "Hallticket": hall,
            "Emails": raw_emails,
            "Location": location,
            "MatchedFile": os.path.basename(matched_file) if matched_file else "❌ Not Found"
        })
    df_map = pd.DataFrame(mapping_rows)
    st.subheader("📋 Mapping Table (Hallticket ↔ PDF)")
    st.dataframe(df_map, width="stretch")

    # Grouping
    groups = defaultdict(lambda: defaultdict(list))
    for _, row in df.iterrows():
        hall = str(row.iloc[0]).strip()
        raw_emails = str(row.iloc[1]).strip()
        location = str(row.iloc[2]).strip()
        emails = [e.strip().lower() for e in re.split(r"[,;\s]+", raw_emails) if e.strip()]
        matched_pdf = pdf_map.get(hall)
        if matched_pdf:
            groups[location][tuple(emails)].append(matched_pdf)

    rows, all_prepared = [], {}
    for loc, recips in groups.items():
        for recip_tuple, files in recips.items():
            recip_str = ", ".join(recip_tuple)
            out_dir = os.path.join(tmpdir, f"{loc}_{recip_str.replace('@','_')}")
            parts = chunk_zip(files, out_dir, f"{loc}", int(max_mb*1024*1024))
            all_prepared[(loc, recip_str)] = parts
            for i, p in enumerate(parts, 1):
                rows.append({
                    "Location": loc,
                    "Recipients": recip_str,
                    "File": os.path.basename(p),
                    "Part": f"{i}/{len(parts)}",
                    "Path": p
                })

    df_summary = pd.DataFrame(rows)
    st.subheader("📦 Prepared ZIP Parts Summary")
    st.dataframe(df_summary, width="stretch")

    # Download links
    for _, row in df_summary.iterrows():
        with open(row["Path"], "rb") as f:
            st.download_button(
                f"Download {row['File']} ({row['Location']} {row['Part']})",
                f.read(),
                file_name=row["File"],
                key=f"dl_{row['Location']}_{row['Recipients']}_{row['File']}_{row['Part']}_{int(time.time()*1000)}"
            )

    # --- Testing Section ---
    st.subheader("🔬 Testing Option")
    test_email = st.text_input("Enter a test email address", sender_email)

    if st.button("📤 Send Test Email"):
        try:
            server = smtplib.SMTP_SSL(smtp_host, int(smtp_port))
            server.login(sender_email, sender_pass)

            for (loc, recip_str), parts in all_prepared.items():
                msg = EmailMessage()
                msg["From"] = sender_email
                msg["To"] = test_email
                msg["Subject"] = f"[TEST] {subject_template.format(location=loc, part=1, total=len(parts))}"
                msg.set_content(
                    body_template.format(location=loc, part=1, total=len(parts))
                    + "\n\n---\n(This is a TEST email, only first part attached.)"
                )
                with open(parts[0], "rb") as f:
                    msg.add_attachment(
                        f.read(),
                        maintype="application",
                        subtype="zip",
                        filename=os.path.basename(parts[0])
                    )
                server.send_message(msg)
                break
            server.quit()
            st.success(f"✅ Test email sent to {test_email}")
        except Exception as e:
            st.error(f"❌ Test failed: {e}")

    # --- Bulk Sending ---
    st.subheader("🚀 Bulk Sending")
    if st.button("Send All Emails"):
        try:
            server = smtplib.SMTP_SSL(smtp_host, int(smtp_port))
            server.login(sender_email, sender_pass)
            prog = st.progress(0)
            total_jobs = sum(len(p) for p in all_prepared.values())
            done = 0
            for (loc, recip_str), parts in all_prepared.items():
                for idx, p in enumerate(parts, 1):
                    msg = EmailMessage()
                    msg["From"] = sender_email
                    msg["To"] = recip_str
                    msg["Subject"] = subject_template.format(location=loc, part=idx, total=len(parts))
                    msg.set_content(body_template.format(location=loc, part=idx, total=len(parts)))
                    with open(p, "rb") as f:
                        msg.add_attachment(f.read(), maintype="application", subtype="zip", filename=os.path.basename(p))
                    server.send_message(msg)
                    done += 1
                    prog.progress(done/total_jobs)
                    time.sleep(float(delay_seconds))
            server.quit()
            st.success("✅ All emails sent successfully.")
        except Exception as e:
            st.error(f"❌ Sending failed: {e}")

    if st.button("🧹 Cleanup temporary files"):
        try:
            shutil.rmtree(tmpdir)
            st.success("Temporary files deleted.")
        except:
            st.warning("Cleanup skipped.")
