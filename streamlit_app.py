import streamlit as st
import pandas as pd
import zipfile, os, io, tempfile, shutil, time, re
from collections import defaultdict
from email.message import EmailMessage
import smtplib

# ---------- Helpers ----------
def human_bytes(n):
    for unit in ['B','KB','MB','GB']:
        if n < 1024:
            return f"{n:.2f}{unit}"
        n /= 1024
    return f"{n:.2f}TB"

def extract_zip_recursive(uploaded_file, extract_to):
    with zipfile.ZipFile(uploaded_file, "r") as z:
        z.extractall(extract_to)
    for root, _, files in os.walk(extract_to):
        for f in files:
            if f.lower().endswith(".zip"):
                nested_path = os.path.join(root, f)
                nested_dir = os.path.join(root, f.replace(".zip", "_nested"))
                os.makedirs(nested_dir, exist_ok=True)
                with zipfile.ZipFile(nested_path, "r") as nested:
                    nested.extractall(nested_dir)

def create_chunked_zip(files, out_dir, base_name, max_size=3*1024*1024):
    os.makedirs(out_dir, exist_ok=True)
    zips = []
    current_files, idx = [], 1

    for f in files:
        current_files.append(f)
        test_zip = os.path.join(out_dir, f"test_{idx}.zip")
        with zipfile.ZipFile(test_zip, "w", compression=zipfile.ZIP_DEFLATED) as z:
            for cf in current_files:
                z.write(cf, os.path.basename(cf))
        if os.path.getsize(test_zip) > max_size:
            current_files.pop()
            final_zip = os.path.join(out_dir, f"{base_name}_part{idx}.zip")
            with zipfile.ZipFile(final_zip, "w", compression=zipfile.ZIP_DEFLATED) as z:
                for cf in current_files:
                    z.write(cf, os.path.basename(cf))
            zips.append(final_zip)
            idx += 1
            current_files = [f]
        os.remove(test_zip)

    if current_files:
        final_zip = os.path.join(out_dir, f"{base_name}_part{idx}.zip")
        with zipfile.ZipFile(final_zip, "w", compression=zipfile.ZIP_DEFLATED) as z:
            for cf in current_files:
                z.write(cf, os.path.basename(cf))
        zips.append(final_zip)

    return zips

# ---------- Streamlit UI ----------
st.set_page_config(page_title="Aiclex Mailer Final", layout="wide")
st.title("📧 Aiclex Mailer — Final Version")

# ---------- Sidebar ----------
with st.sidebar:
    st.header("📧 Email Settings")
    smtp_host = st.text_input("SMTP Host", value="smtp.gmail.com")
    smtp_port = st.number_input("SMTP Port", value=465)
    sender_email = st.text_input("Sender Email", value="info@cruxmanagement.com")
    sender_pass = st.text_input("App Password", type="password")

    subject_template = st.text_input("Subject Template", value="Hall Tickets for {location} - Part {part}")
    body_template = st.text_area(
        "Body Template",
        value="Dear Team,\n\nPlease find attached hall tickets for {location}.\nThis is part {part}.\n\nRegards,\nAiclex Technologies"
    )

    size_limit_mb = st.number_input("Attachment Limit (MB)", value=3.0, step=0.5)
    delay_seconds = st.number_input("Delay Between Emails (seconds)", value=2.0, step=0.5)

    st.markdown("---")
    testing_mode = st.checkbox("Enable Testing Mode", value=False)
    test_email = st.text_input("Test Email (for Testing Mode)", value="info@aiclex.in")

# ---------- File Upload ----------
uploaded_excel = st.file_uploader("Upload Excel", type=["xlsx", "csv"])
uploaded_zip = st.file_uploader("Upload ZIP (pdfs)", type=["zip"])

if uploaded_excel and uploaded_zip:
    # --- Step 1: Read Excel ---
    if uploaded_excel.name.endswith("csv"):
        df = pd.read_csv(uploaded_excel, dtype=str).fillna("")
    else:
        df = pd.read_excel(uploaded_excel, dtype=str).fillna("")

    st.write("### Preview of Uploaded Data")
    st.dataframe(df.head(10))

    # Detect columns
    cols = df.columns.tolist()
    ht_col = st.selectbox("Hallticket Column", cols, index=0)
    email_col = st.selectbox("Email Column", cols, index=1)
    loc_col = st.selectbox("Location Column", cols, index=2)

    # --- Step 2: Extract PDFs ---
    temp_dir = tempfile.mkdtemp()
    extract_zip_recursive(uploaded_zip, temp_dir)

    pdf_map = {}
    for root, _, files in os.walk(temp_dir):
        for f in files:
            if f.lower().endswith(".pdf"):
                pdf_map[f] = os.path.join(root, f)

    st.success(f"Extracted {len(pdf_map)} PDFs")

    # --- Step 3: Group by Location + Email(s) ---
    grouped = defaultdict(list)
    for _, row in df.iterrows():
        hall = str(row[ht_col]).strip()
        location = str(row[loc_col]).strip()
        raw_emails = str(row[email_col]).strip()
        emails = [e.strip().lower() for e in re.split(r"[,;\n]+", raw_emails) if e.strip()]

        matched_pdf = ""
        for fn in pdf_map:
            if hall in fn:
                matched_pdf = pdf_map[fn]
                break

        if matched_pdf:
            key = (location, tuple(sorted(emails)))
            grouped[key].append(matched_pdf)

    st.write("### Grouping Summary")
    summary = []
    for (loc, emails), files in grouped.items():
        summary.append({"Location": loc, "Recipients": ", ".join(emails), "Tickets": len(files)})
    st.dataframe(pd.DataFrame(summary))

    # --- Step 4: Prepare Zips ---
    if st.button("Prepare Zips"):
        prepared = {}
        for (loc, emails), files in grouped.items():
            out_dir = os.path.join(tempfile.mkdtemp(), loc.replace(" ", "_"))
            zips = create_chunked_zip(files, out_dir, loc, max_size=int(size_limit_mb*1024*1024))
            prepared[(loc, emails)] = zips

        st.session_state["prepared"] = prepared
        st.success("ZIPs prepared successfully ✅")

    # --- Step 5: Show Prepared Zips in Table with Download Column ---
    if "prepared" in st.session_state:
        st.subheader("📦 Prepared ZIP Parts Summary")
        rows = []
        for (loc, emails), zips in st.session_state["prepared"].items():
            for idx, zp in enumerate(zips, 1):
                rows.append({
                    "Location": loc,
                    "Recipients": ", ".join(emails),
                    "Part": f"{idx}/{len(zips)}",
                    "File": os.path.basename(zp),
                    "Size": human_bytes(os.path.getsize(zp)),
                    "Path": zp
                })
        df_summary = pd.DataFrame(rows)[["Location", "Recipients", "Part", "File", "Size"]]
        st.dataframe(df_summary, use_container_width=True)

        for idx, row in enumerate(rows):
            with open(row["Path"], "rb") as f:
                st.download_button(
                    label=f"⬇️ Download {row['File']}",
                    data=f.read(),
                    file_name=row["File"],
                    key=f"dl_{idx}"
                )

    # --- Step 6: Send Emails ---
    if "prepared" in st.session_state and st.button("Send All Emails"):
        try:
            server = smtplib.SMTP_SSL(smtp_host, int(smtp_port))
            server.login(sender_email, sender_pass)
            logs = []

            for (loc, emails), zips in st.session_state["prepared"].items():
                for idx, zp in enumerate(zips, 1):
                    msg = EmailMessage()
                    msg["From"] = sender_email
                    if testing_mode:
                        msg["To"] = test_email
                    else:
                        msg["To"] = ", ".join(emails)

                    msg["Subject"] = subject_template.format(location=loc, part=idx)
                    body = body_template.format(location=loc, part=idx)
                    msg.set_content(body)

                    with open(zp, "rb") as f:
                        msg.add_attachment(
                            f.read(),
                            maintype="application",
                            subtype="zip",
                            filename=os.path.basename(zp)
                        )

                    try:
                        server.send_message(msg)
                        logs.append({
                            "Location": loc,
                            "Part": idx,
                            "Recipients": msg["To"],
                            "Status": "Sent"
                        })
                    except Exception as e:
                        logs.append({
                            "Location": loc,
                            "Part": idx,
                            "Recipients": msg["To"],
                            "Status": f"Failed: {e}"
                        })
                    time.sleep(delay_seconds)

            server.quit()
            st.success("All emails attempted ✅")
            st.dataframe(pd.DataFrame(logs))

        except Exception as e:
            st.error(f"SMTP Error: {e}")
