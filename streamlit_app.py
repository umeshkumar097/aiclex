# streamlit_app.py (FINAL UNIQUE KEY FIXED)
import streamlit as st
import pandas as pd
import zipfile, os, io, tempfile, shutil, time, re
from collections import OrderedDict
from email.message import EmailMessage
import smtplib
from datetime import datetime

st.set_page_config(page_title="Aiclex Mailer — Final", layout="wide")
st.title("📧 Aiclex Hallticket Mailer — Final (With Unique Keys + Parts)")

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
                    continue

def create_chunked_zips(file_paths, out_dir, base_name, max_bytes):
    """Pack file_paths into sequential zip files each <= max_bytes."""
    os.makedirs(out_dir, exist_ok=True)
    parts = []
    cur_files, cur_size, part_index = [], 0, 1
    for fp in file_paths:
        fsz = os.path.getsize(fp)
        if fsz > max_bytes and not cur_files:
            zpath = os.path.join(out_dir, f"{base_name}_part{part_index}.zip")
            with zipfile.ZipFile(zpath, "w", compression=zipfile.ZIP_DEFLATED) as z:
                z.write(fp, arcname=os.path.basename(fp))
            parts.append(zpath)
            part_index += 1
            continue
        if cur_files and (cur_size + fsz) > max_bytes:
            zpath = os.path.join(out_dir, f"{base_name}_part{part_index}.zip")
            with zipfile.ZipFile(zpath, "w", compression=zipfile.ZIP_DEFLATED) as z:
                for f in cur_files:
                    z.write(f, arcname=os.path.basename(f))
            parts.append(zpath)
            part_index += 1
            cur_files, cur_size = [fp], fsz
        else:
            cur_files.append(fp)
            cur_size += fsz
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
    st.header("SMTP Settings")
    smtp_host = st.text_input("SMTP host", value="smtp.gmail.com")
    smtp_port = st.number_input("SMTP port", value=465)
    protocol = st.selectbox("Protocol", options=["SMTPS (SSL)", "SMTP (STARTTLS)"], index=0)
    smtp_use_ssl = True if protocol.startswith("SMTPS") else False
    sender_email = st.text_input("Sender email", value="info@cruxmanagement.com")
    sender_password = st.text_input("Sender password (App Password)", type="password", value="norx wxop hvsm bvfu")
    delay_seconds = st.number_input("Delay between emails (s)", value=2.0, step=0.5)
    per_attachment_limit_mb = st.number_input("Per-attachment limit (MB)", value=3.0, step=0.5)

# ---------------- Uploads ----------------
st.header("1) Upload Excel and ZIP")
uploaded_excel = st.file_uploader("Upload Excel (.xlsx/.csv)", type=["xlsx","csv"])
uploaded_zip = st.file_uploader("Upload ZIP (PDFs inside, nested zips ok)", type=["zip"])

if not uploaded_excel or not uploaded_zip:
    st.stop()

# ---------------- Read Excel ----------------
if uploaded_excel.name.lower().endswith(".csv"):
    df = pd.read_csv(uploaded_excel, dtype=str).fillna("")
else:
    df = pd.read_excel(uploaded_excel, dtype=str).fillna("")

cols = list(df.columns)
ht_col = st.selectbox("Hallticket column", cols, index=0)
email_col = st.selectbox("Email column", cols, index=1)
location_col = st.selectbox("Location column", cols, index=2)

st.write("Excel preview")
st.dataframe(df[[ht_col, email_col, location_col]].head(10))

# ---------------- Extract PDFs ----------------
extract_root = tempfile.mkdtemp(prefix="aiclex_extract_")
b = io.BytesIO(uploaded_zip.read())
extract_zip_recursively(b, extract_root)
pdf_paths = []
for root, _, files in os.walk(extract_root):
    for f in files:
        if f.lower().endswith(".pdf"):
            pdf_paths.append(os.path.join(root, f))
st.success(f"Extracted {len(pdf_paths)} PDFs")

# ---------------- Grouping ----------------
groups = OrderedDict()
for _, row in df.iterrows():
    ht = str(row.get(ht_col, "")).strip()
    loc = str(row.get(location_col, "")).strip()
    raw_emails = str(row.get(email_col, "")).strip()
    emails = [e.strip().lower() for e in re.split(r"[,;\n]+", raw_emails) if e.strip()]
    email_key = frozenset(emails)
    key = (loc, email_key)
    if key not in groups:
        groups[key] = {"halltickets": [], "matched_paths": []}
    groups[key]["halltickets"].append(ht)

pdf_lookup = {os.path.basename(p): p for p in pdf_paths}
for key, info in groups.items():
    matched = []
    for ht in info["halltickets"]:
        for fname, path in pdf_lookup.items():
            if ht and ht in fname:
                matched.append(path)
                break
    groups[key]["matched_paths"] = list(dict.fromkeys(matched))

summary = []
for (loc, email_key), info in groups.items():
    total_size = sum(os.path.getsize(p) for p in info["matched_paths"]) if info["matched_paths"] else 0
    summary.append({
        "Location": loc,
        "Recipients": ", ".join(email_key),
        "Tickets": len(info["halltickets"]),
        "Matched PDFs": len(info["matched_paths"]),
        "TotalSize": human_bytes(total_size)
    })
st.write("Group summary")
st.dataframe(pd.DataFrame(summary))

# ---------------- Prepare ZIPs ----------------
if st.button("2) Prepare ZIPs"):
    workdir = tempfile.mkdtemp(prefix="aiclex_zips_")
    max_bytes = int(per_attachment_limit_mb * 1024 * 1024)
    st.session_state.prepared = {}
    for (loc, email_key), info in groups.items():
        matched = info["matched_paths"]
        if not matched:
            continue
        base_name = re.sub(r"[^A-Za-z0-9]+", "_", loc)[:50]
        parts = create_chunked_zips(matched, workdir, base_name, max_bytes)
        st.session_state.prepared[(loc, tuple(email_key))] = {"parts": parts, "recipients": list(email_key)}
    st.success("ZIPs prepared successfully ✅")

if "prepared" in st.session_state:
    st.subheader("📦 Prepared Parts Summary")
    for (loc, emails), pdata in st.session_state.prepared.items():
        st.write(f"📍 Location: {loc} | Recipients: {', '.join(emails)} | Parts: {len(pdata['parts'])}")
        for idx, p in enumerate(pdata["parts"], start=1):
            size = human_bytes(os.path.getsize(p))
            with open(p, "rb") as f:
                st.download_button(
                    label=f"Download {os.path.basename(p)}",
                    data=f.read(),
                    file_name=os.path.basename(p),
                    key=f"dl_{loc}_{'_'.join(emails)}_{idx}_{int(time.time()*1000)}"
                )

# ---------------- Send ----------------
st.subheader("3) Send Emails")
subject_template = st.text_input("Subject template", value="Hall Tickets — {location}")
body_template = st.text_area("Body template", value="Dear Coordinator,\n\nPlease find attached the hall tickets for {location}.\n\nRegards,\nAiclex")

if st.button("🚀 Send All Emails"):
    if "prepared" not in st.session_state:
        st.error("No ZIPs prepared.")
    else:
        logs = []
        total_parts = sum(len(p["parts"]) for p in st.session_state.prepared.values())
        sent_count = 0
        prog = st.progress(0)
        for (loc, emails), pdata in st.session_state.prepared.items():
            for idx, p in enumerate(pdata["parts"], start=1):
                subj = f"{subject_template.format(location=loc)} (Part {idx}/{len(pdata['parts'])})"
                body = body_template.format(location=loc)
                try:
                    send_email_smtp(
                        smtp_host, smtp_port, smtp_use_ssl,
                        sender_email, sender_password,
                        pdata["recipients"], subj, body, [p]
                    )
                    logs.append({"Location": loc, "Recipients": ", ".join(emails), "Part": f"{idx}/{len(pdata['parts'])}", "Zip": os.path.basename(p), "Status": "✅ Sent"})
                except Exception as e:
                    logs.append({"Location": loc, "Recipients": ", ".join(emails), "Part": f"{idx}/{len(pdata['parts'])}", "Zip": os.path.basename(p), "Status": f"❌ Failed {e}"})
                sent_count += 1
                prog.progress(int(sent_count*100/total_parts))
                time.sleep(delay_seconds)
        st.write("📊 Sending Logs")
        st.dataframe(pd.DataFrame(logs))
        st.success("🎉 All sending attempts completed!")

# ---------------- Cleanup ----------------
if st.button("🧹 Cleanup Temporary Files"):
    try:
        shutil.rmtree(extract_root, ignore_errors=True)
        if "prepared" in st.session_state:
            for pdata in st.session_state.prepared.values():
                for p in pdata["parts"]:
                    try: os.remove(p)
                    except: pass
            st.session_state.pop("prepared")
        st.success("Cleaned temporary files.")
    except Exception as e:
        st.error(f"Cleanup failed: {e}")
