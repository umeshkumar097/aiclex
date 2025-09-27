import streamlit as st
import pandas as pd
import zipfile, os, io, tempfile, shutil, time
from collections import defaultdict
from email.message import EmailMessage
import smtplib
from datetime import datetime
import re

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
st.set_page_config(page_title="Aiclex Hallticket Mailer (Preview-enabled)", layout="wide")
st.title("📧 Aiclex Hallticket Mailer (Preview before Send)")

# Sidebar SMTP settings (defaults provided)
with st.sidebar:
    st.header("SMTP Settings (Gmail default)")
    smtp_host = st.text_input("SMTP host", value="smtp.gmail.com")
    smtp_port = st.number_input("SMTP port", value=465)
    protocol = st.selectbox("Protocol", ["SMTPS (SSL)", "SMTP (STARTTLS)"], index=0)
    sender_email = st.text_input("Sender email", value="info@cruxmanagement.com")
    sender_password = st.text_input("Sender password (App password)", value="norx wxop hvsm bvfu", type="password")
    delay_seconds = st.number_input("Delay between emails (sec)", value=2.0, step=0.5)
    st.markdown("---")
    st.subheader("Testing")
    testing_mode = st.checkbox("Enable Testing Mode (send to test email instead of real recipients)", value=True)
    test_email = st.text_input("Test recipient email", value="info@aiclex.in")

# File upload
uploaded_excel = st.file_uploader("Upload Excel (.xlsx/.csv)", type=["xlsx","csv"])
uploaded_zip = st.file_uploader("Upload ZIP (pdfs inside, nested zips supported)", type=["zip"])

# Session-state holders for preview/workdir/prepared
if "workdir" not in st.session_state:
    st.session_state["workdir"] = None
if "pdf_map" not in st.session_state:
    st.session_state["pdf_map"] = {}
if "grouped" not in st.session_state:
    st.session_state["grouped"] = {}
if "prepared" not in st.session_state:
    st.session_state["prepared"] = {}

# When files uploaded -> process mapping and show preview
if uploaded_excel and uploaded_zip:
    # Read Excel (preserve current behavior)
    try:
        if uploaded_excel.name.lower().endswith("csv"):
            df = pd.read_csv(uploaded_excel, dtype=str).fillna("")
        else:
            df = pd.read_excel(uploaded_excel, dtype=str).fillna("")
    except Exception as e:
        st.error("Failed to read Excel: " + str(e))
        st.stop()

    cols = list(df.columns)
    # Let user choose which columns map where (keeps flexibility)
    ht_col = st.selectbox("Hallticket column", cols, index=0)
    email_col = st.selectbox("Email column", cols, index=1 if len(cols)>1 else 0)
    location_col = st.selectbox("Location column", cols, index=2 if len(cols)>2 else 0)

    st.subheader("1) Data Preview (first 10 rows)")
    st.dataframe(df[[ht_col, email_col, location_col]].head(10), width="stretch")

    # Extract ZIPs into a workdir and build pdf_map
    if st.session_state["workdir"] is None:
        st.session_state["workdir"] = tempfile.mkdtemp(prefix="aiclex_zip_")
    workdir = st.session_state["workdir"]

    bio = io.BytesIO(uploaded_zip.read())
    try:
        extract_zip_recursively(bio, workdir)
    except Exception as e:
        st.error("ZIP extraction failed: " + str(e))
        st.stop()

    # collect PDFs
    pdf_files = {}
    for root, _, files in os.walk(workdir):
        for f in files:
            if f.lower().endswith(".pdf"):
                pdf_files[f] = os.path.join(root, f)
    st.success(f"Extracted {len(pdf_files)} PDF files into workspace: {workdir}")
    st.session_state["pdf_map"] = pdf_files

    # Build grouping (location + frozenset(emails)) -> list of hallticket numbers
    grouped = defaultdict(list)
    for _, r in df.iterrows():
        ht = str(r[ht_col]).strip()
        loc = str(r[location_col]).strip()
        raw_emails = str(r[email_col]).strip()
        # split on commas/semicolons/newlines
        emails = [e.strip().lower() for e in re.split(r"[,;\n]+", raw_emails) if e.strip()]
        email_key = frozenset(emails)
        key = (loc, email_key)
        grouped[key].append(ht)
    st.session_state["grouped"] = grouped

    # Create Mapping Table (Hallticket -> MatchedFile or Not Found)
    mapping_rows = []
    for _, r in df.iterrows():
        ht = str(r[ht_col]).strip()
        emails = str(r[email_col]).strip()
        loc = str(r[location_col]).strip()
        matched = ""
        for fn, path in pdf_files.items():
            if ht and ht in fn:
                matched = fn
                break
        mapping_rows.append({
            "Hallticket": ht,
            "Emails": emails,
            "Location": loc,
            "MatchedFile": matched if matched else "NOT FOUND"
        })
    map_df = pd.DataFrame(mapping_rows)
    st.subheader("2) Mapping Table (Hallticket → PDF)")
    st.dataframe(map_df, width="stretch")

    st.markdown("---")
    st.subheader("3) Prepare ZIPs (Preview before sending)")

    # Button: prepare zips from grouped data (but do not send)
    if st.button("Prepare ZIPs (create parts for each group)"):
        max_bytes = int(3 * 1024 * 1024)  # 3 MB default per requirement
        prepared = {}   # key: (loc, recipients_str) -> list of zip paths
        summary_rows = []  # for preview table
        tmp_outroot = tempfile.mkdtemp(prefix="aiclex_out_")
        for (loc, email_set), hts in grouped.items():
            # find matched pdf paths for these halltickets
            matched_paths = []
            for ht in hts:
                found = None
                for fn, path in pdf_files.items():
                    if ht and ht in fn:
                        found = path
                        break
                if found:
                    matched_paths.append(found)
            if not matched_paths:
                # no matched files for this group
                prepared[(loc, ", ".join(sorted(list(email_set))))] = []
                continue
            out_dir = os.path.join(tmp_outroot, f"{loc}_{'_'.join(sorted(list(email_set))).replace('@','_')}")
            os.makedirs(out_dir, exist_ok=True)
            parts = create_chunked_zips(matched_paths, out_dir, base_name=loc.replace(" ", "_")[:60], max_bytes=max_bytes)
            recipients_str = ", ".join(sorted(list(email_set)))
            prepared[(loc, recipients_str)] = parts
            for idx, p in enumerate(parts, start=1):
                summary_rows.append({
                    "Location": loc,
                    "Recipients": recipients_str,
                    "Part": f"{idx}/{len(parts)}",
                    "File": os.path.basename(p),
                    "Size": human_bytes(os.path.getsize(p)),
                    "Path": p
                })
        # save prepared into session so user can preview and then send
        st.session_state["prepared"] = prepared
        st.session_state["summary_rows"] = summary_rows
        st.success("Prepared zips and parts created (preview available below).")

    # If prepared available show preview table + downloads
    if st.session_state.get("summary_rows"):
        st.subheader("4) Prepared ZIP Parts Preview")
        preview_df = pd.DataFrame(st.session_state["summary_rows"])
        # show table
        st.dataframe(preview_df[["Location","Recipients","Part","File","Size"]], width="stretch")

        st.markdown("**Download prepared parts**")
        # show download buttons with unique keys
        for idx, row in enumerate(st.session_state["summary_rows"]):
            ppath = row["Path"]
            fname = row["File"]
            label = f"⬇️ {row['Location']} — {fname} ({row['Part']})"
            try:
                with open(ppath, "rb") as f:
                    st.download_button(
                        label=label,
                        data=f.read(),
                        file_name=fname,
                        key=f"dl_part_{idx}"
                    )
            except Exception as e:
                st.warning(f"Unable to open {ppath}: {e}")

        st.markdown("---")
        st.subheader("5) Send (Test / Bulk)")

        # Test send: sends only first part of first prepared group
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Send Test Email (first prepared part only)"):
                if not st.session_state.get("prepared"):
                    st.error("No prepared parts found. Click 'Prepare ZIPs' first.")
                else:
                    # find first non-empty prepared entry
                    test_sent = False
                    try:
                        if protocol.startswith("SMTPS"):
                            server = smtplib.SMTP_SSL(smtp_host, int(smtp_port), timeout=60)
                        else:
                            server = smtplib.SMTP(smtp_host, int(smtp_port), timeout=60)
                            server.starttls()
                        server.login(sender_email, sender_password)
                        for (loc, recipients_str), parts in st.session_state["prepared"].items():
                            if not parts:
                                continue
                            first_part = parts[0]
                            msg = EmailMessage()
                            msg["From"] = sender_email
                            msg["To"] = test_email if testing_mode else recipients_str
                            msg["Subject"] = f"[TEST] {row.get('Location', loc)} - {os.path.basename(first_part)}"
                            msg.set_content(f"This is a test email for {loc}. Attached: {os.path.basename(first_part)}")
                            with open(first_part, "rb") as af:
                                msg.add_attachment(af.read(), maintype="application", subtype="zip", filename=os.path.basename(first_part))
                            server.send_message(msg)
                            test_sent = True
                            st.success(f"Test email sent to {msg['To']} with {os.path.basename(first_part)}")
                            break
                        server.quit()
                        if not test_sent:
                            st.warning("No parts available to test send.")
                    except Exception as e:
                        st.error("Test send failed: " + str(e))

        with col2:
            if st.button("Send ALL Prepared Parts (Bulk)"):
                if not st.session_state.get("prepared"):
                    st.error("No prepared parts found. Click 'Prepare ZIPs' first.")
                else:
                    total_parts = sum(len(v) for v in st.session_state["prepared"].values())
                    sent_count = 0
                    logs = []
                    prog = st.progress(0)
                    try:
                        if protocol.startswith("SMTPS"):
                            server = smtplib.SMTP_SSL(smtp_host, int(smtp_port), timeout=60)
                        else:
                            server = smtplib.SMTP(smtp_host, int(smtp_port), timeout=60)
                            server.starttls()
                        server.login(sender_email, sender_password)
                        for (loc, recipients_str), parts in st.session_state["prepared"].items():
                            if not parts:
                                logs.append({"Location": loc, "Recipients": recipients_str, "Part": "", "File": "", "Status": "No parts"})
                                continue
                            for idx, ppath in enumerate(parts, start=1):
                                msg = EmailMessage()
                                msg["From"] = sender_email
                                msg["To"] = test_email if testing_mode else recipients_str
                                subj = f"{subject_template if 'subject_template' in locals() else 'Hall Tickets — {location}'}"
                                # allow subject_template and body_template if defined below (keeps backward compat)
                                try:
                                    subject_line = subj.format(location=loc, part=idx, total=len(parts))
                                except Exception:
                                    subject_line = f"{loc} part {idx}/{len(parts)}"
                                msg["Subject"] = subject_line
                                body_txt = ""
                                if 'body_template' in locals():
                                    try:
                                        body_txt = body_template.format(location=loc, part=idx, total=len(parts))
                                    except Exception:
                                        body_txt = f"Please find attached part {idx} for {loc}."
                                else:
                                    body_txt = f"Please find attached part {idx} for {loc}."
                                msg.set_content(body_txt)
                                with open(ppath, "rb") as af:
                                    msg.add_attachment(af.read(), maintype="application", subtype="zip", filename=os.path.basename(ppath))
                                try:
                                    server.send_message(msg)
                                    sent_count += 1
                                    logs.append({"Location": loc, "Recipients": msg["To"], "Part": f"{idx}/{len(parts)}", "File": os.path.basename(ppath), "Status": "Sent", "Time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
                                except Exception as e:
                                    logs.append({"Location": loc, "Recipients": msg["To"], "Part": f"{idx}/{len(parts)}", "File": os.path.basename(ppath), "Status": f"Failed: {e}"})
                                prog.progress(sent_count / total_parts if total_parts>0 else 1.0)
                                time.sleep(float(delay_seconds))
                        server.quit()
                        st.success("Bulk send attempts complete.")
                        st.subheader("Sending Logs")
                        st.dataframe(pd.DataFrame(logs), width="stretch")
                    except Exception as e:
                        st.error("Bulk send failed: " + str(e))

    # Cleanup button for workdir
    st.markdown("---")
    if st.button("🧹 Cleanup workspace (delete extracted files & prepared zips)"):
        wd = st.session_state.get("workdir")
        # also remove prepared parts dirs when present
        try:
            if wd and os.path.exists(wd):
                shutil.rmtree(wd)
            # clear session prepared data
            if "prepared" in st.session_state:
                # try removing each prepared path's parent dirs
                for (_, _), parts in st.session_state["prepared"].items():
                    for p in parts:
                        try:
                            parent = os.path.dirname(p)
                            if parent and os.path.exists(parent):
                                shutil.rmtree(parent)
                        except:
                            pass
                st.session_state["prepared"] = {}
            st.session_state["workdir"] = None
            st.session_state["summary_rows"] = []
            st.success("Workspace cleaned.")
        except Exception as e:
            st.error("Cleanup failed: " + str(e))

else:
    st.info("Upload both Excel and ZIP to begin. After upload: 1) Mapping Table will show, 2) click 'Prepare ZIPs' to create parts and preview, 3) use Test or Bulk send controls.")
