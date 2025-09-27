# streamlit_app.py
import streamlit as st
import pandas as pd
import zipfile, os, io, tempfile, shutil, time, re
from collections import defaultdict
from email.message import EmailMessage
import smtplib
from datetime import datetime
import re

# ---------------- Config ----------------
st.set_page_config(page_title="Aiclex Hallticket Mailer — Final", layout="wide")
st.title("📧 Aiclex Hallticket Mailer — Final (Preview, Test, Send, Compact Downloads)")

# ---------------- Sidebar (settings) ----------------
with st.sidebar:
    st.image("https://aiclex.in/wp-content/uploads/2024/08/aiclex-logo.png", width=140)
    st.header("SMTP & App Settings")
    smtp_host = st.text_input("SMTP Host", value="smtp.gmail.com")
    smtp_port = st.number_input("SMTP Port", value=465)
    protocol = st.selectbox("Protocol", ["SMTPS (SSL)", "SMTP (STARTTLS)"], index=0)
    sender_email = st.text_input("Sender Email", value="info@cruxmanagement.com")
    sender_pass = st.text_input("App Password", value="norx wxop hvsm bvfu", type="password")

    st.markdown("---")
    st.header("Email Templates")
    subject_template = st.text_input("Subject template", value="Hall Tickets — {location} (Part {part}/{total})")
    body_template = st.text_area("Body template", value="Dear Coordinator,\n\nPlease find attached the hall tickets for {location}.\n\nRegards,\nAiclex Technologies", height=140)

    st.markdown("---")
    st.header("Send Controls")
    delay_seconds = st.number_input("Delay between emails (seconds)", value=2.0, step=0.5)
    max_mb = st.number_input("Per-attachment limit (MB)", value=3.0, step=0.5)
    st.markdown("Quick controls:")
    testing_mode_default = st.checkbox("Default: send everything to test email (testing mode)", value=True)
    test_email_default = st.text_input("Default test email", value="info@aiclex.in")
    st.markdown(" ")
    st.caption("Use Preview -> Prepare -> Test -> Send flow. Use Cleanup when finished.")

# ---------------- Helpers ----------------
def extract_zip_recursively(zip_file_like, extract_to):
    """Extract zip (file-like or path) recursively (nested zips)."""
    with zipfile.ZipFile(zip_file_like) as z:
        z.extractall(path=extract_to)
    for root, _, files in os.walk(extract_to):
        for f in files:
            if f.lower().endswith('.zip'):
                nested = os.path.join(root, f)
                nested_dir = os.path.join(root, f"_nested_{os.path.splitext(f)[0]}")
                os.makedirs(nested_dir, exist_ok=True)
                try:
                    with open(nested, 'rb') as nf:
                        extract_zip_recursively(nf, nested_dir)
                except Exception:
                    # skip if nested zip cannot be read
                    continue

def human_bytes(n):
    try:
        n = float(n)
    except:
        return ""
    for unit in ['B','KB','MB','GB','TB']:
        if n < 1024:
            return f"{n:.2f} {unit}"
        n /= 1024
    return f"{n:.2f} PB"

def create_chunked_zips(file_paths, out_dir, base_name, max_bytes):
    """Split list of file_paths into multiple zip parts each <= max_bytes (approx)."""
    os.makedirs(out_dir, exist_ok=True)
    parts = []
    current_files = []
    part_index = 1
    for fp in file_paths:
        current_files.append(fp)
        # test package size
        test_path = os.path.join(out_dir, f"__test_{part_index}.zip")
        with zipfile.ZipFile(test_path, 'w', compression=zipfile.ZIP_DEFLATED) as z:
            for f in current_files:
                z.write(f, arcname=os.path.basename(f))
        size = os.path.getsize(test_path)
        if size <= max_bytes:
            os.remove(test_path)
            continue
        # overflow: remove last, write current part, start new part with last
        last = current_files.pop()
        part_path = os.path.join(out_dir, f"{base_name}_part{part_index}.zip")
        with zipfile.ZipFile(part_path, 'w', compression=zipfile.ZIP_DEFLATED) as z:
            for f in current_files:
                z.write(f, arcname=os.path.basename(f))
        parts.append(part_path)
        part_index += 1
        current_files = [last]
        os.remove(test_path)
    # final part
    if current_files:
        part_path = os.path.join(out_dir, f"{base_name}_part{part_index}.zip")
        with zipfile.ZipFile(part_path, 'w', compression=zipfile.ZIP_DEFLATED) as z:
            for f in current_files:
                z.write(f, arcname=os.path.basename(f))
        parts.append(part_path)
    return parts

def make_download_zip(paths, out_path):
    """Create a single zip containing all given files (temporary)."""
    with zipfile.ZipFile(out_path, 'w', compression=zipfile.ZIP_DEFLATED) as z:
        for p in paths:
            z.write(p, arcname=os.path.basename(p))
    return out_path

# ---------------- Session-state defaults ----------------
if "workdir" not in st.session_state:
    st.session_state.workdir = None
if "pdf_map" not in st.session_state:
    st.session_state.pdf_map = {}
if "grouped" not in st.session_state:
    st.session_state.grouped = {}
if "prepared" not in st.session_state:
    st.session_state.prepared = {}
if "summary_rows" not in st.session_state:
    st.session_state.summary_rows = []
if "cancel_requested" not in st.session_state:
    st.session_state.cancel_requested = False
if "skip_delay" not in st.session_state:
    st.session_state.skip_delay = False
if "status_msg" not in st.session_state:
    st.session_state.status_msg = ""

# small helper to update status placeholder
status_placeholder = st.empty()

# ---------------- Upload area ----------------
st.header("1) Upload files")
col1, col2 = st.columns([2,3])
with col1:
    uploaded_excel = st.file_uploader("Upload Excel (.xlsx or .csv) — columns: Hallticket | Emails | Location", type=["xlsx","csv"], key="excel")
with col2:
    uploaded_zip = st.file_uploader("Upload ZIP (PDFs; nested zips OK)", type=["zip"], key="zip")

if not (uploaded_excel and uploaded_zip):
    st.info("Upload both Excel and ZIP to begin. Use the sidebar to tune SMTP/templates/settings.")
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

# column mapping
cols = list(df.columns)
st.subheader("2) Map columns")
ht_col = st.selectbox("Hallticket column", cols, index=0)
email_col = st.selectbox("Emails column (may include multiple emails separated by comma/semicolon)", cols, index=1 if len(cols)>1 else 0)
loc_col = st.selectbox("Location column", cols, index=2 if len(cols)>2 else 0)

# preview
st.subheader("Data preview (first 10 rows)")
st.dataframe(df[[ht_col, email_col, loc_col]].head(10), width="stretch")

# ---------------- Extract ZIPs into workspace ----------------
if st.session_state.workdir is None:
    st.session_state.workdir = tempfile.mkdtemp(prefix="aiclex_zip_")
workdir = st.session_state.workdir

status_placeholder.info("Extracting ZIP(s) into workspace...")
try:
    bio = io.BytesIO(uploaded_zip.read())
    extract_zip_recursively(bio, workdir)
except Exception as e:
    st.error("ZIP extraction failed: " + str(e))
    st.stop()

# collect pdf files
pdf_map = {}
for root, _, files in os.walk(workdir):
    for f in files:
        if f.lower().endswith(".pdf"):
            pdf_map[f] = os.path.join(root, f)
st.session_state.pdf_map = pdf_map
status_placeholder.success(f"Extracted {len(pdf_map)} PDF files into workspace: {workdir}")

# ---------------- Build mapping table ----------------
mapping_rows = []
for idx, row in df.iterrows():
    hall = str(row[ht_col]).strip() if ht_col in row.index else str(row.iloc[0]).strip()
    raw_emails = str(row[email_col]).strip() if email_col in row.index else str(row.iloc[1]).strip()
    location = str(row[loc_col]).strip() if loc_col in row.index else str(row.iloc[2]).strip()
    matched_fn = ""
    for fn in pdf_map:
        if hall and hall in fn:
            matched_fn = fn
            break
    mapping_rows.append({"Hallticket": hall, "Emails": raw_emails, "Location": location, "MatchedFile": matched_fn or "NOT FOUND"})
map_df = pd.DataFrame(mapping_rows)
st.subheader("3) Mapping Table (Hallticket → Matched PDF)")
st.dataframe(map_df, width="stretch")

# ---------------- Grouping ----------------
grouped = defaultdict(list)
for idx, row in df.iterrows():
    hall = str(row[ht_col]).strip() if ht_col in row.index else str(row.iloc[0]).strip()
    raw_emails = str(row[email_col]).strip() if email_col in row.index else str(row.iloc[1]).strip()
    location = str(row[loc_col]).strip() if loc_col in row.index else str(row.iloc[2]).strip()
    emails = [e.strip().lower() for e in re.split(r"[,;\n]+", raw_emails) if e.strip()]
    email_key = tuple(sorted(emails))
    grouped[(location, email_key)].append(hall)
st.session_state.grouped = grouped

st.subheader("4) Group summary (Location + Recipients)")
summary_list = []
for (loc, email_key), halls in grouped.items():
    matched = 0
    for ht in halls:
        for fn in pdf_map:
            if ht and ht in fn:
                matched += 1
                break
    summary_list.append({"Location": loc, "Recipients": ", ".join(email_key), "Tickets": len(halls), "MatchedPDFs": matched})
summary_df = pd.DataFrame(summary_list)
st.dataframe(summary_df, width="stretch")

# ---------------- Prepare Zips (preview) ----------------
st.markdown("---")
st.subheader("5) Prepare ZIPs (Preview parts before sending)")

prepare_col1, prepare_col2 = st.columns([1,1])
with prepare_col1:
    if st.button("Prepare ZIPs (create parts)"):
        status_placeholder.info("Preparing ZIP parts...")
        st.session_state.cancel_requested = False
        max_bytes = int(max_mb * 1024 * 1024)
        outroot = tempfile.mkdtemp(prefix="aiclex_out_")
        prepared = {}
        summary_rows = []
        groups = list(grouped.items())
        total_groups = len(groups) or 1
        prog_prep = st.progress(0)
        for i, ((loc, email_key), halls) in enumerate(groups, start=1):
            if st.session_state.cancel_requested:
                status_placeholder.warning("Preparation cancelled by user.")
                break
            matched_paths = []
            for ht in halls:
                found = None
                for fn, p in pdf_map.items():
                    if ht and ht in fn:
                        found = p
                        break
                if found:
                    matched_paths.append(found)
            recip_str = ", ".join(email_key)
            if not matched_paths:
                prepared[(loc, recip_str)] = []
                prog_prep.progress(int(i/total_groups*100))
                continue
            out_dir = os.path.join(outroot, f"{loc}_{re.sub(r'[^A-Za-z0-9]', '_', recip_str)[:80]}")
            os.makedirs(out_dir, exist_ok=True)
            parts = create_chunked_zips(matched_paths, out_dir, base_name=loc.replace(" ", "_")[:60], max_bytes=max_bytes)
            prepared[(loc, recip_str)] = parts
            for idx_part, p in enumerate(parts, start=1):
                summary_rows.append({
                    "Location": loc,
                    "Recipients": recip_str,
                    "Part": f"{idx_part}/{len(parts)}",
                    "File": os.path.basename(p),
                    "Size": human_bytes(os.path.getsize(p)),
                    "Path": p
                })
            prog_prep.progress(int(i/total_groups*100))
        st.session_state.prepared = prepared
        st.session_state.summary_rows = summary_rows
        status_placeholder.success("Prepared ZIP parts — preview ready.")

with prepare_col2:
    if st.button("Cancel Preparation"):
        st.session_state.cancel_requested = True
        status_placeholder.warning("Cancel requested. Preparation will stop soon.")

# show preview table if prepared
if st.session_state.summary_rows:
    st.subheader("6) Prepared Parts Preview (select row to download)")
    preview_df = pd.DataFrame(st.session_state.summary_rows)
    st.dataframe(preview_df[["Location","Recipients","Part","File","Size"]], width="stretch")

    # compact download controls:
    options = [f"{i+1}. {r['Location']} — {r['File']} ({r['Part']})" for i,r in enumerate(st.session_state.summary_rows)]
    sel_index = st.selectbox("Select a prepared part to download", options=options, index=0)
    sel_idx = int(sel_index.split(".")[0]) - 1
    sel_row = st.session_state.summary_rows[sel_idx]
    # single download button for selected row
    try:
        with open(sel_row["Path"], "rb") as f:
            st.download_button(
                label=f"⬇️ Download selected: {sel_row['File']}",
                data=f.read(),
                file_name=sel_row["File"],
                key=f"dl_single_{sel_idx}"
            )
    except Exception as e:
        st.warning(f"Cannot open selected file: {e}")

    # Download all prepared parts as single zip
    all_parts_paths = [r["Path"] for r in st.session_state.summary_rows]
    if all_parts_paths:
        tmp_all_zip = os.path.join(tempfile.gettempdir(), f"aiclex_all_parts_{int(time.time())}.zip")
        if st.button("⬇️ Download ALL prepared parts as one ZIP"):
            try:
                make_download_zip(all_parts_paths, tmp_all_zip)
                with open(tmp_all_zip, "rb") as af:
                    st.download_button(label="Download ALL (single ZIP)", data=af.read(), file_name=os.path.basename(tmp_all_zip), key=f"dl_all_{int(time.time())}")
                # remove temp combined zip after offering (not immediate deletion because streamlit needs it for download)
            except Exception as e:
                st.error("Failed to create combined download: " + str(e))

# ---------------- Test send & Bulk send controls ----------------
st.markdown("---")
st.subheader("7) Test send & Bulk send (with progress, skip-delay & cancel)")

col_a, col_b, col_c = st.columns([1,1,1])
with col_a:
    test_email = st.text_input("Test email (overrides recipients when used)", value=test_email_default)
    if st.button("Send Test Email (first available part)"):
        if not st.session_state.prepared:
            st.error("No prepared parts — click Prepare ZIPs first.")
        else:
            status_placeholder.info("Sending test email (first available part)...")
            sent = False
            try:
                if protocol.startswith("SMTPS"):
                    server = smtplib.SMTP_SSL(smtp_host, int(smtp_port), timeout=60)
                else:
                    server = smtplib.SMTP(smtp_host, int(smtp_port), timeout=60)
                    server.starttls()
                server.login(sender_email, sender_pass)
                for (loc, recip_str), parts in st.session_state.prepared.items():
                    if not parts:
                        continue
                    first = parts[0]
                    msg = EmailMessage()
                    msg["From"] = sender_email
                    msg["To"] = test_email
                    try:
                        subj = subject_template.format(location=loc, part=1, total=len(parts))
                    except:
                        subj = f"{loc} part 1/{len(parts)}"
                    msg["Subject"] = f"[TEST] {subj}"
                    try:
                        body_txt = body_template.format(location=loc, part=1, total=len(parts))
                    except:
                        body_txt = f"Test: attached {os.path.basename(first)}"
                    msg.set_content(body_txt + "\n\n(This is a TEST email — only first part attached.)")
                    with open(first, "rb") as af:
                        msg.add_attachment(af.read(), maintype="application", subtype="zip", filename=os.path.basename(first))
                    server.send_message(msg)
                    sent = True
                    status_placeholder.success(f"Test email sent to {test_email} with {os.path.basename(first)}")
                    break
                try:
                    server.quit()
                except:
                    pass
                if not sent:
                    st.warning("No parts available to test send.")
            except Exception as e:
                st.error("Test send failed: " + str(e))

with col_b:
    skip_delay_chk = st.checkbox("Skip delay during sending (push immediately)", value=False, key="ui_skip_delay")
    if st.button("Cancel ongoing operation"):
        st.session_state.cancel_requested = True
        status_placeholder.warning("Cancel requested — sending will stop soon.")

with col_c:
    if st.button("Send ALL Prepared Parts (Bulk)"):
        if not st.session_state.prepared:
            st.error("No prepared parts — click Prepare ZIPs first.")
        else:
            st.session_state.cancel_requested = False
            total_parts = sum(len(p) for p in st.session_state.prepared.values())
            if total_parts == 0:
                st.warning("No parts to send.")
            else:
                status_placeholder.info("Starting bulk send...")
                sent_count = 0
                logs = []
                prog_send = st.progress(0)
                try:
                    if protocol.startswith("SMTPS"):
                        server = smtplib.SMTP_SSL(smtp_host, int(smtp_port), timeout=60)
                    else:
                        server = smtplib.SMTP(smtp_host, int(smtp_port), timeout=60)
                        server.starttls()
                    server.login(sender_email, sender_pass)
                    for (loc, recipients_str), parts in st.session_state.prepared.items():
                        if st.session_state.cancel_requested:
                            status_placeholder.warning("Bulk send cancelled by user.")
                            break
                        if not parts:
                            logs.append({"Location": loc, "Recipients": recipients_str, "Part": "", "File": "", "Status": "No parts"})
                            continue
                        for idx_part, ppath in enumerate(parts, start=1):
                            if st.session_state.cancel_requested:
                                break
                            msg = EmailMessage()
                            msg["From"] = sender_email
                            target_to = test_email if testing_mode_default else recipients_str
                            msg["To"] = target_to
                            try:
                                subject_line = subject_template.format(location=loc, part=idx_part, total=len(parts))
                            except:
                                subject_line = f"{loc} part {idx_part}/{len(parts)}"
                            msg["Subject"] = subject_line
                            try:
                                body_txt = body_template.format(location=loc, part=idx_part, total=len(parts))
                            except:
                                body_txt = f"Please find attached part {idx_part} for {loc}."
                            msg.set_content(body_txt)
                            with open(ppath, "rb") as af:
                                msg.add_attachment(af.read(), maintype="application", subtype="zip", filename=os.path.basename(ppath))
                            try:
                                server.send_message(msg)
                                logs.append({"Location": loc, "Recipients": target_to, "Part": f"{idx_part}/{len(parts)}", "File": os.path.basename(ppath), "Status": "Sent", "Time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
                            except Exception as e:
                                logs.append({"Location": loc, "Recipients": target_to, "Part": f"{idx_part}/{len(parts)}", "File": os.path.basename(ppath), "Status": f"Failed: {e}"})
                            sent_count += 1
                            prog_send.progress(int(sent_count / total_parts * 100))
                            # delay control
                            if not skip_delay_chk:
                                time.sleep(float(delay_seconds))
                    try:
                        server.quit()
                    except:
                        pass
                    status_placeholder.success("Bulk send complete (or stopped).")
                    st.subheader("Sending logs")
                    st.dataframe(pd.DataFrame(logs), width="stretch")
                except Exception as e:
                    st.error("Bulk send failed: " + str(e))

# ---------------- Cleanup ----------------
st.markdown("---")
if st.button("🧹 Cleanup workspace (delete extracted files & prepared parts)"):
    try:
        wd = st.session_state.get("workdir")
        if wd and os.path.exists(wd):
            shutil.rmtree(wd)
        # remove prepared parts directories too
        for (_, _), parts in st.session_state.get("prepared", {}).items():
            for p in parts:
                try:
                    parent = os.path.dirname(p)
                    if parent and os.path.exists(parent):
                        shutil.rmtree(parent)
                except:
                    pass
        st.session_state.workdir = None
        st.session_state.pdf_map = {}
        st.session_state.grouped = {}
        st.session_state.prepared = {}
        st.session_state.summary_rows = []
        st.session_state.cancel_requested = False
        status_placeholder.info("Workspace cleaned.")
    except Exception as e:
        st.error("Cleanup failed: " + str(e))
