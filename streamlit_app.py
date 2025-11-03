# streamlit_app.py
"""
Aiclex Hallticket Mailer — Final with persistent SQLite send log + Resume
Drop-in streamlit app. Run: streamlit run streamlit_app.py
Dependencies: streamlit, pandas
"""

import streamlit as st
import pandas as pd
import zipfile, os, io, tempfile, shutil, time, re, sqlite3, json
from collections import defaultdict
from email.message import EmailMessage
import smtplib
from datetime import datetime

# ---------------- Config ----------------
APP_DB = os.path.join(os.getcwd(), "send_logs.db")  # persistent DB in app working dir
LOG_TABLE = "email_sends"

st.set_page_config(page_title="Aiclex Mailer — Safe with Resume", layout="wide")
st.title("🛡️ Aiclex Hallticket Mailer — Safe + Resume")

# ---------------- DB helpers ----------------
def init_db():
    conn = sqlite3.connect(APP_DB, timeout=30)
    cur = conn.cursor()
    cur.execute(f"""
    CREATE TABLE IF NOT EXISTS {LOG_TABLE} (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp TEXT,
        location TEXT,
        recipients TEXT,
        halltickets TEXT,
        part TEXT,
        file TEXT,
        files_in_part INTEGER,
        status TEXT,
        error TEXT
    )
    """)
    conn.commit()
    return conn

def append_log(conn, row):
    cur = conn.cursor()
    cur.execute(f"""
      INSERT INTO {LOG_TABLE} (timestamp, location, recipients, halltickets, part, file, files_in_part, status, error)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        row.get("location",""),
        row.get("recipients",""),
        json.dumps(row.get("halltickets",[]), ensure_ascii=False),
        row.get("part",""),
        row.get("file",""),
        int(row.get("files_in_part",0)),
        row.get("status",""),
        str(row.get("error",""))
    ))
    conn.commit()

def update_log_status(conn, log_id, status, error=""):
    cur = conn.cursor()
    cur.execute(f"UPDATE {LOG_TABLE} SET status=?, error=? WHERE id=?", (status, str(error), log_id))
    conn.commit()

def fetch_stats(conn):
    cur = conn.cursor()
    cur.execute(f"SELECT COUNT(*) FROM {LOG_TABLE}")
    total = cur.fetchone()[0]
    cur.execute(f"SELECT COUNT(*) FROM {LOG_TABLE} WHERE status='Sent'")
    sent = cur.fetchone()[0]
    cur.execute(f"SELECT COUNT(*) FROM {LOG_TABLE} WHERE status!='Sent'")
    pending = cur.fetchone()[0]
    cur.execute(f"SELECT COUNT(*) FROM {LOG_TABLE} WHERE status='Failed'")
    failed = cur.fetchone()[0]
    return {"total": total, "sent": sent, "pending": pending, "failed": failed}

def fetch_pending_rows(conn):
    cur = conn.cursor()
    cur.execute(f"SELECT id, location, recipients, halltickets, part, file, files_in_part, status FROM {LOG_TABLE} WHERE status!='Sent' ORDER BY id")
    rows = cur.fetchall()
    res = []
    for r in rows:
        res.append({
            "id": r[0],
            "location": r[1],
            "recipients": r[2],
            "halltickets": json.loads(r[3]) if r[3] else [],
            "part": r[4],
            "file": r[5],
            "files_in_part": r[6],
            "status": r[7]
        })
    return res

def clear_pending(conn):
    cur = conn.cursor()
    cur.execute(f"DELETE FROM {LOG_TABLE}")
    conn.commit()

# initialize DB connection
conn = init_db()

# ---------------- Sidebar / Settings ----------------
with st.sidebar:

    st.header("Email templates & sending")
    subject_template = st.text_input("Subject template", value="Hall Tickets — {location} (Part {part}/{total})")
    body_template = st.text_area("Body template", value="Dear Coordinator,\n\nPlease find attached the hall tickets for {location}.\n\nRegards,\nAiclex Technologies", height=140)

    st.markdown("---")
    delay_seconds = st.number_input("Delay between emails (sec)", value=2.0, step=0.5)
    max_mb = st.number_input("Per-attachment limit (MB)", value=3.0, step=0.5)
    st.markdown("Use small delay (1-3s) for deliverability; very long delays can cause SMTP to drop.")
    st.markdown("---")
    st.header("Testing")
    testing_mode_default = st.checkbox("Default: testing mode (override recipients)", value=True)
    test_email_default = st.text_input("Default test email", value=st.secrets.get("smtp_credentials", {}).get("default_test_email", ""))
    # --- START: SMTP Config ---
        st.subheader("SMTP Credentials")
        smtp_creds = st.secrets.get("smtp_credentials", {})
        
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            smtp_host = st.text_input("SMTP Host", value=smtp_creds.get("host", ""))
            sender_email = st.text_input("Sender Email", value=smtp_creds.get("email", ""))
        with col_s2:
            smtp_port = st.text_input("SMTP Port", value=smtp_creds.get("port", "587"))
            sender_pass = st.text_input("Sender Password", value=smtp_creds.get("password", ""), type="password")

        protocol = st.selectbox("Protocol", ["STARTTLS", "SMTPS"], index=0 if smtp_creds.get("protocol", "STARTTLS") == "STARTTLS" else 1)
        # --- END: SMTP Config ---


# show DB stats
stats = fetch_stats(conn)
st.sidebar.markdown(f"**Send log stats** \nTotal attempts: {stats['total']}  \nSent: {stats['sent']}  \nPending/Failed: {stats['pending']}  \nFailed: {stats['failed']}")

# ---------------- Helpers for ZIP / matching ----------------
def extract_zip_recursively(zip_file_like, extract_to):
    if hasattr(zip_file_like, "read"):
        zf = zipfile.ZipFile(zip_file_like)
    else:
        zf = zipfile.ZipFile(zip_file_like, "r")
    try:
        zf.extractall(path=extract_to)
    finally:
        zf.close()
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

def human_bytes(n):
    try: n = float(n)
    except: return ""
    for unit in ['B','KB','MB','GB','TB']:
        if n < 1024:
            return f"{n:.2f} {unit}"
        n /= 1024
    return f"{n:.2f} PB"

def create_chunked_zips_with_counts(file_paths, out_dir, base_name, max_bytes):
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
        last = current_files.pop()
        part_path = os.path.join(out_dir, f"{base_name}_part{part_index}.zip")
        with zipfile.ZipFile(part_path, 'w', compression=zipfile.ZIP_DEFLATED) as z:
            for f in current_files:
                z.write(f, arcname=os.path.basename(f))
        with zipfile.ZipFile(part_path, 'r') as zc:
            names = zc.namelist()
        parts.append({"path": part_path, "files": names, "size": os.path.getsize(part_path)})
        part_index += 1
        current_files = [last]
        os.remove(test_path)
    if current_files:
        part_path = os.path.join(out_dir, f"{base_name}_part{part_index}.zip")
        with zipfile.ZipFile(part_path, 'w', compression=zipfile.ZIP_DEFLATED) as z:
            for f in current_files:
                z.write(f, arcname=os.path.basename(f))
        with zipfile.ZipFile(part_path, 'r') as zc:
            names = zc.namelist()
        parts.append({"path": part_path, "files": names, "size": os.path.getsize(part_path)})
    return parts

def make_download_zip(paths, out_path):
    with zipfile.ZipFile(out_path, 'w', compression=zipfile.ZIP_DEFLATED) as z:
        for p in paths:
            if os.path.exists(p):
                z.write(p, arcname=os.path.basename(p))
    return out_path

# ---------------- Session state defaults ----------------
if "workdir" not in st.session_state: st.session_state.workdir = None
if "pdf_map" not in st.session_state: st.session_state.pdf_map = {}
if "grouped" not in st.session_state: st.session_state.grouped = {}
if "prepared" not in st.session_state: st.session_state.prepared = {}
if "summary_rows" not in st.session_state: st.session_state.summary_rows = []
if "cancel_requested" not in st.session_state: st.session_state.cancel_requested = False
if "skip_delay" not in st.session_state: st.session_state.skip_delay = False
if "verified" not in st.session_state: st.session_state.verified = False

status_ph = st.empty()

# ---------------- Upload UI ----------------
st.header("1) Upload Excel & ZIP")
col1, col2 = st.columns([2,3])
with col1:
    uploaded_excel = st.file_uploader("Upload Excel (.xlsx or .csv) — contains Hallticket, Emails, Location", type=["xlsx","csv"], key="upl_excel")
with col2:
    uploaded_zip = st.file_uploader("Upload ZIP (PDFs; nested zips OK)", type=["zip"], key="upl_zip")

if not (uploaded_excel and uploaded_zip):
    st.info("Upload both Excel and ZIP to begin (mapping, verify, prepare, send).")
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

cols = list(df.columns)
st.subheader("2) Map columns")
ht_col = st.selectbox("Hallticket column", cols, index=0)
email_col = st.selectbox("Emails column (may contain multiple separated by comma/semicolon)", cols, index=1 if len(cols)>1 else 0)
loc_col = st.selectbox("Location column", cols, index=2 if len(cols)>2 else 0)
st.subheader("Data preview (first 8 rows)")
st.dataframe(df[[ht_col, email_col, loc_col]].head(8), width="stretch")

# ---------------- Extract ZIP ----------------
if st.session_state.workdir is None:
    st.session_state.workdir = tempfile.mkdtemp(prefix="aiclex_zip_")
workdir = st.session_state.workdir
status_ph.info("Extracting uploaded ZIP into workspace...")
try:
    bio = io.BytesIO(uploaded_zip.read())
    extract_zip_recursively(bio, workdir)
except Exception as e:
    st.error("ZIP extraction failed: " + str(e))
    st.stop()

pdf_map = {}
for root, _, files in os.walk(workdir):
    for f in files:
        if f.lower().endswith(".pdf"):
            pdf_map[f] = os.path.join(root, f)
st.session_state.pdf_map = pdf_map
status_ph.success(f"Extracted {len(pdf_map)} PDFs into workspace: {workdir}")

# ---------------- Mapping Excel -> PDF ----------------
mapping_rows = []
excel_halls = []
for idx, row in df.iterrows():
    hall = str(row[ht_col]).strip() if ht_col in row.index else str(row.iloc[0]).strip()
    raw_emails = str(row[email_col]).strip() if email_col in row.index else str(row.iloc[1]).strip()
    location = str(row[loc_col]).strip() if loc_col in row.index else str(row.iloc[2]).strip()
    excel_halls.append(hall)
    matched_files = []
    if hall:
        hall_low = hall.lower()
        for fn, path in pdf_map.items():
            fn_low = fn.lower()
            if fn_low.endswith(f"{hall_low}.pdf") or re.search(rf"[^0-9]{re.escape(hall_low)}[^0-9]", fn_low) or hall_low in fn_low:
                matched_files.append(fn)
    matched_files = sorted(set(matched_files))
    mapping_rows.append({
        "Hallticket": hall,
        "Emails": raw_emails,
        "Location": location,
        "MatchedCount": len(matched_files),
        "MatchedFiles": "; ".join(matched_files)
    })
map_df = pd.DataFrame(mapping_rows)
st.subheader("3) Mapping Table (Excel → PDF)")
st.markdown("Download `mapping_check.csv` and verify.")
st.download_button("⬇️ mapping_check.csv", data=map_df.to_csv(index=False), file_name="mapping_check.csv", mime="text/csv", key="dl_map_check")
st.dataframe(map_df, width="stretch")

# ---------------- Reverse mapping PDF -> Excel ----------------
pdf_reverse_rows = []
excel_set = set([str(x).strip().lower() for x in excel_halls if str(x).strip() != ""])
for fn, p in pdf_map.items():
    fn_low = fn.lower()
    digits = re.findall(r"\d{4,20}", fn_low)
    matched_hall = ""
    for d in digits:
        if d in excel_set:
            matched_hall = d
            break
    if not matched_hall and digits:
        last = digits[-1]
        if last in excel_set:
            matched_hall = last
    pdf_reverse_rows.append({"PDFFile": fn, "DetectedHallticket": matched_hall or "", "MatchedInExcel": bool(matched_hall)})
pdf_rev_df = pd.DataFrame(pdf_reverse_rows)
st.subheader("4) Reverse mapping (PDF → Excel detect)")
st.markdown("Download `extra_in_zip.csv` (PDFs not matched to any Excel hallticket).")
extra_csv = pdf_rev_df[pdf_rev_df["MatchedInExcel"]==False].to_csv(index=False)
st.download_button("⬇️ extra_in_zip.csv", data=extra_csv, file_name="extra_in_zip.csv", mime="text/csv", key="dl_extra")
st.dataframe(pdf_rev_df, width="stretch")

# missing (Excel halltickets with zero matches)
missing_df = map_df[map_df["MatchedCount"] == 0][["Hallticket","Emails","Location"]]
st.download_button("⬇️ missing_in_zip.csv", data=missing_df.to_csv(index=False), file_name="missing_in_zip.csv", mime="text/csv", key="dl_missing")
st.markdown("---")

# verification gate
st.subheader("⚠️ Verification required")
st.markdown("Please review the three CSVs above. After manual verification, check the box to enable Prepare & Send.")
st.session_state.verified = st.checkbox("I have reviewed mapping_check.csv, missing_in_zip.csv, extra_in_zip.csv and confirm accuracy", value=False, key="verify_final")
if not st.session_state.verified:
    st.warning("Prepare & Send disabled until you verify mappings.")
    st.stop()

# ---------------- Grouping (Location + row-level recipients) ----------------
grouped = defaultdict(list)
for idx, row in df.iterrows():
    hall = str(row[ht_col]).strip() if ht_col in row.index else str(row.iloc[0]).strip()
    raw_emails = str(row[email_col]).strip() if email_col in row.index else str(row.iloc[1]).strip()
    location = str(row[loc_col]).strip() if loc_col in row.index else str(row.iloc[2]).strip()
    emails = [e.strip().lower() for e in re.split(r"[,;\n]+", raw_emails) if e.strip()]
    recip_key = tuple(sorted(emails))
    grouped[(location, recip_key)].append(hall)
st.session_state.grouped = grouped

st.subheader("5) Group summary (Location + Recipients)")
summary_rows = []
for (loc, recip_key), halls in grouped.items():
    matched_count = sum(1 for ht in halls for fn in pdf_map if ht and ht in fn)
    summary_rows.append({"Location": loc, "Recipients": ", ".join(recip_key), "Tickets": len(halls), "MatchedPDFs": matched_count})
summary_df = pd.DataFrame(summary_rows)
st.dataframe(summary_df, width="stretch")

# ---------------- Prepare ZIPs ----------------
st.markdown("---")
st.subheader("6) Prepare ZIPs (create parts with counts & preview)")
prep_col1, prep_col2 = st.columns([1,1])
with prep_col1:
    if st.button("Prepare ZIPs (create parts)"):
        st.session_state.cancel_requested = False
        status_ph.info("Preparing ZIP parts...")
        max_bytes = int(max_mb * 1024 * 1024)
        outroot = tempfile.mkdtemp(prefix="aiclex_out_")
        prepared = {}
        summary_rows = []
        groups = list(grouped.items())
        total = max(1, len(groups))
        prog = st.progress(0)
        for i, ((loc, recip_key), halls) in enumerate(groups, start=1):
            if st.session_state.cancel_requested:
                status_ph.warning("Preparation cancelled.")
                break
            matched_paths = []
            for ht in halls:
                for fn, p in pdf_map.items():
                    if ht and ht in fn:
                        matched_paths.append(p)
            recip_str = ", ".join(recip_key)
            if not matched_paths:
                prepared[(loc, recip_str)] = []
                prog.progress(int(i/total*100))
                continue
            out_dir = os.path.join(outroot, f"{loc}_{re.sub(r'[^A-Za-z0-9]', '_', recip_str)[:80]}")
            os.makedirs(out_dir, exist_ok=True)
            parts = create_chunked_zips_with_counts(matched_paths, out_dir, base_name=loc.replace(" ", "_")[:60], max_bytes=max_bytes)
            prepared[(loc, recip_str)] = parts
            total_files_in_group = sum(len(pinfo["files"]) for pinfo in parts)
            for idx_part, pinfo in enumerate(parts, start=1):
                summary_rows.append({
                    "Location": loc,
                    "Recipients": recip_str,
                    "Part": f"{idx_part}/{len(parts)}",
                    "File": os.path.basename(pinfo["path"]),
                    "Size": human_bytes(pinfo["size"]),
                    "FilesInPart": len(pinfo["files"]),
                    "TotalFilesInGroup": total_files_in_group,
                    "Path": pinfo["path"]
                })
            prog.progress(int(i/total*100))
        st.session_state.prepared = prepared
        st.session_state.summary_rows = summary_rows
        status_ph.success("Prepared ZIP parts created — preview ready.")
with prep_col2:
    if st.button("Cancel Preparation"):
        st.session_state.cancel_requested = True
        status_ph.warning("Cancel requested — preparation will stop soon.")

# preview prepared parts
if st.session_state.get("summary_rows"):
    st.subheader("7) Prepared Parts Preview")
    prep_df = pd.DataFrame(st.session_state["summary_rows"])
    st.download_button("⬇️ prepared_summary.csv", data=prep_df.to_csv(index=False), file_name="prepared_summary.csv", mime="text/csv", key="dl_prep")
    st.dataframe(prep_df[["Location","Recipients","Part","File","Size","FilesInPart","TotalFilesInGroup"]], width="stretch")

    # compact download: select row
    opts = [f"{i+1}. {r['Location']} — {r['File']} ({r['Part']}) [{r['FilesInPart']} files]" for i,r in enumerate(st.session_state["summary_rows"])]
    sel = st.selectbox("Select a prepared part to download", opts, index=0, key="sel_part_ui")
    sel_idx = int(sel.split(".")[0]) - 1
    sel_row = st.session_state["summary_rows"][sel_idx]
    try:
        with open(sel_row["Path"], "rb") as f:
            st.download_button(label=f"⬇️ Download selected part", data=f.read(), file_name=sel_row["File"], key=f"dl_sel_{sel_idx}")
    except Exception as e:
        st.warning(f"Cannot open selected prepared part: {e}")

    # download all combined
    all_paths = [r["Path"] for r in st.session_state["summary_rows"] if os.path.exists(r["Path"])]
    if all_paths:
        if st.button("⬇️ Download ALL prepared parts as single ZIP"):
            tmp_all = os.path.join(tempfile.gettempdir(), f"aiclex_all_parts_{int(time.time())}.zip")
            try:
                make_download_zip(all_paths, tmp_all)
                with open(tmp_all, "rb") as af:
                    st.download_button(label="Download combined ZIP", data=af.read(), file_name=os.path.basename(tmp_all), key=f"dl_all_{int(time.time())}")
            except Exception as e:
                st.error("Failed to create combined download: " + str(e))

# ---------------- Test & Bulk send with DB logging ----------------
st.markdown("---")
st.subheader("8) Test send, Bulk Send & Resume (persistent log)")

col_test, col_opts, col_send = st.columns([1,1,1])
with col_test:
    test_email = st.text_input("Test email (overrides recipients)", value=test_email_default, key="test_email_input")
    if st.button("Send Test Email (first available part)"):
        if not st.session_state.get("prepared"):
            st.error("No prepared parts — click Prepare ZIPs first.")
        else:
            status_ph.info("Sending test email (first prepared part)...")
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
                    first = parts[0]["path"]
                    msg = EmailMessage()
                    msg["From"] = sender_email
                    msg["To"] = test_email
                    try:
                        subj = subject_template.format(location=loc, part=1, total=len(parts))
                    except:
                        subj = f"{loc} part 1/{len(parts)}"
                    msg["Subject"] = "[TEST] " + subj
                    try:
                        body_txt = body_template.format(location=loc, part=1, total=len(parts))
                    except:
                        body_txt = f"Test: attached {os.path.basename(first)}"
                    msg.set_content(body_txt + "\n\n(This is a TEST email — only first part attached.)")
                    with open(first, "rb") as af:
                        msg.add_attachment(af.read(), maintype="application", subtype="zip", filename=os.path.basename(first))
                    server.send_message(msg)
                    # log test send as Sent in DB (so resume won't re-send)
                    append_log(conn, {"location": loc, "recipients": test_email, "halltickets": [], "part": "1/1", "file": os.path.basename(first), "files_in_part": len(parts[0]["files"]) if parts else 0, "status": "Sent", "error": ""})
                    sent = True
                    status_ph.success(f"Test email sent to {test_email} with {os.path.basename(first)}")
                    break
                try:
                    server.quit()
                except:
                    pass
                if not sent:
                    st.warning("No parts available to test send.")
            except Exception as e:
                st.error("Test send failed: " + str(e))

with col_opts:
    skip_delay_chk = st.checkbox("Skip delay during sending (push immediately)", value=False, key="skip_delay_send")
    if st.button("Cancel ongoing operation"):
        st.session_state.cancel_requested = True
        status_ph.warning("Cancel requested — operation will stop shortly.")

    # Resume pending sends (DB-based)
    if st.button("Resume Pending Sends (DB)"):
        pending = fetch_pending_rows(conn)
        if not pending:
            st.info("No pending entries to resume.")
        else:
            status_ph.info(f"Resuming {len(pending)} pending sends...")
            prog = st.progress(0)
            total_pending = len(pending)
            sent_count = 0
            try:
                if protocol.startswith("SMTPS"):
                    server = smtplib.SMTP_SSL(smtp_host, int(smtp_port), timeout=60)
                else:
                    server = smtplib.SMTP(smtp_host, int(smtp_port), timeout=60)
                    server.starttls()
                server.login(sender_email, sender_pass)
                RECONNECT_EVERY = 100
                rc = 0
                for i, item in enumerate(pending, start=1):
                    if st.session_state.cancel_requested:
                        status_ph.warning("Resume cancelled by user.")
                        break
                    # build email
                    msg = EmailMessage()
                    msg["From"] = sender_email
                    # recipients stored as comma-separated; allow override if testing default on
                    target_to = test_email if testing_mode_default else item["recipients"]
                    msg["To"] = target_to
                    try:
                        msg["Subject"] = subject_template.format(location=item["location"], part=item["part"].split("/")[0], total=item["part"].split("/")[-1])
                    except:
                        msg["Subject"] = f"{item['location']} {item['part']}"
                    msg.set_content(f"Resuming send for {item['location']} — part {item['part']}")
                    # attach file by full path: find prepared summary row matching file
                    fname = item["file"]
                    # locate file path in summary_rows
                    ppath = None
                    for r in st.session_state.get("summary_rows", []):
                        if r["File"] == fname:
                            ppath = r["Path"]
                            break
                    if not ppath or not os.path.exists(ppath):
                        # mark failed in DB
                        append_log(conn, {"location": item["location"], "recipients": item["recipients"], "halltickets": item.get("halltickets",[]), "part": item["part"], "file": fname, "files_in_part": item.get("files_in_part",0), "status": "Failed", "error": "Prepared file missing on server"})
                        continue
                    with open(ppath, "rb") as af:
                        msg.add_attachment(af.read(), maintype="application", subtype="zip", filename=os.path.basename(ppath))
                    try:
                        server.send_message(msg)
                        append_log(conn, {"location": item["location"], "recipients": target_to, "halltickets": item.get("halltickets",[]), "part": item["part"], "file": fname, "files_in_part": item.get("files_in_part",0), "status": "Sent", "error": ""})
                    except Exception as e:
                        append_log(conn, {"location": item["location"], "recipients": target_to, "halltickets": item.get("halltickets",[]), "part": item["part"], "file": fname, "files_in_part": item.get("files_in_part",0), "status": "Failed", "error": str(e)})
                    sent_count += 1
                    rc += 1
                    prog.progress(int(i/total_pending*100))
                    if rc >= RECONNECT_EVERY:
                        try: server.quit()
                        except: pass
                        if protocol.startswith("SMTPS"):
                            server = smtplib.SMTP_SSL(smtp_host, int(smtp_port), timeout=60)
                        else:
                            server = smtplib.SMTP(smtp_host, int(smtp_port), timeout=60)
                            server.starttls()
                        server.login(sender_email, sender_pass)
                        rc = 0
                    if not skip_delay_chk:
                        time.sleep(float(delay_seconds))
                try: server.quit()
                except: pass
                status_ph.success("Resume finished (see DB logs).")
            except Exception as e:
                st.error("Resume failed: " + str(e))

with col_send:
    if st.button("Send ALL Prepared Parts (Bulk)"):
        if not st.session_state.get("prepared"):
            st.error("No prepared parts — Prepare ZIPs first.")
        else:
            st.session_state.cancel_requested = False
            total_parts = sum(len(parts) for parts in st.session_state.prepared.values())
            if total_parts == 0:
                st.warning("No parts to send.")
            else:
                status_ph.info("Starting bulk send...")
                sent_count = 0
                logs = []
                prog = st.progress(0)
                try:
                    if protocol.startswith("SMTPS"):
                        server = smtplib.SMTP_SSL(smtp_host, int(smtp_port), timeout=60)
                    else:
                        server = smtplib.SMTP(smtp_host, int(smtp_port), timeout=60)
                        server.starttls()
                    server.login(sender_email, sender_pass)
                    RECONNECT_EVERY = 100
                    rc = 0
                    for (loc, recip_str), parts in st.session_state.prepared.items():
                        if st.session_state.cancel_requested:
                            status_ph.warning("Bulk send cancelled by user.")
                            break
                        if not parts:
                            logs.append({"Location": loc, "Recipients": recip_str, "Part": "", "File": "", "Status": "No parts"})
                            continue
                        for idx_part, pinfo in enumerate(parts, start=1):
                            if st.session_state.cancel_requested:
                                break
                            # determine recipient target (use testing mode if set)
                            target_to = test_email if testing_mode_default else recip_str
                            msg = EmailMessage()
                            msg["From"] = sender_email
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
                            with open(pinfo["path"], "rb") as af:
                                msg.add_attachment(af.read(), maintype="application", subtype="zip", filename=os.path.basename(pinfo["path"]))
                            # Log BEFORE sending as Pending (so resume can pick it up if crash occurs)
                            append_log(conn, {"location": loc, "recipients": recip_str, "halltickets": [], "part": f"{idx_part}/{len(parts)}", "file": os.path.basename(pinfo["path"]), "files_in_part": len(pinfo["files"]), "status": "Pending", "error": ""})
                            try:
                                server.send_message(msg)
                                # update log as Sent by inserting new row (keeps history)
                                append_log(conn, {"location": loc, "recipients": target_to, "halltickets": [], "part": f"{idx_part}/{len(parts)}", "file": os.path.basename(pinfo["path"]), "files_in_part": len(pinfo["files"]), "status": "Sent", "error": ""})
                                logs.append({"Location": loc, "Recipients": target_to, "Part": f"{idx_part}/{len(parts)}", "File": os.path.basename(pinfo["path"]), "FilesInPart": len(pinfo["files"]), "Status": "Sent"})
                            except Exception as e:
                                append_log(conn, {"location": loc, "recipients": target_to, "halltickets": [], "part": f"{idx_part}/{len(parts)}", "file": os.path.basename(pinfo["path"]), "files_in_part": len(pinfo["files"]), "status": "Failed", "error": str(e)})
                                logs.append({"Location": loc, "Recipients": target_to, "Part": f"{idx_part}/{len(parts)}", "File": os.path.basename(pinfo["path"]), "FilesInPart": len(pinfo["files"]), "Status": f"Failed: {e}"})
                            sent_count += 1
                            rc += 1
                            prog.progress(int(sent_count / total_parts * 100))
                            if rc >= RECONNECT_EVERY:
                                try: server.quit()
                                except: pass
                                if protocol.startswith("SMTPS"):
                                    server = smtplib.SMTP_SSL(smtp_host, int(smtp_port), timeout=60)
                                else:
                                    server = smtplib.SMTP(smtp_host, int(smtp_port), timeout=60)
                                    server.starttls()
                                server.login(sender_email, sender_pass)
                                rc = 0
                            if not skip_delay_chk:
                                time.sleep(float(delay_seconds))
                    try: server.quit()
                    except: pass
                    status_ph.success("Bulk send complete (or stopped). See logs below.")
                    st.subheader("Immediate send log (recent attempts)")
                    st.dataframe(pd.DataFrame(logs), width="stretch")
                except Exception as e:
                    st.error("Bulk send failed: " + str(e))

# ---------------- Resume info & manual DB controls ----------------
st.markdown("---")
st.subheader("9) Persistent send log (resume / audit)")
st.markdown("Use these buttons to inspect, export or reset the persistent send log (useful after crash).")
col_a, col_b, col_c = st.columns([1,1,1])
with col_a:
    if st.button("Show send log (last 200 rows)"):
        cur = conn.cursor()
        cur.execute(f"SELECT id, timestamp, location, recipients, part, file, files_in_part, status, error FROM {LOG_TABLE} ORDER BY id DESC LIMIT 200")
        rows = cur.fetchall()
        df_logs = pd.DataFrame(rows, columns=["id","timestamp","location","recipients","part","file","files_in_part","status","error"])
        st.dataframe(df_logs, width="stretch")
with col_b:
    if st.button("Download full send_log.csv"):
        cur = conn.cursor()
        cur.execute(f"SELECT id, timestamp, location, recipients, part, file, files_in_part, status, error FROM {LOG_TABLE} ORDER BY id")
        rows = cur.fetchall()
        df_logs = pd.DataFrame(rows, columns=["id","timestamp","location","recipients","part","file","files_in_part","status","error"])
        st.download_button("⬇️ Download CSV (send_log.csv)", data=df_logs.to_csv(index=False), file_name="send_log.csv", mime="text/csv", key="dl_sendlog")
with col_c:
    if st.button("Start New Batch (CLEAR send_log)"):
        if st.confirm("Are you sure? This will delete the send_log and cannot be undone. Use Resume if you want to re-attempt pending sends."):
            clear_pending(conn)
            st.success("send_log cleared. Starting fresh.")

# ---------------- Cleanup workspace ----------------
st.markdown("---")
if st.button("🧹 Cleanup workspace (delete extracted & prepared files)"):
    try:
        wd = st.session_state.get("workdir")
        if wd and os.path.exists(wd):
            shutil.rmtree(wd)
        for key, parts in st.session_state.get("prepared", {}).items():
            for p in parts:
                try:
                    parent = os.path.dirname(p["path"]) if isinstance(p, dict) else os.path.dirname(p)
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
        st.session_state.verified = False
        status_ph.info("Workspace cleaned and verification reset.")
    except Exception as e:
        st.error("Cleanup failed: " + str(e))
