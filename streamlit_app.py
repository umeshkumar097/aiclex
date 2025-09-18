# streamlit_app_ocr.py
# Aiclex — Final Hall Ticket Mailer with OCR & Result Formation
# - Two modes: "Email Send" and "Result Formation"
# - OCR-enabled PDF text extraction (pdf2image + pytesseract) when available
# - Nested ZIP extraction, per-location grouping, recipient edit, optional compression,
#   chunking to per-attachment limit, preview, delay between sends
# - Result Formation: immediate CSV mapping on upload (for quick check) + full Excel (Summary + Details)
#
# Run:
# streamlit run streamlit_app_ocr.py

import streamlit as st
import pandas as pd
import zipfile, io, os, tempfile, shutil, time, re
from pathlib import Path
from collections import defaultdict
from datetime import datetime
import smtplib
from email.message import EmailMessage

# optional libs for OCR & compression
try:
    from pdf2image import convert_from_path
    PDF2IMAGE_AVAILABLE = True
except Exception:
    PDF2IMAGE_AVAILABLE = False

try:
    import pytesseract
    PYTESSERACT_AVAILABLE = True
except Exception:
    PYTESSERACT_AVAILABLE = False

try:
    import pikepdf
    PIKEPDF_AVAILABLE = True
except Exception:
    PIKEPDF_AVAILABLE = False

st.set_page_config(page_title="Aiclex Mailer — OCR & Results", layout="wide")

# ---------------- Helpers ----------------
def human_bytes(n):
    try:
        n = float(n)
    except:
        return ""
    for unit in ['B','KB','MB','GB','TB']:
        if n < 1024:
            return f"{n:.2f} {unit}"
        n /= 1024
    return f"{n:.2f} TB"

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

def compress_pdf_pikepdf(src, dst):
    """Try to compress with pikepdf; return True on success (dst created)."""
    if not PIKEPDF_AVAILABLE:
        return False
    try:
        with pikepdf.Pdf.open(src) as pdf:
            pdf.save(dst, optimize_streams=True, linearize=True)
        return os.path.exists(dst)
    except Exception:
        return False

def create_chunked_zips(file_paths, out_dir, base_name, max_bytes):
    """Pack file_paths into sequential zip files each <= max_bytes. Return list of (zip_path, files)."""
    os.makedirs(out_dir, exist_ok=True)
    parts = []
    current_files = []
    part_index = 1
    for fp in file_paths:
        current_files.append(fp)
        # write test zip
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
            parts.append((part_path, list(current_files)))
            part_index += 1
            current_files = [last]
            os.remove(test_path)
    if current_files:
        part_path = os.path.join(out_dir, f"{base_name}_part{part_index}.zip")
        with zipfile.ZipFile(part_path, 'w', compression=zipfile.ZIP_DEFLATED) as z:
            for f in current_files:
                z.write(f, arcname=os.path.basename(f))
        parts.append((part_path, list(current_files)))
    return parts

# OCR function using pdf2image + pytesseract
def extract_text_from_pdf_ocr(pdf_path, dpi=250, first_n_pages=None):
    """Return extracted text (string) using OCR. Requires pdf2image & pytesseract & poppler & tesseract."""
    if not (PDF2IMAGE_AVAILABLE and PYTESSERACT_AVAILABLE):
        return ""
    try:
        pages = convert_from_path(pdf_path, dpi=dpi)
    except Exception as e:
        # conversion failed
        return ""
    if first_n_pages:
        pages = pages[:first_n_pages]
    texts = []
    for page_img in pages:
        try:
            txt = pytesseract.image_to_string(page_img)
        except Exception:
            txt = ""
        texts.append(txt)
    return "\n".join(texts)

# simple heuristic parser to get a numeric mark and status from text
def parse_marks_and_status_from_text(txt):
    if not txt:
        return None, None
    t = txt.replace("\xa0", " ").replace("\r", " ")
    # try patterns like "XYZ 41 FAIL" or "41 FAIL"
    m = re.search(r"(\b\d{1,3}\b)\s*(?:/|\-|\s)\s*(PASS|FAIL|FAILED|ABSENT)", t, re.IGNORECASE)
    if m:
        marks = m.group(1)
        status = m.group(2).upper().rstrip(".")
        return marks, status
    # try "<label> 41 FAIL"
    m2 = re.search(r"(\b\d{1,3}\b)\s+(PASS|FAIL|FAILED|ABSENT)", t, re.IGNORECASE)
    if m2:
        return m2.group(1), m2.group(2).upper().rstrip(".")
    # try "Marks Obtained: 54" or "Marks: 54"
    m3 = re.search(r"(?:Marks?\s*(?:Obtained)?)\s*[:\-]?\s*(\d{1,3})", t, re.IGNORECASE)
    if m3:
        marks = m3.group(1)
        # try find PASS/FAIL elsewhere
        if re.search(r"\bPASS\b", t, re.IGNORECASE):
            return marks, "PASS"
        if re.search(r"\bFAIL\b", t, re.IGNORECASE):
            return marks, "FAIL"
        if re.search(r"\bABSENT\b", t, re.IGNORECASE):
            return marks, "ABSENT"
        return marks, None
    # fallback: look for ABSENT
    if re.search(r"\bABSENT\b", t, re.IGNORECASE):
        return None, "ABSENT"
    return None, None

# ---------------- UI Header ----------------
st.title("Aiclex Technologies — Email Send & Result Formation (OCR-enabled)")
st.markdown(
    """
    Upload Excel and ZIP files. Use **Email Send** to group by location and send location ZIP(s).
    Use **Result Formation** to automatically generate a mapping CSV and a Results Excel with Marks/Status extracted from marksheet PDFs (OCR).
    """
)

# ---------------- Sidebar settings ----------------
with st.sidebar:
    st.header("Shared Settings")
    smtp_host = st.text_input("SMTP host", value="smtp.hostinger.com")
    smtp_port = st.number_input("SMTP port", value=465)
    protocol = st.selectbox("Protocol", options=["SMTPS (SSL)", "SMTP (STARTTLS)"], index=0)
    sender_email = st.text_input("Sender email", value="info@aiclex.in")
    sender_password = st.text_input("Sender password", type="password")
    st.markdown("---")
    st.write("OCR availability:")
    st.write(f"pdf2image: {'Yes' if PDF2IMAGE_AVAILABLE else 'No'}")
    st.write(f"pytesseract: {'Yes' if PYTESSERACT_AVAILABLE else 'No'}")
    st.write(f"pikepdf compression: {'Yes' if PIKEPDF_AVAILABLE else 'No'}")
    st.markdown(
        """
        If OCR is not available, install system packages: `poppler` and `tesseract`, and python packages `pdf2image` and `pytesseract`.
        On Mac: `brew install poppler tesseract`.
        """
    )

# ---------------- Mode selection ----------------
mode = st.radio("Select mode:", ["Email Send", "Result Formation"], index=0)

# ---------------- EMAIL SEND MODE ----------------
if mode == "Email Send":
    st.header("Email Send")
    st.markdown("Upload **Excel** and **ZIP** (top-level) — app groups by location, lets you edit recipient for each location, prepares compressed/chunked zips, previews, and sends with delay to avoid spam.")

    uploaded_excel = st.file_uploader("Upload Excel (.xlsx/.csv) with Hallticket, Email, Location (optional: LocationRecipientEmail)", type=["xlsx","xls","csv"], key="es_excel")
    uploaded_zip = st.file_uploader("Upload top-level ZIP (PDFs / nested zips)", type=["zip"], key="es_zip")

    if not uploaded_excel or not uploaded_zip:
        st.info("Please upload both Excel and ZIP to proceed.")
    else:
        # read excel
        try:
            if uploaded_excel.name.lower().endswith("csv"):
                df = pd.read_csv(uploaded_excel, dtype=str).fillna("")
            else:
                df = pd.read_excel(uploaded_excel, dtype=str).fillna("")
        except Exception as e:
            st.error("Could not read Excel: " + str(e))
            df = None

        # extract pdfs from zip
        pdf_files = {}
        temp_zip_dir = None
        if df is not None:
            try:
                temp_zip_dir = tempfile.mkdtemp(prefix="aiclex_es_zip_")
                bio = io.BytesIO(uploaded_zip.read())
                extract_zip_recursively(bio, temp_zip_dir)
                for root, _, files in os.walk(temp_zip_dir):
                    for f in files:
                        if f.lower().endswith(".pdf"):
                            pdf_files[f] = os.path.join(root, f)
                st.success(f"Extracted PDFs: {len(pdf_files)}")
            except Exception as e:
                st.error("ZIP extraction failed: " + str(e))

        if df is not None:
            st.subheader("Column selection")
            cols = list(df.columns)
            detected_ht = next((c for c in cols if 'hall' in c.lower() or 'ticket' in c.lower()), cols[0])
            detected_email = next((c for c in cols if 'email' in c.lower() or 'mail' in c.lower()), cols[1] if len(cols)>1 else cols[0])
            detected_loc = next((c for c in cols if 'loc' in c.lower() or 'center' in c.lower() or 'city' in c.lower()), cols[2] if len(cols)>2 else cols[0])

            ht_col = st.selectbox("Hallticket column", cols, index=cols.index(detected_ht))
            email_col = st.selectbox("Candidate Email column", cols, index=cols.index(detected_email))
            location_col = st.selectbox("Location column", cols, index=cols.index(detected_loc))

            candidate_loc_recipient_cols = [c for c in cols if 'recipient' in c.lower() or 'coord' in c.lower() or 'contact' in c.lower()]
            location_recipient_col = None
            if candidate_loc_recipient_cols:
                location_recipient_col = st.selectbox("Location Recipient Email column (optional)", ["--none--"]+candidate_loc_recipient_cols)
                if location_recipient_col == "--none--":
                    location_recipient_col = None

            st.markdown("Preview (first 5 rows):")
            preview_cols = [c for c in [ht_col, email_col, location_col] if c in df.columns]
            st.dataframe(df[preview_cols].head(5))

            # build entries and group by location
            entries = []
            for _, r in df.iterrows():
                ht = str(r.get(ht_col,"")).strip()
                em = str(r.get(email_col,"")).strip()
                loc = str(r.get(location_col,"")).strip()
                loc_rec = str(r.get(location_recipient_col,"")).strip() if location_recipient_col else ""
                entries.append({"hallticket": ht, "email": em, "location": loc, "location_recipient": loc_rec})
            grouped = defaultdict(list)
            for e in entries:
                grouped[e['location']].append(e)

            st.subheader("Group summary — edit recipient per location")
            if 'location_recipients' not in st.session_state:
                st.session_state['location_recipients'] = {}
            # show groups and recipient edit fields
            for loc, items in grouped.items():
                matched_files = []
                total_bytes = 0
                for it in items:
                    candidates = [f"{it['hallticket']}.pdf", f"{it['hallticket'].upper()}.pdf", f"{it['hallticket'].lower()}.pdf"]
                    mp = ""
                    for c in candidates:
                        if c in pdf_files:
                            mp = pdf_files[c]; break
                    if not mp:
                        for fn,p in pdf_files.items():
                            if it['hallticket'] and it['hallticket'] in fn:
                                mp = p; break
                    if mp:
                        matched_files.append(mp)
                        total_bytes += os.path.getsize(mp)
                default_rec = items[0]['location_recipient'] if items and items[0]['location_recipient'] else st.session_state['location_recipients'].get(loc,"")
                col1, col2, col3 = st.columns([3,2,1])
                with col1:
                    st.markdown(f"**{loc or '(empty)'}** — Rows: {len(items)} — Matched: {len(matched_files)} — Size: {human_bytes(total_bytes)}")
                with col2:
                    new_rec = st.text_input(f"Recipient for {loc}", value=default_rec, key=f"recipient_es_{loc}")
                    st.session_state['location_recipients'][loc] = new_rec.strip()
                with col3:
                    st.markdown("✅" if matched_files else "⚠️ No PDFs")
            st.markdown("---")

            # Prepare & send settings
            st.subheader("Prepare zips & sending options")
            compress_choice = st.selectbox("Compression option", ["No compression", "Try pikepdf compression (if installed)"], index=1 if PIKEPDF_AVAILABLE else 0)
            attachment_limit_mb = st.number_input("Per-attachment limit (MB)", value=3.0, step=0.1)
            delay_seconds = st.number_input("Delay between each email (seconds)", value=2.0, step=0.5)
            subject_template = st.text_input("Subject template (use {location})", value="Hall Tickets — {location}")
            body_template = st.text_area("Body template (use {location} and {footer})", value="Dear Coordinator,\n\nPlease find attached the hall tickets for {location}.\n\n{footer}", height=140)
            footer_text = st.text_input("Footer text", value="Regards,\nAiclex Technologies\ninfo@aiclex.in")

            prepare_btn = st.button("Prepare location ZIPs (compress & chunk)")

            if prepare_btn:
                st.info("Preparing zips — this may take time.")
                workdir = tempfile.mkdtemp(prefix="aiclex_es_prep_")
                prepared_info = {}
                max_bytes = int(float(attachment_limit_mb) * 1024 * 1024)
                for loc, items in grouped.items():
                    matched_paths = []
                    for it in items:
                        candidates = [f"{it['hallticket']}.pdf", f"{it['hallticket'].upper()}.pdf", f"{it['hallticket'].lower()}.pdf"]
                        mp = ""
                        for c in candidates:
                            if c in pdf_files:
                                mp = pdf_files[c]; break
                        if not mp:
                            for fn,p in pdf_files.items():
                                if it['hallticket'] and it['hallticket'] in fn:
                                    mp = p; break
                        if mp:
                            matched_paths.append(mp)
                    if not matched_paths:
                        prepared_info[loc] = {"parts": [], "note":"no matched pdfs"}
                        continue
                    loc_dir = os.path.join(workdir, loc.replace(" ","_")[:60])
                    os.makedirs(loc_dir, exist_ok=True)
                    copied = []
                    for src in matched_paths:
                        dst = os.path.join(loc_dir, os.path.basename(src))
                        if compress_choice.startswith("Try") and PIKEPDF_AVAILABLE:
                            tmpc = dst + ".cmp.pdf"
                            ok = compress_pdf_pikepdf(src, tmpc)
                            if ok and os.path.exists(tmpc) and os.path.getsize(tmpc) < os.path.getsize(src):
                                os.replace(tmpc, dst)
                            else:
                                if os.path.exists(tmpc): os.remove(tmpc)
                                shutil.copy2(src, dst)
                        else:
                            shutil.copy2(src, dst)
                        copied.append(dst)
                    parts = create_chunked_zips(copied, out_dir=loc_dir, base_name=loc.replace(" ","_")[:40], max_bytes=max_bytes)
                    parts_info = []
                    for ppath, fls in parts:
                        parts_info.append({"path": ppath, "size": os.path.getsize(ppath), "num_files": len(fls)})
                    prepared_info[loc] = {"parts": parts_info, "recipient": st.session_state['location_recipients'].get(loc,"")}
                st.success("Preparation complete.")
                st.session_state['prepared_info'] = prepared_info
                st.session_state['workdir_es'] = workdir

            # Review & Send UI
            st.markdown("---")
            st.subheader("Review & Send")
            if 'prepared_info' not in st.session_state:
                st.info("Prepare location zips first.")
            else:
                prepared_info = st.session_state['prepared_info']
                sel_loc = st.selectbox("Select location to preview/send", options=list(prepared_info.keys()))
                info = prepared_info[sel_loc]
                st.write("Recipient:", st.session_state['location_recipients'].get(sel_loc,"(none)"))
                if info.get("parts"):
                    for idx,p in enumerate(info['parts'], start=1):
                        st.write(f"Part {idx}: {os.path.basename(p['path'])} — {human_bytes(p['size'])} — files: {p['num_files']}")
                        with open(p['path'],'rb') as f:
                            data = f.read()
                        st.download_button(label=f"Download {os.path.basename(p['path'])}", data=data, file_name=os.path.basename(p['path']), mime="application/zip")
                    if st.button(f"Send location '{sel_loc}' now"):
                        recipient = st.session_state['location_recipients'].get(sel_loc,"").strip()
                        if not recipient:
                            st.error("Recipient not set for this location.")
                        else:
                            try:
                                if protocol.startswith("SMTPS") or int(smtp_port) == 465:
                                    server = smtplib.SMTP_SSL(smtp_host, int(smtp_port), timeout=60)
                                else:
                                    server = smtplib.SMTP(smtp_host, int(smtp_port), timeout=60)
                                    if int(smtp_port) == 587:
                                        server.starttls()
                                if sender_password:
                                    server.login(sender_email, sender_password)
                            except Exception as e:
                                st.error("SMTP connection/login failed: " + str(e))
                                server = None
                            if server:
                                logs = []
                                for idx,p in enumerate(info['parts'], start=1):
                                    try:
                                        msg = EmailMessage()
                                        msg['From'] = sender_email
                                        msg['To'] = recipient
                                        subj = subject_template.format(location=sel_loc)
                                        msg['Subject'] = f"{subj} — part {idx} of {len(info['parts'])}"
                                        body = body_template.format(location=sel_loc, footer=footer_text)
                                        msg.set_content(body)
                                        with open(p['path'],'rb') as f:
                                            data = f.read()
                                        msg.add_attachment(data, maintype='application', subtype='zip', filename=os.path.basename(p['path']))
                                        server.send_message(msg)
                                        logs.append({"Location": sel_loc, "Part": idx, "Zip": os.path.basename(p['path']), "Recipient": recipient, "Status": "Sent", "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
                                        time.sleep(float(delay_seconds))
                                    except Exception as e:
                                        logs.append({"Location": sel_loc, "Part": idx, "Zip": os.path.basename(p['path']), "Recipient": recipient, "Status": "Failed", "Error": str(e)})
                                try:
                                    server.quit()
                                except:
                                    pass
                                st.success("Send attempts complete. See log:")
                                st.dataframe(pd.DataFrame(logs))
                else:
                    st.warning("No parts prepared for this location.")

            # Send ALL
            if 'prepared_info' in st.session_state and st.button("Send ALL prepared location zips (confirm)"):
                prepared_info = st.session_state['prepared_info']
                missing = [loc for loc,info in prepared_info.items() if info.get("parts") and not st.session_state['location_recipients'].get(loc,"").strip()]
                if missing:
                    st.error("Set recipients for locations first: " + ", ".join(missing))
                else:
                    try:
                        if protocol.startswith("SMTPS") or int(smtp_port) == 465:
                            server = smtplib.SMTP_SSL(smtp_host, int(smtp_port), timeout=60)
                        else:
                            server = smtplib.SMTP(smtp_host, int(smtp_port), timeout=60)
                            if int(smtp_port) == 587:
                                server.starttls()
                        if sender_password:
                            server.login(sender_email, sender_password)
                    except Exception as e:
                        st.error("SMTP connect/login failed: " + str(e))
                        server = None
                    if server:
                        all_logs = []
                        for loc, info in prepared_info.items():
                            recipient = st.session_state['location_recipients'].get(loc,"")
                            if not info.get("parts"):
                                continue
                            for idx,p in enumerate(info['parts'], start=1):
                                try:
                                    msg = EmailMessage()
                                    msg['From'] = sender_email
                                    msg['To'] = recipient
                                    subj = subject_template.format(location=loc)
                                    msg['Subject'] = f"{subj} — part {idx} of {len(info['parts'])}"
                                    body = body_template.format(location=loc, footer=footer_text)
                                    msg.set_content(body)
                                    with open(p['path'],'rb') as f:
                                        data = f.read()
                                    msg.add_attachment(data, maintype='application', subtype='zip', filename=os.path.basename(p['path']))
                                    server.send_message(msg)
                                    all_logs.append({"Location": loc, "Part": idx, "Zip": os.path.basename(p['path']), "Recipient": recipient, "Status": "Sent", "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
                                    time.sleep(float(delay_seconds))
                                except Exception as e:
                                    all_logs.append({"Location": loc, "Part": idx, "Zip": os.path.basename(p['path']), "Recipient": recipient, "Status": "Failed", "Error": str(e)})
                        try:
                            server.quit()
                        except:
                            pass
                        st.success("All sends attempted.")
                        st.dataframe(pd.DataFrame(all_logs))

            # cleanup
            if st.button("Cleanup temporary workspace (Email Send)"):
                wd = st.session_state.get('workdir_es', None)
                td = locals().get('temp_zip_dir', None)
                try:
                    if wd and os.path.exists(wd):
                        shutil.rmtree(wd)
                    if td and os.path.exists(td):
                        shutil.rmtree(td)
                    st.success("Temporary files removed.")
                except Exception as e:
                    st.error("Cleanup failed: " + str(e))

# ---------------- RESULT FORMATION MODE ----------------
elif mode == "Result Formation":
    st.header("Result Formation")
    st.markdown(
        """
        Upload an Excel (master/results) and optionally a ZIP of marksheet PDFs.
        As soon as both are available (or even with only Excel), the app will produce a downloadable CSV mapping (row-level) for quick verification.
        You can then generate the final Result Formation Excel (Summary + Details) where two new columns are added: `Marks` and `ResultStatus`.
        """
    )

    res_excel = st.file_uploader("Upload Excel (.xlsx/.csv)", type=["xlsx","xls","csv"], key="rf_excel")
    res_zip = st.file_uploader("Optional: Upload ZIP (PDFs for mark extraction)", type=["zip"], key="rf_zip")

    use_ocr = st.checkbox("Use OCR (pdf images) for mark extraction (requires poppler+tesseract)", value=True)
    ocr_dpi = st.number_input("OCR DPI (higher = slower but better)", value=250, step=50)

    rdf = None
    if res_excel:
        try:
            if res_excel.name.lower().endswith("csv"):
                rdf = pd.read_csv(res_excel, dtype=str).fillna("")
            else:
                rdf = pd.read_excel(res_excel, dtype=str).fillna("")
            st.success(f"Loaded Excel: {len(rdf)} rows")
        except Exception as e:
            st.error("Could not read Excel: " + str(e))
            rdf = None

    pdf_files_rf = {}
    temp_zip_dir_rf = None
    if res_zip and rdf is not None:
        try:
            temp_zip_dir_rf = tempfile.mkdtemp(prefix="aiclex_rf_zip_")
            bio = io.BytesIO(res_zip.read())
            extract_zip_recursively(bio, temp_zip_dir_rf)
            for root, _, files in os.walk(temp_zip_dir_rf):
                for f in files:
                    if f.lower().endswith(".pdf"):
                        pdf_files_rf[f] = os.path.join(root, f)
            st.success(f"Extracted PDFs: {len(pdf_files_rf)}")
        except Exception as e:
            st.error("ZIP extraction failed: " + str(e))

    if rdf is not None:
        st.subheader("Select columns")
        cols = list(rdf.columns)
        detected_ht = next((c for c in cols if 'hall' in c.lower() or 'ticket' in c.lower()), cols[0])
        detected_email = next((c for c in cols if 'email' in c.lower()), cols[1] if len(cols)>1 else cols[0])
        detected_loc = next((c for c in cols if 'loc' in c.lower() or 'center' in c.lower() or 'city' in c.lower()), cols[2] if len(cols)>2 else cols[0])
        ht_col_r = st.selectbox("Hallticket column", options=cols, index=cols.index(detected_ht))
        email_col_r = st.selectbox("Email column", options=cols, index=cols.index(detected_email))
        location_col_r = st.selectbox("Location column", options=cols, index=cols.index(detected_loc))
        candidate_result_cols = [c for c in cols if 'result' in c.lower() or 'status' in c.lower()]
        result_col_r = None
        if candidate_result_cols:
            result_col_r = st.selectbox("Result column (if present)", options=["--none--"]+candidate_result_cols)
            if result_col_r == "--none--":
                result_col_r = None

        st.markdown("A row-level mapping CSV will be created immediately for quick verification below.")
        # Build mapping rows immediately
        mapping_rows = []
        sr = 1
        # Prepare OCR texts cache if zip present and OCR enabled
        pdf_text_cache = {}
        if pdf_files_rf and use_ocr:
            if not (PDF2IMAGE_AVAILABLE and PYTESSERACT_AVAILABLE):
                st.warning("OCR not available in this environment. Install pdf2image & pytesseract and poppler+tesseract to enable OCR.")
            else:
                st.info("Running OCR on PDFs (may take time) — only first page of each PDF is processed by default for speed.")
                for fname, p in pdf_files_rf.items():
                    try:
                        txt = extract_text_from_pdf_ocr(p, dpi=ocr_dpi, first_n_pages=1)
                        pdf_text_cache[fname] = txt
                    except Exception:
                        pdf_text_cache[fname] = ""

        # Build pdf lookup (by numbers found in filename)
        pdf_lookup = {}
        for fname in pdf_files_rf:
            digits = re.findall(r"\d{4,}", fname)
            for d in digits:
                pdf_lookup.setdefault(d, []).append(fname)
            # also map filename without ext
            key = os.path.splitext(fname)[0]
            pdf_lookup.setdefault(key, []).append(fname)

        for _, r in rdf.iterrows():
            ht = str(r.get(ht_col_r,"")).strip()
            em = str(r.get(email_col_r,"")).strip()
            loc = str(r.get(location_col_r,"")).strip()
            matched_name = ""
            matched_flag = False
            filesize_mb = ""
            if pdf_files_rf and ht:
                # try exact keys
                mp = None
                if ht in pdf_lookup:
                    mp = pdf_lookup[ht][0]
                else:
                    # substring numeric match
                    for k,v in pdf_lookup.items():
                        if ht and ht in k:
                            mp = v[0]; break
                if not mp:
                    # try filename containing hallticket
                    for fn in pdf_files_rf:
                        if ht and ht in fn:
                            mp = fn; break
                if mp:
                    matched_name = mp
                    matched_flag = True
                    try:
                        filesize_mb = round(os.path.getsize(pdf_files_rf[mp]) / 1024 / 1024, 2)
                    except:
                        filesize_mb = ""
            mapping_rows.append({"Sr No": sr, "Hallticket No": ht, "Email": em, "Location": loc, "MatchedFile": matched_name, "Matched": "Yes" if matched_flag else "No", "FileSizeMB": filesize_mb})
            sr += 1

        map_df = pd.DataFrame(mapping_rows)
        # prepare CSV bytes for download
        csv_buf = io.StringIO()
        map_df.to_csv(csv_buf, index=False)
        csv_buf.seek(0)
        st.success("Mapping CSV ready for download.")
        st.download_button("Download Mapping CSV (mapping_check.csv)", data=csv_buf.getvalue(), file_name="mapping_check.csv", mime="text/csv")
        st.markdown("---")
        st.subheader("Mapping preview (first 100 rows)")
        st.dataframe(map_df.head(100), use_container_width=True)

        # Option to generate final Result Formation Excel with OCR-derived marks & statuses
        st.markdown("---")
        st.subheader("Generate Result Formation Excel")
        threshold = st.number_input("Pass threshold (marks >= threshold => Pass). Default 50", value=50, step=1)
        if st.button("Generate Result Formation Excel (Summary + Details)"):
            st.info("Generating Result Formation. This will attempt to extract marks/status from matched PDFs (using OCR if enabled).")
            details = []
            summary_map = {}
            sr2 = 1
            # Precompute parsed info for matched PDFs
            parsed_pdf_info = {}
            for fname, p in pdf_files_rf.items():
                text = ""
                if use_ocr and pdf_text_cache.get(fname) is not None:
                    text = pdf_text_cache.get(fname,"")
                else:
                    # try text extraction fallback using PyPDF2 if OCR not used/available
                    try:
                        import PyPDF2
                        with open(p,"rb") as fh:
                            reader = PyPDF2.PdfReader(fh)
                            txt_parts = []
                            for pg in reader.pages[:2]:
                                try:
                                    txt_parts.append(pg.extract_text() or "")
                                except:
                                    txt_parts.append("")
                            text = "\n".join(txt_parts)
                    except Exception:
                        text = ""
                marks, status = parse_marks_and_status_from_text(text)
                parsed_pdf_info[fname] = {"marks": marks, "status": status, "text_preview": text[:500]}

            # build details rows
            for _, r in rdf.iterrows():
                ht = str(r.get(ht_col_r,"")).strip()
                em = str(r.get(email_col_r,"")).strip()
                loc = str(r.get(location_col_r,"")).strip()
                resv = str(r.get(result_col_r,"")).strip() if result_col_r else ""
                # find matched pdf for this row (same matching logic as mapping)
                mp = ""
                if pdf_files_rf and ht:
                    if ht in pdf_lookup:
                        mp = pdf_lookup[ht][0]
                    else:
                        for k,v in pdf_lookup.items():
                            if ht and ht in k:
                                mp = v[0]; break
                    if not mp:
                        for fn in pdf_files_rf:
                            if ht and ht in fn:
                                mp = fn; break
                # decide marks/status
                final_marks = ""
                final_status = ""
                if mp:
                    info = parsed_pdf_info.get(mp, {})
                    marks_val = info.get("marks")
                    status_val = info.get("status")
                    if marks_val and marks_val.isdigit():
                        mk = int(marks_val)
                        final_marks = str(mk)
                        final_status = "Pass" if mk >= int(threshold) else "Fail"
                    else:
                        # if status_val indicates ABSENT
                        if status_val and status_val.upper().startswith("ABSENT"):
                            final_marks = ""
                            final_status = "Absent"
                        elif resv and isinstance(resv, str) and resv.strip() != "":
                            # fallback to given result column value if provided
                            final_marks = ""
                            final_status = resv
                        else:
                            final_marks = ""
                            final_status = "Pending"
                else:
                    # no PDF matched
                    if resv and isinstance(resv, str) and resv.strip() != "":
                        # if result column exists use it
                        if resv.strip().lower() == "absent":
                            final_marks = ""
                            final_status = "Absent"
                        else:
                            final_marks = ""
                            final_status = resv
                    else:
                        final_marks = ""
                        final_status = "No PDF matched"
                details.append({"Sr No": sr2, "Hallticket No": ht, "Email": em, "Location": loc, "Marks": final_marks, "ResultStatus": final_status})
                sr2 += 1
                # summary aggregation
                if loc not in summary_map:
                    summary_map[loc] = {"rows":0, "matched":0, "bytes":0}
                summary_map[loc]["rows"] += 1
                if mp:
                    summary_map[loc]["matched"] += 1
                    try:
                        summary_map[loc]["bytes"] += os.path.getsize(pdf_files_rf[mp])
                    except:
                        pass

            sum_rows = []
            for loc, v in summary_map.items():
                sum_rows.append({"Location": loc, "Rows": v["rows"], "MatchedPDFs": v["matched"], "TotalBytes": v["bytes"], "TotalSize": human_bytes(v["bytes"])})
            sum_df = pd.DataFrame(sum_rows)
            details_df = pd.DataFrame(details)

            # write Excel in-memory
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
                sum_df.to_excel(writer, sheet_name="Summary", index=False)
                details_df.to_excel(writer, sheet_name="Details", index=False)
            out.seek(0)
            st.success("Result Formation Excel ready.")
            st.download_button("Download Result Formation Excel", data=out.getvalue(), file_name="result_formation.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # cleanup button for RF
    if st.button("Cleanup temporary workspace (Result Formation)"):
        try:
            td = locals().get('temp_zip_dir_rf', None)
            if td and os.path.exists(td):
                shutil.rmtree(td)
            st.success("Cleaned result temp workspace.")
        except Exception as e:
            st.error("Cleanup failed: " + str(e))

# ---------------- Footer ----------------
st.markdown("---")
st.info("Notes: 1) If PDFs are scanned images you must enable OCR (install poppler & tesseract) for marks extraction. 2) For very large sends use SendGrid/SES and S3 signed links.")
