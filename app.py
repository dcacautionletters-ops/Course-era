"""Streamlit app:  streamlit run app.py"""
import io
import subprocess
import sys
import tempfile
from pathlib import Path

import streamlit as st

from certificate_linker import process
from pdf_downloader import download_certificates, parse_ranges, read_jobs, zip_folder

st.set_page_config(page_title="Coursera Certificate Tools", page_icon="🎓", layout="wide")
st.title("🎓 Coursera Certificate Tools")
tab1, tab2 = st.tabs(["1️⃣ Paste certificate links", "2️⃣ Download certificates as PDF"])

# =============================================================== TAB 1
with tab1:
    st.caption("Upload the abstract + Coursera export file(s). Get the abstract back with a Certificate link column.")
    abstract = st.file_uploader("Abstract workbook (one sheet per class)", type=["xlsx"], key="abs")
    exports = st.file_uploader("Coursera LearnerActivity export(s) - upload all of them",
                               type=["xlsx"], accept_multiple_files=True, key="exp")

    if st.button("Generate", type="primary", disabled=not (abstract and exports)):
        try:
            buf = io.BytesIO()
            with st.spinner("Matching students and courses..."):
                allrows, issues = process(io.BytesIO(abstract.getvalue()),
                                          [io.BytesIO(f.getvalue()) for f in exports], buf)
            st.session_state["linked_xlsx"] = buf.getvalue()      # reused by tab 2
            st.session_state["linker_issues"] = issues
            st.session_state["linker_counts"] = (len(allrows), int((allrows["Link status"] == "Found").sum()))
        except Exception as e:
            st.error(f"Could not process: {e}")

    if "linked_xlsx" in st.session_state:
        total, found = st.session_state["linker_counts"]
        c1, c2, c3 = st.columns(3)
        c1.metric("Rows", total); c2.metric("Links pasted", found)
        c3.metric("Link not available", total - found)
        st.download_button("⬇️ Download result", st.session_state["linked_xlsx"],
                           file_name="Abstract_with_certificate_links.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        issues = st.session_state["linker_issues"]
        if len(issues):
            st.subheader("Rows needing attention")
            st.dataframe(issues, use_container_width=True, hide_index=True)
        else:
            st.success("Every row matched with a link. Name / Roll No cross-check is clean.")
        st.info("Full match details for every row are in the 'Verification' sheet of the downloaded file.")

# =============================================================== TAB 2
@st.cache_resource(show_spinner="Setting up the browser (first time only)...")
def ensure_browser():
    subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)
    return True


with tab2:
    st.caption("Saves each certificate link as a PDF named  RollNo_CourseName.pdf, one folder per semester.")
    src = None
    if "linked_xlsx" in st.session_state:
        st.success("Using the workbook just generated in tab 1.")
        src = io.BytesIO(st.session_state["linked_xlsx"])
    up = st.file_uploader("...or upload a workbook that already has a 'Certificate link' column",
                          type=["xlsx"], key="up2")
    if up:
        src = io.BytesIO(up.getvalue())

    if src:
        jobs = read_jobs(io.BytesIO(src.getvalue()))
        real = sum(j["url"].lower().startswith("http") for j in jobs)
        st.write(f"**{real}** certificate links found across "
                 f"**{len({j['sheet'] for j in jobs})}** sheet(s).")
        all_sheets = list(dict.fromkeys(j["sheet"] for j in jobs))
        chosen = st.multiselect("Sheets / semesters", all_sheets, default=all_sheets)
        n_sel = sum(j["url"].lower().startswith("http") for j in jobs if j["sheet"] in chosen)
        c1, c2 = st.columns([2, 1])
        rng = c1.text_input(f"Certificate numbers (1 to {n_sel}) - e.g.  1-50   or   1-50, 80, 100-120   (blank = all)",
                            value="1-10")
        delay = c2.slider("Pause between downloads (seconds)", 1.0, 10.0, 2.0, 0.5)
        try:
            picked = parse_ranges(rng)
            count = n_sel if picked is None else len([x for x in picked if x <= n_sel])
            st.caption(f"➡️ {count} certificate(s) will be downloaded. Next batch tip: use {min(n_sel, (max(picked) if picked else 0) + 1)}-{min(n_sel, (max(picked) if picked else 0) + 50)}.")
        except ValueError as e:
            st.error(str(e)); picked = False
        cert_only = st.checkbox("Certificate only (crop out the rest of the page)", value=True)
        st.warning("Large runs take long (about 5-10 s per certificate). For all ~1,000, use "
                   "`python pdf_downloader.py file.xlsx` on your own computer; this tab is best for test batches.")

        if st.button("Download PDFs", type="primary", disabled=(picked is False or not chosen)):
            try:
                ensure_browser()
                tmp = Path(tempfile.mkdtemp()) / "certificates"
                bar, msg = st.progress(0.0), st.empty()
                rep = download_certificates(
                    io.BytesIO(src.getvalue()), tmp, select=rng, sheets=chosen, delay=delay, certificate_only=cert_only,
                    progress=lambda d, t, m: (bar.progress(d / max(t, 1)), msg.write(f"{d}/{t}  {m}")))
                bar.progress(1.0)
                st.session_state["pdf_report"] = rep
                zbuf = io.BytesIO(); zip_folder(tmp, zbuf)
                st.session_state["pdf_zip"] = zbuf.getvalue()
            except Exception as e:
                st.error(f"Download failed: {e}")

    if "pdf_report" in st.session_state:
        rep = st.session_state["pdf_report"]
        st.write(rep["status"].value_counts().to_frame("count"))
        st.download_button("⬇️ Download ZIP of PDFs", st.session_state["pdf_zip"],
                           file_name="certificates.zip", mime="application/zip")
        full = rep[rep["error"].astype(str).str.contains("not detected", na=False)]
        if len(full):
            st.warning(f"{len(full)} file(s) could not be cropped to the certificate and were saved as full pages.")
        bad = rep[rep["status"].str.contains("FAILED|No link", na=False)]
        if len(bad):
            st.subheader("Not downloaded")
            st.dataframe(bad[["no", "sheet", "roll", "name", "course", "status", "error"]],
                         use_container_width=True, hide_index=True)
