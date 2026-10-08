"""
pdf_downloader.py  --  save every Coursera certificate link as  <RollNo>_<Course>.pdf

Reads the workbook produced by certificate_linker.py (any sheet that has a
"Certificate link" column), opens each link in a headless Chrome (Playwright) and
prints the page to PDF.  One folder per sheet / semester.  Safe to re-run: files that
already exist are skipped, so a stopped run resumes where it left off.

One-time setup on your computer:
    pip install playwright pandas openpyxl
    python -m playwright install chromium

Command line:
    python pdf_downloader.py Abstract_with_certificate_links.xlsx -o certificates --range 1-50
    python pdf_downloader.py file.xlsx --range "1-50, 80, 100-120" --sheets "I MCA Coursera Report"
"""
import argparse
import re
import time
import zipfile
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

SKIP_SHEETS = {"verification", "summary"}
ID_ALIASES = {"external id", "roll no", "roll no.", "register no", "register no.", "reg no",
              "reg no.", "register number", "roll number", "usn"}


# ----------------------------------------------------------------------------- helpers
def safe(text, maxlen=80):
    """Make text safe for a file/folder name."""
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', " ", str(text or "")).strip()
    s = re.sub(r"\s+", "_", s)
    return s[:maxlen].strip("._") or "untitled"


def parse_ranges(text):
    """'1-50, 80, 100-120'  ->  {1..50, 80, 100..120}   (both ends inclusive). Empty -> None (= all)."""
    if text is None or not str(text).strip():
        return None
    nums = set()
    cleaned = re.sub(r"\s*(?:-|\u2013|to|:)\s*", "-", str(text).strip().lower())   # '10 to 20' -> '10-20'
    for part in re.split(r"[,\s;]+", cleaned):
        if not part:
            continue
        m = re.fullmatch(r"(\d+)-(\d+)", part) or re.fullmatch(r"(\d+)", part)
        if not m:
            raise ValueError(f"Cannot understand '{part}'. Use e.g.  1-50, 80, 100-120")
        a = int(m.group(1)); b = int(m.group(2)) if m.lastindex == 2 else a
        if a < 1 or b < a:
            raise ValueError(f"Bad range '{part}' (numbers start at 1 and must go upward)")
        nums.update(range(a, b + 1))
    return nums


def read_jobs(workbook):
    """Return list of dicts: sheet, roll, name, course, url (only rows holding a real link)."""
    wb = load_workbook(workbook, read_only=False)
    jobs = []
    for ws in wb.worksheets:
        if ws.title.strip().lower() in SKIP_SHEETS:
            continue
        head = {str(c.value).strip().lower(): c.column for c in ws[1] if c.value is not None}
        id_c = next((head[k] for k in head if k in ID_ALIASES), None)
        name_c = head.get("name")
        course_c = next((head[k] for k in head if k in {"course", "course name"}), None)
        link_c = head.get("certificate link")
        if not (id_c and course_c and link_c):
            continue
        for r in range(2, ws.max_row + 1):
            link_cell = ws.cell(r, link_c)
            url = (link_cell.hyperlink.target if link_cell.hyperlink else None) or link_cell.value
            url = str(url).strip() if url else ""
            roll = ws.cell(r, id_c).value
            if not roll:
                continue
            jobs.append(dict(sheet=ws.title, roll=str(roll).strip(),
                             name=ws.cell(r, name_c).value if name_c else "",
                             course=ws.cell(r, course_c).value, url=url))
    return jobs


def plan_paths(jobs, out_dir):
    """Give every job a unique  <out>/<sheet>/<Roll>_<Course>.pdf  path."""
    seen = set()
    for j in jobs:
        folder = Path(out_dir) / safe(j["sheet"], 60)
        base = f"{safe(j['roll'], 20)}_{safe(j['course'])}"
        p, n = folder / f"{base}.pdf", 2
        while p in seen:                       # same roll + same course twice -> _2, _3 ...
            p = folder / f"{base}_{n}.pdf"
            n += 1
        seen.add(p)
        j["path"] = p
    return jobs


# ----------------------------------------------------------------------------- browser
FIND_CERT_JS = """
() => {
  const els = [...document.querySelectorAll('img,canvas,svg,iframe,embed,object')];
  let best = null, area = 0;
  for (const e of els) {
    const r = e.getBoundingClientRect(), st = getComputedStyle(e);
    if (st.visibility === 'hidden' || st.display === 'none' || r.width < 400 || r.height < 250) continue;
    if (e.tagName === 'IMG' && !e.complete) continue;
    const a = r.width * r.height;
    if (a > area) { area = a; best = e; }
  }
  document.querySelectorAll('[data-cert-target]').forEach(x => x.removeAttribute('data-cert-target'));
  if (!best) return false;
  best.setAttribute('data-cert-target', '1');
  return true;
}
"""


def _png_size(png_bytes):
    import struct
    return struct.unpack(">II", png_bytes[16:24])


class ChromePrinter:
    """
    Opens a link and saves a PDF.
      certificate_only=True : finds the big certificate image on the page, saves ONLY that
                              (page size = certificate size). Falls back to the whole page
                              if no certificate is detected, and says so in the report.
      certificate_only=False: whole page as one tall PDF page.
    """
    SCALE = 3          # sharpness of the certificate image

    def __init__(self, timeout_ms=60000, certificate_only=True):
        from playwright.sync_api import sync_playwright
        self._pw = sync_playwright().start()
        self._browser = self._pw.chromium.launch(headless=True)
        self._ctx = self._browser.new_context(viewport={"width": 1400, "height": 1000},
                                              device_scale_factor=self.SCALE if certificate_only else 1)
        self._timeout = timeout_ms
        self._only = certificate_only
        self._debug_saved = 0

    def _check_page(self, page, resp):
        if resp is not None and resp.status >= 400:
            raise RuntimeError(f"HTTP {resp.status}")
        text = (page.inner_text("body") or "").lower()
        if any(t in text for t in ("page not found", "doesn't exist", "does not exist",
                                   "invalid certificate", "no longer available")):
            raise RuntimeError("Coursera says the certificate page is not available")

    def _save_certificate_only(self, page, path):
        page.wait_for_timeout(1500)
        if not page.evaluate(FIND_CERT_JS):
            return False
        png = page.locator('[data-cert-target="1"]').first.screenshot(type="png")
        wpx, hpx = _png_size(png)
        w, h = wpx / self.SCALE, hpx / self.SCALE
        import base64
        html = (f'<html><body style="margin:0"><img style="display:block;width:{w}px;height:{h}px" '
                f'src="data:image/png;base64,{base64.b64encode(png).decode()}"></body></html>')
        pdf_page = self._ctx.new_page()
        try:
            pdf_page.set_content(html)
            path.parent.mkdir(parents=True, exist_ok=True)
            pdf_page.pdf(path=str(path), width=f"{w}px", height=f"{h}px", print_background=True,
                         page_ranges="1", margin=dict(top="0", bottom="0", left="0", right="0"))
        finally:
            pdf_page.close()
        return True

    def __call__(self, url, path):
        page = self._ctx.new_page()
        try:
            resp = page.goto(url, wait_until="networkidle", timeout=self._timeout)
            self._check_page(page, resp)
            if self._only and self._save_certificate_only(page, path):
                return ""
            note = ""
            if self._only:
                note = "Certificate image not detected - saved full page"
                if self._debug_saved < 3:               # keep a copy so the detector can be tuned
                    dbg = path.parent.parent / "_debug"
                    dbg.mkdir(parents=True, exist_ok=True)
                    (dbg / f"{path.stem}.html").write_text(page.content(), encoding="utf-8")
                    self._debug_saved += 1
            page.wait_for_timeout(500)
            height = page.evaluate("Math.max(document.body.scrollHeight, "
                                   "document.documentElement.scrollHeight)")
            page.emulate_media(media="screen")
            path.parent.mkdir(parents=True, exist_ok=True)
            page.pdf(path=str(path), width="1400px", height=f"{min(int(height) + 20, 14000)}px",
                     print_background=True, margin=dict(top="0", bottom="0", left="0", right="0"))
            return note
        finally:
            page.close()

    def close(self):
        self._browser.close()
        self._pw.stop()


# ----------------------------------------------------------------------------- main routine
def download_certificates(workbook, out_dir, limit=None, delay=2.0, retries=2,
                          progress=None, render=None, certificate_only=True,
                          select=None, sheets=None):
    """
    workbook : path / file-like of the linker output
    select   : which certificates, by number, e.g. '1-50' or '1-50, 80, 100-120' (None = all).\n               Numbers count the downloadable links in the chosen sheets, in sheet order (1,2,3...)\n    sheets   : list of sheet names to include (None = all sheets)\n    limit    : only the first N of the selection (use 10 for a trial run)
    delay    : seconds to wait between downloads (be polite to Coursera)
    progress : optional callback(done, total, message)
    render   : optional function(url, path) -> used instead of Chrome (for testing)
    Returns a DataFrame report (one row per link).
    """
    out_dir = Path(out_dir)
    jobs = plan_paths(read_jobs(workbook), out_dir)
    report = []
    wanted = parse_ranges(select)
    todo, n = [], 0
    for j in jobs:
        if sheets is not None and j["sheet"] not in sheets:
            continue
        if not j["url"].lower().startswith("http"):
            report.append({**j, "no": "", "status": "No link (Link not available)", "error": ""})
            continue
        n += 1
        j["no"] = n
        if wanted is None or n in wanted:
            todo.append(j)
    if limit:
        todo = todo[:int(limit)]

    printer = render or ChromePrinter(certificate_only=certificate_only)
    try:
        for i, j in enumerate(todo, 1):
            row = {k: j[k] for k in ("no", "sheet", "roll", "name", "course", "url")}
            row["file"] = str(j["path"])
            if j["path"].exists() and j["path"].stat().st_size > 0:
                row.update(status="Skipped (already downloaded)", error="")
            else:
                err, note = "", ""
                for attempt in range(retries + 1):
                    try:
                        note = printer(j["url"], j["path"]) or ""
                        err = ""
                        break
                    except Exception as e:                      # noqa: BLE001
                        err = str(e).splitlines()[0][:200]
                        time.sleep(delay * (attempt + 1))
                row.update(status="Downloaded" if not err else "FAILED", error=err or note)
                time.sleep(delay)
            report.append(row)
            if progress:
                progress(i, len(todo), f"{j['roll']} - {str(j['course'])[:50]}")
    finally:
        if render is None:
            printer.close()

    df = pd.DataFrame(report)
    out_dir.mkdir(parents=True, exist_ok=True)
    df.drop(columns=["path"], errors="ignore").to_csv(out_dir / "download_report.csv", index=False)
    return df.drop(columns=["path"], errors="ignore")


def zip_folder(folder, zip_path_or_buffer):
    folder = Path(folder)
    with zipfile.ZipFile(zip_path_or_buffer, "w", zipfile.ZIP_DEFLATED) as z:
        for p in folder.rglob("*"):
            if p.is_file():
                z.write(p, p.relative_to(folder))


# ----------------------------------------------------------------------------- CLI
if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Download Coursera certificate links as PDFs.")
    ap.add_argument("workbook")
    ap.add_argument("-o", "--out", default="certificates")
    ap.add_argument("--range", default=None, help="certificate numbers, e.g. 1-50  or  1-50,80,100-120")
    ap.add_argument("--sheets", nargs="*", default=None, help="only these sheet names")
    ap.add_argument("--limit", type=int, default=None, help="only first N of the selection (trial run)")
    ap.add_argument("--delay", type=float, default=2.0)
    ap.add_argument("--full-page", action="store_true", help="save the whole page, not just the certificate")
    a = ap.parse_args()
    rep = download_certificates(a.workbook, a.out, a.limit, a.delay, certificate_only=not a.full_page, select=a.range, sheets=a.sheets,
                                
                                progress=lambda d, t, m: print(f"[{d}/{t}] {m}"))
    print(rep["status"].value_counts().to_string())
    print(f"Report: {Path(a.out) / 'download_report.csv'}")
