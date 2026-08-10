# -*- coding: utf-8 -*-
"""Render a CV JSON (master-cv.json schema) into a designed, ATS-safe, one-page PDF.

Usage:
    python render_cv.py <cv.json> <output.pdf>

Pipeline: JSON -> HTML (embedded template) -> Microsoft Edge headless print-to-pdf.
Design: single column, blue accents, Segoe UI - real text only (ATS friendly).
AUTO-FIT: all sizes are rem-based; an inline script binary-searches the root
font-size so the content always fills exactly one full A4 page before printing.
"""
import html
import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from string import Template

EDGE_PATHS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
]

PAGE = Template(r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>$name - CV</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  @page { size: A4; margin: 0; }
  html { font-size: 11.5px; }
  html, body { background: #ffffff; }
  body {
    font-family: "Segoe UI", Calibri, Arial, sans-serif;
    color: #24292f;
    width: 210mm;
    padding: 12mm 15mm 10mm 15mm;
    font-size: 1rem; line-height: 1.5;
    -webkit-print-color-adjust: exact; print-color-adjust: exact;
  }
  a { color: #2563a8; text-decoration: none; }
  b, strong { font-weight: 600; color: #1b2733; }

  .header { display: flex; justify-content: space-between; align-items: flex-end;
            border-bottom: 3px solid #2563a8; padding-bottom: 0.78rem; }
  .name { font-size: 2.52rem; font-weight: 700; letter-spacing: 0.8px; color: #16385e; line-height: 1.1; }
  .role-title { font-size: 1.22rem; color: #2563a8; font-weight: 600; margin-top: 0.43rem; letter-spacing: 0.3px; }
  .contact { text-align: right; font-size: 0.94rem; color: #444d56; line-height: 1.65; }
  .contact .sep { color: #b6c2cf; margin: 0 5px; }

  .section { margin-top: 1.04rem; }
  .section-title {
    font-size: 1.07rem; font-weight: 700; text-transform: uppercase; letter-spacing: 1.5px;
    color: #2563a8; border-bottom: 1.5px solid #cfe0f0; padding-bottom: 0.26rem; margin-bottom: 0.61rem;
  }
  .summary { color: #333c45; }

  .item { margin-bottom: 0.78rem; }
  .item:last-child { margin-bottom: 0; }
  .item-head { display: flex; justify-content: space-between; align-items: baseline; }
  .item-title { font-weight: 700; font-size: 1.07rem; color: #1b2733; }
  .item-title .at { font-weight: 600; color: #2563a8; }
  .period { font-size: 0.91rem; color: #6a737d; white-space: nowrap; margin-left: 10px; }
  .item-note { font-size: 0.91rem; color: #57606a; font-style: italic; margin-top: 0.17rem; }
  ul { margin: 0.35rem 0 0 1.4rem; }
  li { margin-bottom: 0.22rem; }

  .skills-table { width: 100%; border-collapse: collapse; }
  .skills-table td { padding: 0.17rem 0; vertical-align: top; }
  .skill-cat { font-weight: 700; color: #1b2733; width: 10.9rem; padding-right: 10px; white-space: nowrap; }

  .edu-high { margin: 0.35rem 0 0 1.4rem; }
  .award { margin-top: 0.61rem; color: #333c45; }
  .langs { color: #333c45; }
</style>
</head>
<body>
  <div class="header">
    <div>
      <div class="name">$name</div>
      <div class="role-title">$title</div>
    </div>
    <div class="contact">
      <div>$phone <span class="sep">|</span> <a href="mailto:$email">$email</a></div>
      <div><a href="$linkedin">LinkedIn: $linkedin_label</a> <span class="sep">|</span> <a href="$github">GitHub: $github_label</a></div>
    </div>
  </div>

  <div class="section">
    <div class="section-title">Professional Summary</div>
    <div class="summary">$summary</div>
  </div>

  <div class="section">
    <div class="section-title">Work Experience</div>
    $experience_html
  </div>

  <div class="section">
    <div class="section-title">Education</div>
    $education_html
  </div>

  $projects_section

  <div class="section">
    <div class="section-title">Skills</div>
    <table class="skills-table">$skills_html</table>
  </div>

  <div class="section">
    <div class="section-title">Military Service</div>
    <div class="summary">$military</div>
  </div>

  <div class="section">
    <div class="section-title">Languages</div>
    <div class="langs">$languages</div>
  </div>

<script>
(function () {
  // Fill exactly one A4 page: 297mm at 96dpi = 1122.5px. Keep a hair of slack.
  var TARGET = 1119;
  var root = document.documentElement;
  var lo = 8.5, hi = 16.0, best = lo;
  for (var i = 0; i < 16; i++) {
    var mid = (lo + hi) / 2;
    root.style.fontSize = mid + 'px';
    if (document.body.scrollHeight <= TARGET) { best = mid; lo = mid; }
    else { hi = mid; }
  }
  root.style.fontSize = best + 'px';
  // Safety: never overflow to page 2.
  while (document.body.scrollHeight > TARGET && best > 8.5) {
    best -= 0.1;
    root.style.fontSize = best + 'px';
  }
})();
</script>
</body>
</html>""")

SKILL_LABELS = {
    "languages": "Languages", "frameworks": "Frameworks", "data": "Databases",
    "concepts": "Concepts", "tools": "Tools",
}


def esc(text):
    """HTML-escape, then convert **bold** markers to <b>."""
    out = html.escape(str(text))
    return re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", out)


def bullets_html(items):
    return "<ul>" + "".join(f"<li>{esc(b)}</li>" for b in items) + "</ul>"


def build_html(cv):
    c = cv.get("contact", {})

    exp_parts = []
    for e in cv.get("experience", []):
        note = f'<div class="item-note">{esc(e["note"])}</div>' if e.get("note") else ""
        exp_parts.append(
            '<div class="item"><div class="item-head">'
            f'<span class="item-title">{esc(e["role"])} <span class="at">| {esc(e["company"])}</span></span>'
            f'<span class="period">{esc(e.get("period", ""))}</span></div>'
            f'{note}{bullets_html(e.get("bullets", []))}</div>'
        )

    edu_parts = []
    for ed in cv.get("education", []):
        highs = "".join(f"<li>{esc(h)}</li>" for h in ed.get("highlights", []))
        edu_parts.append(
            '<div class="item"><div class="item-head">'
            f'<span class="item-title">{esc(ed["degree"])} <span class="at">| {esc(ed["institution"])}</span></span>'
            f'<span class="period">{esc(ed.get("period", ""))}</span></div>'
            f'<ul class="edu-high">{highs}</ul></div>'
        )
    for award in cv.get("awards", []):
        edu_parts.append(f'<div class="award"><b>{esc(award.split(" - ")[0])}</b> - {esc(" - ".join(award.split(" - ")[1:]))}</div>'
                         if " - " in award else f'<div class="award">{esc(award)}</div>')

    projects = cv.get("projects", [])
    if projects:
        proj_parts = []
        for p in projects:
            proj_parts.append(
                '<div class="item">'
                f'<span class="item-title">{esc(p["name"])} <span class="at">({esc(p.get("tech", ""))})</span></span>'
                f'{bullets_html(p.get("bullets", []))}</div>'
            )
        projects_section = ('<div class="section"><div class="section-title">Relevant Projects</div>'
                            + "".join(proj_parts) + "</div>")
    else:
        projects_section = ""

    skills_rows = []
    for key, vals in cv.get("skills", {}).items():
        label = SKILL_LABELS.get(key, key.replace("_", " ").title())
        skills_rows.append(f'<tr><td class="skill-cat">{esc(label)}:</td><td>{esc(", ".join(vals))}</td></tr>')

    return PAGE.substitute(
        name=esc(cv["name"]),
        title=esc(cv.get("title", "")),
        phone=esc(c.get("phone", "")),
        email=esc(c.get("email", "")),
        linkedin=html.escape(c.get("linkedin", "#"), quote=True),
        linkedin_label=esc(c.get("linkedin_label", "Profile")),
        github=html.escape(c.get("github", "#"), quote=True),
        github_label=esc(c.get("github_label", "Profile")),
        summary=esc(cv.get("summary", cv.get("summary_base", ""))),
        experience_html="".join(exp_parts),
        education_html="".join(edu_parts),
        projects_section=projects_section,
        skills_html="".join(skills_rows),
        military=esc(cv.get("military", {}).get("text", "")),
        languages=" | ".join(esc(l) for l in cv.get("languages_spoken", [])),
    )


def find_browser():
    for p in EDGE_PATHS:
        if Path(p).exists():
            return p
    raise FileNotFoundError("No Edge/Chrome found for PDF rendering")


def render_pdf(html_path, pdf_path):
    browser = find_browser()
    profile_dir = Path(tempfile.gettempdir()) / "edge-headless-cv-profile"
    cmd = [
        browser, "--headless", "--disable-gpu", "--no-first-run",
        "--no-default-browser-check", "--disable-extensions",
        f"--user-data-dir={profile_dir}",
        "--no-pdf-header-footer",
        f"--print-to-pdf={pdf_path}",
        Path(html_path).as_uri(),
    ]
    res = subprocess.run(cmd, capture_output=True, text=True, timeout=90)
    if not Path(pdf_path).exists():
        raise RuntimeError(f"PDF not created. Browser output:\n{res.stderr[-2000:]}")


def page_count(pdf_path):
    try:
        from pypdf import PdfReader
        return len(PdfReader(str(pdf_path)).pages)
    except Exception:
        return -1


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    cv_json, out_pdf = Path(sys.argv[1]).resolve(), Path(sys.argv[2]).resolve()
    cv = json.loads(cv_json.read_text(encoding="utf-8"))
    html_out = out_pdf.with_suffix(".html")
    html_out.write_text(build_html(cv), encoding="utf-8")
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    render_pdf(html_out, out_pdf)
    pages = page_count(out_pdf)
    print(f"OK: {out_pdf} ({pages} page{'s' if pages != 1 else ''})")
    if pages > 1:
        print("WARNING: CV exceeds 1 page even at minimum font size - trim content in the JSON and re-render.")


if __name__ == "__main__":
    main()
