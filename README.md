# JobHunt

My local job hunting setup. A scheduled Claude Code agent scans company career sites and job boards every morning, scores new postings against my profile, and drops them into an Excel tracker with an RTL Hebrew dashboard on top. From the dashboard, one click generates a CV tailored to a specific posting, and another click sends a browser agent to fill the actual application form in Chrome.

Built for exactly one user (me), on Windows, but the pieces are generic enough to steal.

## The moving parts

```
scripts/
  add_jobs.py        appends scan results to jobs.xlsx, dedupes by link / job id
  make_dashboard.py  renders dashboard.html from the tracker: score badges, filters,
                     batch pipeline, live task console. Single self-contained file,
                     no CDN, works over file://
  render_cv.py       CV JSON -> HTML -> Edge headless print-to-pdf. One page,
                     ATS-safe real text, binary-searches the root font size until
                     the content fills an A4 exactly
  cv_server.py       localhost server behind the dashboard buttons. Runs headless
                     Claude Code to generate CVs, runs Codex CLI to fill application
                     forms in the real Chrome, streams both logs into a task console
                     in the page
  archive_stale.py   auto-archives stale postings (21 days for new, 30 for rejected)
  set_cv.py          records a generated CV on its tracker row
  delete_jobs.py     removes rows, remembers them so scans never re-add
  make_web_report.py read-only phone-friendly copy of the tracker
```

`start-cv-server.bat` starts the local server minimized. `HOW-IT-WORKS.md` is the user manual (Hebrew).

## What is not here

The actual data: `jobs.xlsx`, `master-cv.json`, `profile.md`, `SOURCES.md` (per-site scan recipes), generated CVs and dashboards. All personal, all gitignored. The scripts expect those files to exist next to them, so cloning this repo gets you the tooling, not a working setup.

## Requirements

Python 3.11+ with `openpyxl`. Microsoft Edge for the PDF printing. Claude Code and Codex CLI on PATH if you want the dashboard buttons to do anything.
