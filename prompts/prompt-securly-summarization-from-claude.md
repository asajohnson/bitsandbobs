# Prompt to recreate securly-summarization.py

Create a Python script called `securly_report.py` that I can run from my Windows desktop. It connects to my local Outlook 2016 (classic) application via COM (pywin32), reads HTML emails from the folder **Inbox > Securly Reports**, and generates per-class Word document summaries.

**Dependencies:** `pywin32`, `beautifulsoup4`, `python-docx`, `lxml`

**Date range:** The script determines the current business week (Monday–Friday) based on the day it's run. It pulls all emails within that range.

**Email format:** Each email is a Securly Classroom session summary. The subject contains the class name and section code (e.g., `DESIGN & MODELING - KDM078_01 session summary (03/27/2026 01:45 PM - 02:40 PM)`). The HTML body contains:

- A class name in a `<td>` with `font-size: 30px` style (includes a section code suffix like ` - KDM078/01` that should be stripped)
- A date in `MM/DD/YYYY` format in a `<td>`
- Student browsing blocks in `<div>` elements with `border-radius: 10px` and `border: 1px` styles
- Student names in `<td>` with `font-weight: 600` and `font-size: 16px` styles, in "Last, First" format
- Website rows in `<table>` elements with `padding-left: 24px` style, containing an `<a>` tag with the domain and `<span>` elements with `font-weight: 600` for minutes (followed by a sibling `<span>` containing "min")

**Class name mapping:**

- `STEM-ADV ROBOTICS & AUTOMATION` → `Adv. Robotics`
- `CS DISCOVERIES 2- GAME DESIGN` → `Game Design`
- `DESIGN & MODELING` → `Design & Modeling`
- Any unmapped class falls back to title case

**Output:** One `.docx` file per class, saved to the Desktop. Filename format: `<ClassName>-Week of <YYYY-MM-DD>-<YYYY-MM-DD>.docx`

**Document formatting:**

- Font: Arial throughout
- Narrow margins (0.75")
- Centered title: `<ClassName> - Week of <Month Day> - <Month Day, Year>`
- A single table with "Table Grid" style, centered, with blue header shading (`#D9E2F3`)
- Font size: 8pt for table content, 14pt bold for title

**Table columns:**

1. **Student Name** – De-duped: when a student has multiple website rows, merge the Student Name cells vertically so the name appears once in a merged cell spanning all their rows. Students sorted alphabetically.
2. **Websites Visited** – The domain visited. Sorted alphabetically within each student.
3. **Dates Visited** – Dates the site was visited, in `MM/DD` format (no year), comma-separated.
4. **Total Minutes Visited** – Total minutes across all dates for that website, shown with one decimal place.
5. **Participation Deduction** – `0` if <10 min, `-1` if 10–19.99 min, `-2` if ≥20 min.
6. **Referral Qualified** – `Yes` if ≥10 min, `-` if <10 min. **Exception:** always `-` for sites matching `*vex.com*`, `*instructure.com*`, `*office.com*`, or `*code.org*`.
7. **Total Instructional Minutes Spent** – Same minutes value but with trailing zeros removed (e.g., `12` not `12.0`, but `12.5` stays).
8. **Total Instructional %** – `(minutes / 235) × 100`, displayed as a percentage with one decimal (e.g., `5.3%`).

**Desktop path:** `%USERPROFILE%\OneDrive - issaquah.wednet.edu\Desktop`
