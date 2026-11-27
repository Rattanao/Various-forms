---
name: bundle-enter-check
description: "Check EDI for a bundled ENTER: one PDF that combines many forwarders' forms. Splits it page by page, reads each page with the reader for its form, compares everything with the MF / MANIFEST (.xls or MF PDF) and writes the colour-coded EDI.xlsx. Pages of forwarders that have a dedicated reader (Sinotrans, LX Pantos, Freight Links, Shipco, Planet, Union World, AGN, Handle Inter, Able, Quality, MAC-NELS, CMS / THE M CORP / Eastern Air / Penanshin, DSV) are detected and read with that reader automatically. Other known forms: Heung-A arrival notices / B/L copies + ATTACHED RIDER, MARK&NO. sheets, AMENDMENT tables (Cargoport / ASIA ATLANTIC / INITIAL / Handle / UNION WORLD / UNICORN), free-text amendment letters, letter forms with a saved profile (H&FRIENDS, BILLION, LOGISALL, YOUR SHIPPING, ACE INTER FREIGHT, Gebruder Weiss, HIPPO, VISION, T.V.L.), MAC-NELS, THE M CORP/CMS, DSV, sideways pages, scanned/outlined pages via OCR. This is also where every NEW ENTER form the user shows is saved. Use whenever the user says \"check-edi input\" / \"edi input\" for a bundled ENTER, \"ENTER ชุดรวม\", \"ENTER หลายฟอร์ม\", \"เทียบ MF กับ ENTER ชุดรวม\", \"จำฟอร์มนี้\" / \"เพิ่มแบบฟอร์ม ENTER\", or points to a job folder whose input/ ENTER PDF bundles several forwarders. Separate from manifest-enter-edi-check: never edit that skill or other forwarders' checks for this."
---

# ENTER ชุดรวมหลายฟอร์ม × MF → EDI.xlsx

An ENTER that is **one PDF combining many forwarders' forms**. The program identifies each page's form,
reads it with that form's reader, and compares everything it actually sees with the MF (MANIFEST).
The user keeps adding ENTER forms; **every form is remembered here** (section "Form registry") — this
skill is the single home for all of them.

**Rules from the user**
- Do not change or upload anything for the other forwarders' EDI checks (`manifest-enter-edi-check`,
  its websites/repos). Fix and extend only this skill's `scripts/bundle_edi.py` (and its own
  `scripts/build_edi.py` copy).
- Do not bring up the old job/website this skill started from; it was deleted on purpose (2026-11-06).
- **User rule for this check (2026-11-27), overrides the global ENTER-only / SHED rules here:** "การตรวจ m/f
  แยกละฟอร์ม เราจะไม่ตรวจจับ SHED และหาก BL ไหนไม่มี ENTER ให้ใส่สีเหลืองช่อง B/L NO. เรียงลำดับตาม m/f" →
  - **SHED is not checked**: the SHED column just shows the MF value (no colour, no note, not counted);
    do not ask the user for a SHED.
  - **Every MF B/L is listed, in MF order** (split B/Ls `…A/…B` right after their main B/L).
  - An MF B/L with no ENTER: **only its B/L NO. cell is dark yellow `FFD966`** (user: "ขอเหลืองเข้ม", 2026-11-27; web class `yeldk`) (the rest of the row stays white; the whole
    row is orange only when its container points at an unreadable scan page). The TOTAL row's MF sums leave
    those B/Ls out, and its note says how many.
  - An ENTER B/L missing from MF is still a red critical row at the end.
- **User rule (2026-11-27, same day):** "UN CLASS ตรงกันไม่ต้องแจ้งเตือน ; TEMP ต้องกัน ไม่ต้องแจ้งแต่หาก m/f VENT
  ไม่มีต้องแจ้งเตือน ; … ถูกต้องตามกฎ = แสดง - (แสดงค่าจริงเฉพาะตอนไม่ตรง/ต้องตรวจ) … แค่ตัดเรื่องตรวจ shed ออกไปเท่านั้น" →
  - **DG**: ENTER and MF give the same UN numbers (and the same CLASS when both sides give one; CLASS only = same
    CLASS) → white, no warning, not counted; cell shows the value. One side only / different = red as before.
    MF / ENTER `CL9` is read as CLASS 9 (`find_dg`).
  - **REEFER**: same TEMP on both sides → white, cell `-` (was `0°C  (VENT 15CBM/H)` until the later update); but **MF text has no `VENT`** → red
    `… (MF ไม่มี VENT)`. One side only / different TEMP = red as before.
  - **TRANSIT/TRANSHIPMENT**: both sides give the same transit and MF ประเทศปลายทาง is correct and CARGO MOVEMENT
    is not wrong → `-` (note keeps `TRANSIT ตรงกันและถูกต้องตามกฎ (…)`); otherwise red as before.
  - SHED / ประเทศปลายทาง / CARGO MOVEMENT / TAX ID: correct = `-` (SHED is always `-`, not checked); real value
    only when wrong / to check — same as the other forwarder checks.
- **User rule (2026-11-27, later): SHED is checked for one case only — DG at DISCHARGE LAEM CHABANG must be 2826.**
  "เพิ่มกฎการตรวจ SHED 2826 DG สำหรับ DISCHARGE LAEM CHABANG เท่านั้น หาก M/F ปรากฏ dg class un แล้วใส่ shed ไม่ถูกให้
  แจ้งเตือน เพิ่มหัวข้อ SHED (MANIFEST) และหาก ENTER ไม่มี class un ให้แจ้งเตือน" →
  - **Update (same day): "ตอนนี้ให้ว่าง SHED อื่นไปก่อน ตรวจเฉพาะ DG ย้าย SHED (MANIFEST) ไปไว้ก่อน DG"** — the column
    sits between REEFER TEMP and DG (`_OLD_TO_NEW` 24→23, DG 22→24, TRANSIT 23→25) and is **blank** for every B/L
    that is not checked; only DG at LAEM CHABANG shows the MF SHED (red if ≠ 2826). Regression by column name: only
    SHED cells changed in job-01…07.
  - **Update 2: "ถ้ามันถูกต้องและตรงให้ใส่ - ถ้ามันไม่ตรงก็แจ้งเตือน"** — DG SHED 2826 correct → `-`; DG that matches
    (CLASS/UN complete) → `-`; wrong/missing → red with the values. REEFER that matches (TEMP same + MF has VENT) also shows `-` ("REEFER ที่ตรงกันให้ใส่ - ด้วย").
- **Every matching pair = `-` (2026-11-27, "ช่องที่ตรงกันอื่นๆ ให้ใส่ - ด้วย"):** CNEE, CNTRS NO., TOTAL PKG, PKG.,
  G.W., MEAS., MARKS, DESC. — when ENTER and MANIFEST match, both cells show `-` (`build_report`, before `_reorder`);
  a mismatch keeps both values (red / orange). Rows read by OCR keep their values (check by eye); MF B/Ls without
  ENTER are unchanged; notes (red box / extra text / unboxed numbers) stay. `annotate_report` no longer writes the
  full ENTER MARKS/DESC over a `-`. Regression: every changed cell in job-01…07 became `-`, no fill changed.
- **STATUS (2026-11-27):** "STATUS (ENTER) / STATUS (MANIFEST) status เดียวกัน เช่น ลากตู้, CY ตรงกับ m/f ที่เป็น CY
  เปลี่ยนเป็น -" → when both sides give a STATUS and they mean the same (CY,CY/CY,FCL,ลากตู้=CY ; LCL,ขน(ส่ง)=LCL ;
  CFS,LCL/CFS,เปิดตู้=LCL/CFS) **both cells show `-`**; the customer's own wording (ด่วน / HOT BOX / เร่งเปิดตู้) moves
  to the note `STATUS ตรง (CY) — ENTER: …`. Different = red with both values; ENTER without STATUS still shows
  `(ไม่ระบุใน ENTER)` + the MF value. Regression: only the two STATUS cells (+ note) changed in job-01…07.
  - **Update 3: header renamed to `SHED (DG_M/F)`** ("SHED (MANIFEST) เปลี่ยนเป็น SHED (DG_M/F)").
  - (Before the update) the column was **SHED (MANIFEST)** and always showed the MF value. When MF port of discharge is LAEM CHABANG
    and MF has DG (CLASS or UN), SHED ≠ 2826 → red `… (ต้องเป็น 2826 — DG LAEM CHABANG)`; every other B/L is not
    checked (white). TOTAL cell `SHED DG ไม่ตรง n B/L`.
  - DG passes only when ENTER gives **every** CLASS and UN the MF gives, with the same values; a missing one →
    red `DG ENTER ไม่ระบุ CLASS / UN`.
  - `DG_RE` / `find_dg` now read `CLASS: 9`, `UN: 3077` and `UN: 3077HS CODE` (colon / glued HS) — MF of SWDN
    1001886 was not seen as DG before.
  - Regression: no wrong DG SHED in job-01…07 (job-02 / job-07 DG at Laem Chabang all 2826, job-06 DG is BANGKOK
    0121 = not checked); newly seen DG B/Ls (CLASS:/UN: with colon) mostly pass; OCR / one-sided ones warn.
- **User rule (2026-11-27, SWDN 1100797):** "enter ไม่ได้ใส่กรอบแสดงว่าไม่ใส่ข้อมูลนั้น ถ้าหากมีกรอบแสดงว่าเอาส่วนนั้น"
  → on a Heung-A B/L page (BLFORM), when a red box covers one of the number cells (TOTAL PKG / G.W. / MEAS.),
  the number cells **outside** any red box were not given by the customer: they are not checked (ENTER value
  replaced by the MF value; note `MEAS. ไม่อยู่ในกรอบแดงของลูกค้า = ไม่ได้แจ้ง ไม่ตรวจ (ENTER พิมพ์ 55.000)`).
  **No red box on the page (e.g. FreeText `ตามหน้า BL  CY`) = every value on the B/L page is the customer's
  data and is checked in full** (user confirmed 2026-11-27 with `HASLC…797_ตามหน้า BL .pdf` / `_CY.pdf`: MEAS. 55
  vs MF 0.000 stays red). A short FreeText ending in `CY` / `CY/CY` / `FCL` / `LCL` (≤ 30 characters) is the
  customer's STATUS (regression: 8 STATUS cells in job-02 / job-06 went from "(ไม่ระบุใน ENTER)" to CY, all = MF).
  **Boxed MARKS/DESC = what the customer declared (2026-11-27, SWDN 1002618 "ENTER มีตีกรอบ คือสิ่งที่ลูกค้า ENTER
  ไม่ตรงกับ M/F ขึ้นเตือนด้วย"):** the MF must **equal** the boxed text (before: it was enough that the MF contained
  it). Not counted on either side: TEMP / VENT / HS CODE, DG tokens (`_DG_TOK`: UN NO: 2265, CLASS:8+6.1, CL9,
  9/3082 — checked in DG), the MF package prefix `(15 METAL BOXES)`, the transit sentence (`_TR_TOK`, checked in
  TRANSIT). Boxes over **part of a line** are read word by word (`parse_blform(words=page.get_text('words'))`, word
  centre inside a red box that overlaps a MARKS/DESC line) — p.9 boxes half of one line + the next line.
  Regression: 8 B/Ls in job-01…07 now warn DESC, each a real difference (MF lists more products / text than the
  box: a second product, a grade / model suffix, quantity and packing lines);
  other changes are notes only (newly found partial boxes that match).
  **MEAS. outside the box (2026-11-27, SWDN 1002364 "enter ไม่ได้ตีกรอบ แต่ M/F เป็น 60.000 MTQ มันผิดต้องแจ้งเตือน"):**
  the customer did not declare MEAS., so the MF must be **0.000** — ENTER MEAS. is set to 0 and compared: MF 0.000 =
  `-`, MF 60.000 = red; note `MEAS. ไม่อยู่ในกรอบแดงของลูกค้า = ไม่ได้แจ้ง → MF ต้องเป็น 0.000 (ENTER พิมพ์ 60.000)`.
  TOTAL PKG / G.W. outside the box are still not checked (MF value used). Regression: only SWDN 1002364 (MF 60) and
  1002539 (MF 40) turned red; every other unboxed MEAS. in job-01…07 has MF 0.000.
  `red_boxes()` now also finds red boxes drawn as plain red lines in the page (`get_drawings`, no fill,
  > 25×10 pt), not only Square annotations. Regression: 50 B/Ls in job-01…07 changed, all by this rule (mostly
  MF MEAS. 0.000 with an unboxed ENTER MEAS. → pass); no new mismatch.
  - Regression 2026-11-27: only summary / REEFER / DG / TRANSIT / SHED / note cells changed in job-01…07;
    the warnings left are one-sided DG/TEMP, MF without VENT (job-01 x2, job-04 x1) and transit whose MF rule
    lines are missing (job-04).
- Other global rules still built in: **WOODEN CASE ≠ CASE**; Laos = **(LA)**.

## GitHub repo (2026-11-27)

The program is published at **https://github.com/Rattanao/Various-forms** (local clone `Downloads/Various-forms`,
branch `main`; guide page https://rattanao.github.io/Various-forms/ — live 2026-11-27; Pages serves `main` / root,
so the page is the root `index.html` (+ `.nojekyll`); keep `docs/index.html` identical). It holds `SKILL.md`, `scripts/`, `README.md`,
`run-check-edi.bat` (asks SHED, writes `output/EDI-<date-time>.xlsx`), `requirements.txt`, `.gitignore`
(input/, output/, tests/, *.xls, *.xlsx, *.pdf never go up).

**Web vs PC test 2026-11-27 (SWDN, job-07):** every text page identical to the PC report; only the 5 scanned
(orange) B/Ls differ (Tesseract vs Windows OCR) plus the totals that follow from them. Found on the way: Pyodide's
PyMuPDF 1.26.3 keeps double spaces inside a line (`Gross Weigh  Volume`) where the PC's 1.28.2 has one, so
`apply_profile` now matches header titles with spaces collapsed — any new matching on raw line text must do the same.
Test recipe: copy the clone's `index.html`, `app/`, `scripts/` + `testdata/{MANIFEST.xls,ENTER.pdf,pc.xlsx}` into a
scratch `site/`, serve it on localhost, `CheckVarious.run(mf, en)`, then compare `/work/output/EDI.xlsx` with pc.xlsx.

**Online checker (the link to send to friends): https://rattanao.github.io/Various-forms/** (2026-11-27).
`index.html` + `app/app.js` run the *same* `scripts/bundle_edi.py` + `build_edi.py` in the browser with Pyodide
0.28.3 (`loadPackage(['pymupdf','xlrd','micropip'])` + `micropip.install('openpyxl')`); nothing is uploaded.
Scanned pages: pass 1 runs with `bundle_edi.BROWSER_OCR = {}` and collects `OCR_WANTED` page numbers; the page
renders each at 300 dpi gray (Python `_png`), OCRs it with Tesseract.js 5 (`blocks:true`), splits each Tesseract line
at word gaps > 1.1 × line height and at lone `:`/`|` (Tesseract joins columns that Windows OCR keeps apart), maps
`©`→0, then pass 2 runs with `BROWSER_OCR = Map{page: lines}` (number keys!). Test 2026-11-27 on job-05: text pages
identical to the PC run, 16/16 B/Ls found; scan pages differ (Tesseract is weaker on some fields, better on others) —
they stay orange. Big scan-heavy bundles are slow online; the PC version (Windows OCR) is the accurate one.
Look = the other forwarder pages (M INTER etc.): navy header + `app/hal-logo.png` + "Project by …", one drop zone
with MANIFEST/ENTER chips, SHED (one number or LCL/CFS + CY), result card with VESSEL/ENTER/OCR info, count chips,
colour legend, search + "only mismatches", table read back from the EDI.xlsx (`_table`), download button.
`window.CheckVarious.run(mfFile, enterFile, {shed})` / `.boot()` / `.py()` are test hooks. Bump `app.js?v=` in
`index.html` when app.js changes. `guide.html` (= `docs/index.html`) is the guide + ZIP download. **To update the repo after changing this skill:** copy `SKILL.md` + `scripts/*.py|ps1`
into the clone, run the sanitizer (customer names used as examples → ABC/XYZ, real B/L digits → 0, real
container digits → 0000000; regexes are untouched), check `ast.parse`, re-run one test job from the clone and
compare with the skill's own output (must be 0 differences), then commit + push. Customer files never go up.

## Files

- `scripts/bundle_edi.py` — the bundle reader (page types, form profiles, OCR) + report post-processing.
- `scripts/build_edi.py` — this skill's own copy of the shared checker (`parse_manifest`, `build_report`,
  MAC-NELS / CMS / DSV page parsers). Loaded before any other copy, so the folder works alone on another PC.
  Changes that live **only in this copy**: MF PDF header labels dropped and `1)…4)` rule lines go to
  column 15; the B/L prefix guess needs ≥5 digits (2026-11-03); `canon_pkg` keeps `WOODEN CASE` and the
  PACKAGING check never matches WOODEN CASE with CASE (2026-11-06); DSV layout G takes the B/L from
  `MB/L NUMBER :` when the B/L CHANGED cell is empty, and `find_temp` reads MF `TEMP: +2CVENT: 25 CBM/H`
  (C glued to VENT) (2026-11-24); report rows sorted by B/L so split B/Ls stay together, and `canon_pkg` knows
  PAIL / TRAY / TANK (2026-11-25). Every such change is marked `[bundle-enter-check]` in the code.
- `scripts/verify_trace.py` — proves the comparison is real (below).
- `scripts/win_ocr.ps1` — Windows' built-in OCR for pages with no text layer.

Requirements: Python 3.10+, `pip install pymupdf openpyxl xlrd`. OCR needs Windows 10/11.

## Run

From the job folder (the one that contains `input/`):

```bash
python ~/.claude/skills/bundle-enter-check/scripts/bundle_edi.py --name EDI.xlsx
```

- Finds `input/MANIFEST*.xls` (or MF PDF) and `input/ENTER*.pdf` automatically (or `--manifest` / `--enter`).
- No SHED check (user rule 2026-11-27). `--shed …` is accepted but ignored. Old behaviour only if the user
  asks: `--check-shed --shed 0141` (or `--shed-lcl X --shed-other Y`).
- Every MF B/L is listed in MF order; the console lists the MF B/Ls with no ENTER (yellow B/L cell).
  `--enter-only` = old behaviour (only B/Ls in ENTER). `--all-mf` is still accepted (now the default).
- Writes `output/<name>`; it never overwrites an existing report (pick a new `--name`). If Excel has the
  file open the save fails — use another name/outdir.

**Trace check (prove the comparison is real):**

```bash
python ~/.claude/skills/bundle-enter-check/scripts/verify_trace.py input/MANIFEST.xls input/ENTER.pdf
```

Every value the report uses is looked up again in the raw files (ENTER values on the page(s) listed for
that B/L, MF values in the B/L's raw rows); "not found" should be 0. Values stated once per document
(sheet-top STATUS, header container) list that page too. Reads only an .xls MF. Values fixed on purpose
(Thai font repair, OCR) do not trace literally — explain them.

Relay to the user: DG / REEFER / STATUS first, then MF B/Ls with no ENTER (yellow), MEAS. 0.000 in MF, MARKS/DESC, then pages that
could not be read.

## How the bundle is read (page types)

| Type | Recognised by | Reader |
|---|---|---|
| BLFORM | `SHIPPER'S LOAD` (Heung-A ARRIVAL NOTICE / B/L COPY, also no-space text) | coordinates relative to the SHIPPER'S LOAD row |
| RIDER | `== ATTACHED RIDER ==` | its marks replace `AS PER ATTACHED RIDER` |
| MARKNO | `MARK&NO.` + `CNTR :` + `B/L NO. :` (2 B/L per page) | own |
| TABLE | header row CONSIGNEE / QUANTITY / DESCRIPTION / GROSS WEIGHT | columns = header x − 45; split B/L (`…104` + `9`) joined; STATUS/container stated once apply to all rows; containers also from the HB/L column |
| LETTER | anything else with text | form profile if one matches, else label/value (`CNEE`, `CONTAINER NO`, `Status` …) or a small MARKS/PKGS/DESCRIPTION table; MARKS/DESC mismatch = orange |
| SHARED_MACNELS / CMS / DSV | MAC-NELS ATTACHED SHEETS / INBOUND CARGO MANIFEST / AMENDMENT OF MANIFEST + B/L CHANGED | only those pages go to `build_edi.parse_enter` |
| SCAN | no text layer | OCR (below); always orange, check by eye |

`page_lines()` / `page_annots()` first turn **sideways-drawn pages** (text direction (0,±1) while
`page.rotation == 0`) back to reading orientation.

## Form registry — every ENTER form remembered

When the user shows a new form (screenshot with arrows, or a PDF page), add it here **and** in the code,
then run the regression (end of this section). Prefer **one `FORM_PROFILES` entry** over a new parser.

| Form (forwarder) | How it is read | Added |
|---|---|---|
| Heung-A ARRIVAL NOTICE / B/L COPY + ATTACHED RIDER | BLFORM + RIDER | 2026-10-26 |
| MARK&NO. / NO. OF PKGS / DESCRIPTION OF GOODS sheet | MARKNO | 2026-10-26 |
| Cargoport, ASIA ATLANTIC, INITIAL FREIGHT, Handle Inter AMENDMENT | TABLE | 2026-10-26 |
| UNION WORLD AMENDMENT (sideways page) | TABLE + rotation fix | 2026-11-03 |
| UNICORN (single-B/L table) | TABLE + profile | 2026-11-03 |
| MAC-NELS ATTACHED SHEETS | SHARED_MACNELS | 2026-10-26 |
| THE M CORP / CMS INBOUND CARGO MANIFEST (+ cover letter "AS PER ATTACHED SHEET") | SHARED_CMS + LETTER | 2026-10-26 |
| DSV AMENDMENT OF MANIFEST | SHARED_DSV | 2026-10-26 |
| Free-text amendment letters (Lux/Nippon, APC, Freight Rangers, Uni-Global, Together …) | LETTER label/value | 2026-10-26 |
| H&FRIENDS, BILLION, LOGISALL (KAE LAT), YOUR SHIPPING, ACE INTER FREIGHT, Gebruder Weiss, VISION, T.V.L., HIPPO | LETTER + `FORM_PROFILES` | 2026-11-03 |
| Scanned / outlined-text pages | SCAN + OCR | 2026-11-03 |
| **Sinotrans** (MB/L NO. letter) | SHARED_SINOTRANS → layout L | 2026-11-06 |
| **LX Pantos** (AMENDMENT OF MANIFEST OF) | SHARED_LXPANTOS → layout J | 2026-11-06 |
| **Freight Links** (IMPORT CARGO MANIFEST + cover page) | SHARED_FREIGHTLINKS → layout M | 2026-11-06 |
| **Shipco** (CUSTOMS IMPORT MANIFEST) | SHARED_SHIPCO → layout E (+ customer's typed STATUS, which layout E alone misses) | 2026-11-06 |
| **Planet / UNITEX** (B/L CHANGE NO. + HB/L NO + `CNTR NO.:`) | SHARED_PLANET → layout K | 2026-11-06 |
| **Union World** (B/L CHANGE NO. + HB/L NO + `STATUS =`) | SHARED_UNIONWORLD → layout I | 2026-11-06 |
| **AGN / Handle Inter** (B/L CHANGE NO., no HB/L NO) | SHARED_AGN → layout D (+ every container on a single-B/L page) | 2026-11-06 |
| **Able / Quality** (MARKS & NOS / DESCRIPTIONS OF GOODS / GROSS WT.& MEAS) | SHARED_LAYOUTB → layout B; a page with this header starts a new document | 2026-11-06 |
| **ORCA FORWARD** (letter "ARRANGE TO AMEND MANIFEST AS WELL AS DELIVERY ORDER", one B/L per page: `STATUS : CY/CY` on the FROM line, `B/L NO.:` + `CONTR.` (2nd container printed under the 1st), table MARKS AND NOS. / NO.OF P'KGS / DESCRIPTION OF GOODS, `GROSS WT. =`, `CBM. =`, `CNEE :`) | LETTER label/value — read correctly with no code change (tests/job-03). ORCA sends **one PDF per B/L** (`sfimmblcover_attach01*.pdf`): merge them into one ENTER.pdf in the scratchpad and pass `--enter`. Web page: repo `Rattanao/Orca` (`docs/`), synced into the combined site `Rattanao/Check-MNF-Forwarder` as `tools/orca/` by `sync.sh` | 2026-11-11 |
| **STAR CONCORD** ("Sea Freight Summary Manifest": BILL NUMBER / WGT / VOL / PKG / CONSIGNEE / ORIGIN/DEST / MARK / GOODS DESCRIPTION, `STATUS :`, `Cargo movement = 7`) | LETTER + profile; `MOVEMENT` rx gives the customer's cargo-movement code | 2026-11-24 |
| **NEW HANYUAN LOGISTICS** (MARKS / NO. OF PKGS. / DESCRIPTION, `CNEE :`, `… KGS`, `… CBM`, `STATUS :`) | LETTER + profile | 2026-11-24 |
| **NIPPON EXPRESS** (SHIPPING MARKS / QUANTITY & DESCRIPTIONS / CONSIGNEE, `PKG :` `W/T :` `M/M :`, `SERVICE TYPE :`) — B/L suffix letter typed as a separate red FreeText after the B/L (`…0966` + `A`); page "ATTACHED SHEET (DESCRIPTION)" fills the description of the B/L already read | LETTER + profile (`skip` option) | 2026-11-24 |
| **SINTOP INTERNATIONAL** (header line `MARKS :` then columns by x: marks / pkgs / description / G.W.+M.) | LETTER + profile `bands` | 2026-11-24 |
| ARRIVAL NOTICE **scanned** (no text layer) | SCAN → OCR → same BLFORM reader on OCR lines (garbled labels fixed, `PACKAGESi SHIPPER'S LOAD` split); falls back to the old OCR reader when pkgs/G.W. are not found (tilted scans) | 2026-11-24 |
| ARRIVAL NOTICE with the **scanner's own OCR text layer** (B/L `HASLS0000■000000`, `SHIPPERIS LOAD`) | `■` resolved against MF B/Ls (one match only) → BLFORM; row orange like OCR | 2026-11-24 |
| ATTACHED RIDER `<< Description Information >>` | replaces `AS PER ATTACHED DESCRIPTIONS` of that B/L | 2026-11-24 |
| **MAXTOP LOGISTICS** ("AMENDMENT OF MANIFEST", MBL NO / HBL NO; table MARKS / NO OF P’ KGS. / DESCRIPTION OF GOODS; `CONSIGNEE NAME:`, `STATUS : CY/CY`, `CONTAINER NO :`, `GW : … KGS`, `CBM : … CBM`) | LETTER + profile | 2026-11-24 (job-05) |
| **SINO LOGISTICS** (same "ARRANGE TO AMEND MANIFEST AS WELL AS DELIVERY ORDER" letter as ORCA, `CNEE` + `: name` on the next pt) | LETTER label/value; a line starting `:` ends the goods table | 2026-11-24 (job-05) |
| ECONIC LOGISTICS letter, H & FRIENDS, "WE SHOULD VERY MUCH APPRECIATE … AMEND THE CONSIGNEE AND DESCRIPTION" letter | read with no change (job-05) | 2026-11-24 |
| ARRIVAL NOTICE / B/L COPY whose scanner text layer is unreadable (`HASLKO000LOO00L0`, no MF B/L exactly on the page, one MF B/L within 2 characters) | page re-read with Windows OCR (`_needs_ocr`) → BLFORM reader | 2026-11-24 (job-05 p10) |
| Letter fixes from the user's screenshots (2026-11-25, job-06): **CARGO NETWORK** `5,251` / `CARTONS` on two lines (header `MARKS & NOS. QUANTITY` is one text line) and **SVI/ISC** `19` / `Pallets` (header `NO.OF` / `PACKAGES` on two lines) → a bare number with a package word right below is the count; 2nd header row `WEIGHT(KGS.) MEAS.(CBM)` / `GOODS (KGS) (M3)` is not goods; **ABC/XYZ** `CONSIGNEE:` with a far-right `QUANTITY 38 CASES` → right-side value must be within 100pt, else the line below; `QUANTITY 38 CASES` / `Quantity : 1105 CARTONS` = package count; `MEAS. 29.6200 CBM.`; **WISMA** `11,080.000 KGM.` (KGM = KGS), `(TOTAL : NINETY …)` ends the table; **CARGO SERVE/ORCA-type** `GROSS WT. = 9,550.000 KGS.` printed under the table wins over `IN 900KG (NET)` in the description, `10 IBCS` (IBC; MF `WA (INTERMEDIATE BULK CONTAINER(s))` = IBC); **Masterwork** `AS THE SHIPMENT TERM…` ends the table | LETTER reader rules | 2026-11-25 |
| **KUEHNE + NAGEL** "CORRECTION MANIFEST" (monospace, columns separated by long spaces; page 1 `MBL NO. : HASL…`, `CONSIGNEE` + name below, `MOVEMENT : CY/CY`; page **ATTACHMENT**: one block per container `SKHU0000000  1 40' HC  CONTAINER SAID TO CONTAIN:  8214.350  66.514` / `SEAL NO.…  1730 CARTON(S)` / `ABC  ABC ARTICLES` / `HS CODE`, `TOTAL: 13252 CARTONS`, total row `4 40' HC … 33223.217 264.396`). No web checker existed for it. | own reader `parse_kn_page` (MARKS = left column words, DESC = right column words, both de-duplicated; containers from every block; pkgs from `TOTAL:`; G.W./MEAS from the total row) | 2026-11-25 (job-06 HASLS00000000000) |
| **GENIUS INTERFREIGHT** (`CONSIGNEE NAME :` centred, name on the next line; MARKS / NO.OF PACKAGE / DESCRIPTION OF GOODS in 2 sub-columns / WEIGHT, `3,783.00` `KGS` / `27.600` `CBM` as separate pieces, Thai note `*** ชื่อผู้รับมัดจำตู้ …` inside the table) | LETTER + profile `bands` + `skip` | 2026-11-25 |
| **SHIP MASTER** AMENDMENT table (B/L NO / MARK & NO / QUANTITY / DESCRIPTION / CONSIGNEE / `G.W.` / HB/L NO, `***ลากตู้***` `1X20'DC` `CY/CY` at the end of DESCRIPTION, `CREATED BY :` footer) | TABLE (`G.W.` accepted as weight header; those 3 lines → STATUS, not DESC; `CREATED BY` ends the table) | 2026-11-25 |
| Letter page 2 with no B/L (`Page 2 of 2`, e.g. `CY-CY (ลากตต)`) | STATUS / containers go to the B/L of the page before | 2026-11-25 |
| **M INTER-CORP LOGISTICS** INBOUND CARGO MANIFEST (`** STATUS: CY/CY **` in DESC, extra containers on lines under `CNTR NO:`, MF `100%` cells) | SHARED_CMS → layout F of the re-synced build_edi.py — 9/9 B/Ls pass (1022076, A–H), same as the user's M-INTER web page | 2026-11-25 (job-06) |
| **LEO GLOBAL** (wide table D/O NO. / MARKS & NUMBERS / QUANTITY / DESCRIPTION / CONSIGNEE / WEIGHT/VOLUME / HB/L NO., many sub-B/Ls over 5 pages, B/L printed `HASLC000000000` + `29A` on the next line) | TABLE (WEIGHT/VOLUME accepted as G.W. header; description text that spills into QUANTITY moved back to DESC) | 2026-11-24 (job-06) |
| AMENDMENT table "B/L CHANGE NO. / … / HB/L NO" without CNTR/STATUS lines (SHENZHEN JW agent: `HASLC0000000` + `2029A`) and **BUGATTI** (`Page 1 of 1` above the table, `MARK & NUMBER`) | TABLE — B/L tail of up to 6 chars on the next line joined while the B/L is shorter than 16 | 2026-11-24 (job-06) |
| **FUKUYAMA GRAND / INTERCOSPED** ("WE SHALL BE OBLIGED IF YOU KINDLY AMEND THE FOLLOWING MANIFEST…", MARKS / NO.OF PACKAGES / DESCRIPTION OF GOODS / GROSS WEIGHT / MEASUERMENT, `CNEE :`, `STATUS * CY *`) | LETTER + profile | 2026-11-24 (job-06, also job-02 p5) |
| Letter with **two B/Ls** on one page (ABC CHEMICAL: `ENTER B/L …5764` … `____` … `ENTER B/L …5764A`, labels `NO OF P’KGS:`, `G.W.:`, `MEASUREMENT:`, `DESCRIPTION OF GOODS :` value on next line) | `_letter_blocks` splits by the first line of each B/L (needs ≥2 consignee labels) | 2026-11-24 (job-06) |
| Letter with the **B/L NO. left blank** (EXTRA MARITIME) | matched to an MF B/L without ENTER by G.W. + (pkgs or ≥2 consignee words), only when unique; row orange + note | 2026-11-24 (job-06) |
| Blank separator pages in scanned bundles | kind BLANK — skipped, listed, not "unreadable" | 2026-11-24 (job-06: 54 pages) |
| Scan pages that are copies of a B/L already read from text, or page `2/3` "MARKS (ต่อ)" after a letter | page added to that B/L, not "unreadable" | 2026-11-24 |
| Scanner text layer B/L with 0→8 errors (`HASLC00000000000`) | lookalike digits (0/8, 6/8, 1/7…) cost 0.5 in the B/L distance → re-read by OCR | 2026-11-24 |
| OCR B/L unreadable or ambiguous (`HASLce1261æe11e`, `HASLC0000000000THSC`) | containers on the page and the MF consignee name decide; lower-case `HASLc` accepted | 2026-11-24 |
| ARRIVAL NOTICE with count glued to the label (`8 TANK SHIPPER'S LOAD…`) | split before reading (used to crash) | 2026-11-24 |
| Scanned ATTACHED RIDER `<< Mark Information >>` | OCR → its marks replace `AS PER ATTACHED RIDER` of the same B/L (B/L on the rider may be read as `HASL1022611øe793` — letter after HASL lost) | 2026-11-24 (job-05 p18) |
| Scanned ATTACHED RIDER `<< Description Information >>` | OCR → `parse_rider` on the OCR lines; its description replaces `AS PER ATTACHED DESCRIPTIONS` | 2026-11-27 (job-07 p18) |
| **CK LINE (THAILAND)** AMENDMENT (one-row table B/L NO. / D/O NO. / CONSIGNEE / MARKS & NUMBER / QUANTITY / DESCRIPTION / `G.W` / `Measurment`; values left-aligned under centred titles; total row ~200pt lower) | LETTER + profile `CK LINE` with `'gap': 60` (new profile key: stop reading at a vertical gap) — before, CNEE was read as `MARKS & NUMBER` | 2026-11-27 (job-07 p12, also job-02 p3) |
| **EXTRA MARITIME** letter with a table (Marks & Number / Nbr of Pkgs / Description of Goods / `Gross Weigh Volume` = one title over `12,600.00` and `30.000 CBM`) | LETTER + profile — before, G.W. was read as 30 | 2026-11-27 (job-07 p14) |
| **RICH FREIGHT FORWARDER** `*** ATTACHED LIST ***` (Page 2/2: MARKS / NO.OF PACKAGES / DESCRIPTION OF GOODS / GROSS WEIGHT / MEASUERMENT, then `CNEE :`, `DO NO. : <B/L>A`, customer FreeText CTN/STATUS box over the CNEE line) | LETTER + profile; a form line that starts left of a FreeText box and is not its text is kept (`in_annot`) | 2026-11-27 (job-07 p19) |
| **UNICORN GLOBAL LOGISTICS** AMENDMENT with `B/L CHANGE NO.` (no HB/L NO) | no longer routed to SHARED_AGN (layout D misread it) — TABLE + the UNICORN profile; `***јѥдшѬҖ***` = ลากตู้ | 2026-11-27 (job-07 p26) |
| **Sahathai Terminal** (สหไทย) "pre-lading manifest" letter (Consignee Name / No of package / Description of goods / G.W. / Mark & No.; `Meas.` printed under G.W.; `(HEAD OFFICE)` on the line under the name) | LETTER + profile `SAHATHAI TERMINAL`; `(HEAD OFFICE)` / `(BRANCH …)` joined to the name | 2026-11-27 (job-07 p27) |
| MF that prints the **container count `X1` / `X2`** as the first line of the description column (SAWASDEE DENEB 2612S) | `X<n>` removed from MF DESC when n = the B/L's container count | 2026-11-27 (job-07) |
| ENTER package `FLEXIBAG` / `FLEXITANK` | = TANK (MF `TK (TANK)`) | 2026-11-27 (job-07 p10) |
| STATUS typed `ลากตุ้` (ุ instead of ู) | repaired to ลากตู้ = CY | 2026-11-27 (job-07 p32) |
| ENTER `CL9` / `CL 9` / `CL.9` vs MF `CLASS 9` (user: "CL ก็คือ CLASS ถูกต้อง") | `apply_bundle_rules` turns `CL n` into `CLASS n` in ENTER DESC/MARKS (also the red-box text) only when that B/L's MF has `CLASS n`; ENTER DG gets the CLASS from it. Regression: job-07 13 cells (1002364 DESC now matches, 3 DG B/Ls show CLASS 9), job-01…06 0 | 2026-11-27 (job-07 p6, p7, p16) |
| Scanned ARRIVAL NOTICE: weight row 4–5pt off the SHIPPER'S LOAD row, `24,016.000 KGS :` (column rule read as `:`), `55. eoeo` / `49, øee` (space inside a number), `CBM` unreadable, `& WEIGHT` split off | OCR row tolerance 6pt (OCR only), leading/trailing `:` dropped, number pieces joined, a bare decimal right of the weight = MEAS., `KGS 89.0000 CBM` in one piece → number after KGS | 2026-11-27 (job-07 p24, p16) |

### Forwarder forms that already have a dedicated reader (automatic, 2026-11-06)

`FORWARDER_PAGES` in bundle_edi.py holds one signature per forwarder (same tests as that forwarder's own web
page). A matching page goes, with the pages next to it of the same form, to the dedicated reader in this
skill's `build_edi.py` (`parse_shared`). `FORWARDER_MEMBER` pulls in continuation / cover pages that do not
repeat the signature (Able/Quality page 2…, Freight Links cover). B/Ls printed in two pieces (`…2381` + `3A`)
are kept when they are MF B/Ls or their base number is on the page. Signature-routed B/Ls are never marked
low-confidence. Pages the skill already reads with its own tested readers (ARRIVAL NOTICE, MARK&NO. =
Gateway, Cargoport / ASIA ATLANTIC / INITIAL tables, MAC-NELS, CMS family, DSV) are not rerouted.

To add another forwarder that has a dedicated reader: add its signature to `FORWARDER_PAGES` (and
`FORWARDER_MEMBER` if only page 1 carries it), add its sample job to `scripts/test_forwarders.py` `JOBS`, run it.

**Keep `build_edi.py` current:** it is the newest `manifest-enter-edi-check` checker (re-copied **2026-11-25**,
with M INTER-CORP in layout F, LEO GLOBAL layout N, and the MF `%` cell fix) plus this skill's own changes — every
one is marked `[bundle-enter-check]`: `_PDF_MAN_HDR`/`_PDF_MAN_RULE`, B/L prefix needs ≥5 digits, DSV `MB/L NUMBER`
fallback, `find_temp` C+VENT, `canon_pkg` PAIL/TRAY/TANK, rows `sorted(MAN)`; and the old job name in comments is
written "ENTER ชุดรวม". How to re-sync: `diff -u <upstream backup the copy was made from> <this copy>`, keep only the
hunks that contain `[bundle-enter-check]` (plus the `_PDF_MAN_*` / prefix hunks), `patch` them onto the new upstream
file, rename the old job name in comments, check `ast.parse`, then run the bundle on a job and `test_forwarders.py`.
Lesson 2026-11-25: the user checks some forwarders on that skill's web pages (e.g. https://rattanao.github.io/M-INTER/docs/);
if a B/L passes there but not here, the copy is out of date — re-sync first.

### Forwarder test (`scripts/test_forwarders.py`)

Merges the sample ENTERs of 10 forwarders (`Downloads/check edi-<name>/input`) into one 100-page bundle and
checks every B/L against that forwarder's dedicated reader. 2026-11-06: 138 / 138 B/Ls equal, 0 missing,
0 extra (Shipco STATUS excluded — the bundle also reads the customer's typed STATUS).

### Adding a form profile (`FORM_PROFILES` in bundle_edi.py)

```python
{'name': 'HIPPO LOGISTICS', 'sig': r'HIPPO\s+LOGISTICS',          # text only that form has
 'hdr': {'MARKS': 'MARKS', 'NO.OF PKGS.': 'PKG', 'DESCRIPTION OF GOODS': 'DESC', 'GROSS WEIGHT': 'GW'},
 'rx': {'STATUS': r'^STATUS\s*:\s*(.+)$', 'CONS': r'^CONSIGNEE\s*:\s*(.+)$'},   # 'label : value' outside the table
 'stop': r'^(YOURS|THANK)'},                                       # first line under the table that is not goods
```

- `hdr` = column titles exactly as printed → MARKS / PKG / DESC / GW / MEAS / CONS / BL / QTYDESC
  (quantity+description in one column) / REMARK. Two-row titles are fine. Lines go to the column whose
  **centre** is nearest the title centre.
- Any `n KGS` / `n CBM` cell counts as G.W. / MEAS. wherever it is; a marks line ending in `1,234 CARTONS`
  is split into marks + count; the B/L column's own number is removed, its leftover text goes to MARKS.
- `gap: 60` = stop reading the table at the first vertical gap larger than 60pt (a total row far below, CK LINE).
- `skip` = regex of cells to ignore (labels such as `Meas.` / `KGS.` printed inside the table).
- `mode: 'spaced'` for forms printed as one text line separated by long spaces (YOUR SHIPPING).
- `skip` = regex of table cells to ignore (NIPPON's `PKG :` pieces); `bands` = `{'hdr': <regex of the header
  line>, 'cols': [(right_x, field), …]}` when the column titles do not line up with the values (SINTOP);
  `rx` key `MOVEMENT` = customer's cargo-movement code.
- If the logo is an image, use a sentence of the letter as `sig`.
- Applies to LETTER pages and single-B/L TABLE pages. DESC `AS PER ATTACHED DESCRIPTIONS` is replaced by
  the RIDER page that follows (LOGISALL).

### One test folder for every form + regression after any change

All test jobs live in **one place**: `Downloads/8-Various forms/tests/<job>/input/` (MF + bundled ENTER).
`job-01` and `job-02` are the two real bundles this skill was built on (copied 2026-11-06), `job-03` = ORCA,
`job-04` = SAWASDEE VEGA 2614S (STAR CONCORD, NEW HANYUAN, NIPPON, SINTOP, DSV, scanned arrival notices),
`job-05` = DONGJIN CONFIDENT 0154S (MAXTOP, SINO, ECONIC, 5 scanned pages incl. rotated scans, highlighter, ticks),
`job-06` = SAWASDEE MIMOSA 2611S (292 pages, 266 MF B/Ls, ~40 scans + 54 blank pages, LEO GLOBAL, FUKUYAMA/INTERCOSPED,
BUGATTI, many letter forms; 6 MF B/Ls truly absent from that ENTER). `job-07` = SAWASDEE DENEB 2612S (33 pages,
32 B/Ls; CK LINE, EXTRA MARITIME table, RICH FREIGHT, UNICORN B/L CHANGE NO., Sahathai, 6 scans; MF with `X1`
container-count lines; MF SHEDs 2820 / 2826 / 0332 / 0110). OCR results are cached in
`%TEMP%\bundle_enter_ocr_cache` (by page image), so re-runs of big bundles take minutes, not tens of minutes. When the user
brings a bundle with a new form, copy its `input/` there as the next `job-NN` (customer files stay on this
PC only).

Before editing, copy the **whole** `scripts/` folder somewhere (the base needs `build_edi.py` and
`win_ocr.ps1` next to it, or OCR is silently skipped and the comparison is wrong), make the change, then:

```bash
python ~/.claude/skills/bundle-enter-check/scripts/run_tests.py --base <copy>/bundle_edi.py
```

It runs every test job with `--all-mf --shed 0141` on both versions and lists every cell (value and fill)
that changed; each difference must be explained by the change. Last run 2026-11-06: 0 differences on
job-01 and job-02 (after the WOODEN CASE rule, which alone changed 5 cells in job-01). 2026-11-27 (no SHED check,
MF order, yellow B/L cell): compared per B/L, every difference in job-01…07 is the SHED cell / SHED note / SHED
summary (so the summary now shows the `ต้องตรวจสอบ:` list it used to hide) or the no-ENTER row fills; MF order =
the old order in all 7 jobs; forwarder test 138/138. A position compare now needs both sides in the same mode. 2026-11-24 (OCR/DSV/TEMP
changes): 13 cells, all intended — DSV `SHIPMENT NO / COST CENTER` removed from DESC (job-01), MF `0'CVENT` read
as 0°C (job-02), pages read by OCR no longer listed as unreadable; job-03 0. Forwarder test still 138/138.
Later the same day (job-05 changes): job-02 page 8 scan now read by the BLFORM reader (consignee, MEAS, marks,
desc now match MF) and `TO THE ORDER OF XYZ BANK` → Notify; job-04 report identical; job-03 0; 138/138.
2026-11-27 (job-07 changes, base = the pushed GitHub copy): job-03/04 0; job-02 p3 CK LINE CNEE now correct;
job-01/05/06 only OCR rows improved (G.W. `76,324` / `41,000` / MEAS `25.000` now read, scanned DESCRIPTION rider
used, stray `:` gone). A first try of the `in_annot` change added `D/O : BKK` to a job-01 DESC → narrowed to
lines that start left of the FreeText box. Forwarder test: every forwarder 0 different/missing.

## Comparison rules of this skill

- Containers use ISO 6346 `[A-Z]{3}[UJZ]\d{7}` so SHIPMENT IDs and B/L tails are not taken.
- Lines drawn by a FreeText annotation are recognised by its rectangle (a real marks line that repeats the
  annotation text is kept). Red text the customer typed on the PDF = the customer's real instruction
  (STATUS, DG, REEFER, consignee correction, containers).
- Customer STATUS = any typed / annotation line with STATUS / ลากตู้ / เปิดตู้ / ขนส่ง / CFS (only that line).
  A FreeText STATUS at the top of a shared sheet applies to all its B/Ls; a cover letter "CONSIGNEE AS PER
  ATTACHED SHEET" passes its STATUS to all sub-B/Ls.
- B/L consignee `TO ORDER OF …` → the Notify name is compared (noted).
- BLFORM only: when MF MARKS/DESC is part of the B/L text (in order), it is a match; the ENTER column keeps
  the full text and the note lists the extra. Amendment forms stay strict.
- `*AS PER ATTACHED SHEET*` in MF: compare only the **first 30 letters/digits** before it with the start of
  ENTER's text — equal = match, no warning (user rule 2026-11-25; before that the whole prefix was compared).
- MF DESC starting `LCL AGENT'S LABOUR …` (MF STATUS empty) = STATUS **LCL**; the phrase is removed from DESC
  (user rule 2026-11-25, HASLC00000000000).
- MF package code with a full name in brackets wins: `240 PL (PAIL(s))` = PAIL (not PALLET); PAIL / TRAY / TANK
  are known package words, plural ENTER `PAILS` / `TRAYS` match (user rule 2026-11-25, HASLC00000000000).
- Customer's **red box** (red Square annotation) around MARKS / DESC on an ARRIVAL NOTICE: only the boxed text
  is checked; it matches when MF contains it (note `ตรวจเฉพาะในกรอบแดงของลูกค้า: …`) (user rule 2026-11-25,
  HASLC00000000000). Only the box that holds the **first** marks/desc line counts — other red boxes in the
  cargo area are customer notes (`PAY DO BKK / PICK UP LCB`, job-02). Red boxes are common on these
  ARRIVAL NOTICEs (Heung-A staff marks); regression 2026-11-25 changed only notes, no results, in job-01–05.
- **Split B/Ls are listed together**: rows are sorted by B/L text, so `…0900664` is followed by `…0900664A`–`L`
  (this skill's build_edi copy; before, sub-B/Ls were grouped at the end of the report) (user rule 2026-11-25).
- ARRIVAL NOTICE G.W. = the number right before `KGS`; `720 PKGS` in the count column is not a weight
  (HASLC00000000000).
- ENTER B/L typo: an ENTER B/L not in MF one character away from exactly one MF B/L without ENTER is
  compared with it; summary cell red `⚠ เลข B/L ใน ENTER ผิด (…)`.
- Thai font repair: garbled STATUS words are matched to known phrases (`_TH_LITERAL`).
- `REEL` = ROLL; MF `TEMP : - 18'C` / `MINUS 18 DEGREES` = -18°C; `PLTS` = PALLET; WOODEN CASE ≠ CASE.
- OCR (SCAN pages and letters whose text is only outlines): 300 dpi, slashed zeros fixed inside numbers,
  B/L matched to an MF B/L without ENTER within 2 characters (never when ambiguous). OCR rows are always
  orange `⚠ อ่านด้วย OCR — ตรวจด้วยตา`; an unreadable value stays empty. Handwriting (e.g. "BASKETS",
  Thai notes on scans) is not read — check it by eye.
- OCR of a scanned ARRIVAL NOTICE goes through the BLFORM reader when the G.W. column is found (else the old
  OCR reader), and fields it leaves empty are filled from the old reader (e.g. printed CY/CY). OCR repairs
  (`_ocr_blform_lines`): slashed zero / `e` / `B`→8 inside numbers, containers (`JPHU25eø588`), pen circle or
  highlighter on the count (`1±350 CTNS`, `1:800 BAGS`, `ld PALLETS`=10), highlighter hiding `CBM`
  (`45,351.000 KGS : 72.0000`), size×count before SAID TO CONTAIN dropped, pieces of one name joined
  (`ABC FOODS CO.` + `LTD.`), the customer's tick after a name dropped (`ABC CO.,LTD. 7`), handwritten notes
  below a gap > 70pt not taken as marks. Garbled form labels (`Cons i onee`, `I gn•e`, `Corel!re`,
  `btlfy Party`) are recognised; `TO THE ORDER` is never a label.
- `TO ORDER` / `TO THE ORDER (OF …)` consignee → Notify name compared.
- Report cells that start with `=` (customer typed `STATUS ===> ลากตู้`) are forced to text before saving —
  otherwise openpyxl writes a formula and Excel opens the file with "Repairs … Removed Records: Formula"
  and blanks the cell (EDI-SWMM.xlsx, 2026-11-24). Leading `=>-` is also dropped from STATUS text.
- Form labels (Shipper / Consignee / Notify Party) are never treated as annotation text; `。` counts as a space.
- MF with no `1) SHED NO. / 2) country / 3) CARGO MOVEMENT` lines at all: those three columns are orange
  `ไม่มีใน MF ฉบับนี้ (ต้องเป็น …)` — check them in the system — instead of red mismatches.
- `ตู้ลาก` = ลากตู้ (= CY). `TEMP: 6 DEGREE C` = 6°C. DSV internal `SHIPMENT NO: … COST CENTER : …` is not
  part of the description.
- Note column starts with `[ENTER หน้า n]` for every B/L (the PDF page(s) the values came from).
- The bundle has no single grand total; the TOTAL row compares the reported B/Ls only.

## When a bundle has a form this does not read

Dump each page's lines with x/y and its annotations (`page.get_text('dict')`, `page.annots()`), see which
type `page_kind` gives it, ask the user which values are which (or use their screenshot with arrows),
add a `FORM_PROFILES` entry (or a reader), add the form to the **Form registry**, and run the regression.
