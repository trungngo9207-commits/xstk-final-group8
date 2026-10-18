# AGENTS.md — briefing for AI agents working on this repo

Read this first. It is the single source of truth for what has been done, what the numbers are, and how to work safely in the team's shared Excel file.

## 1. What this is

- **Course:** TOAE203 *Statistics and Probability in Economics and Business* (FTU).
- **Assessment:** Final project *Wine Prices and Regional Price Premia*. Full brief: `course/project/Final_Project_Details.pdf` (text: `course/text/Final_Project_Details.txt`).
- **Team:** Group 8. The repo owner (Trung) is responsible for **parts B1 and B2**. Other parts belong to teammates. Do not edit their parts unless the owner asks.
- **Deadline:** both files (Excel workbook + Word/PDF report) are due **18 Oct 2026**, with the team's final review meeting the same day. After that there is an **individual Q&A (25 points)** with the lecturer.

### Project rules that affect every answer
- Every statistic, test and regression must be computed **in the Excel workbook**, with formulas or the Data Analysis ToolPak. Formulas must be visible. Raw output without interpretation earns no credit.
- Every hypothesis test is reported in 4 steps: (1) H₀ and H₁; (2) test statistic plus the Excel formula/output it comes from; (3) p-value (or critical value) at α = 0.05; (4) decision and its meaning for wine prices.
- Report ≤ **8 pages**, answers in question order. For each: key number/table from Excel, then 2–4 sentences of interpretation.
- File names: `Group_8.xlsx`, `Group_8.docx`. List members and contributions on page 1.

## 2. Repo map

| Path | Contents |
|---|---|
| `course/project/` | Final project brief (PDF) |
| `course/lectures/` | Lectures 0–12 (PDF). Lecture 8 = estimation/CI, 9 = hypothesis tests, 10 = two samples (pooled, Welch), 11 = ANOVA ↔ regression, 12 = logs/χ² |
| `course/Course_Review.pdf` | Lecturer's summary: the "model ladder" B1→B6 is built on |
| `course/text/` | Plain-text versions of all PDFs (use these to read quickly) |
| `data/Group_8_Data.csv` | Group 8 dataset. **Decimal commas** (`6,7`), UTF-8 BOM, 806 rows, has an `ln(Price)` column |
| `data/Wine_Dataset.xlsx` | Same 806 wines, original columns A–O |
| `workbooks/Group_8_local_snapshot_2026-10-11.xlsx` | Local copy of the team workbook from 11 Oct (Parts A–C; B1/B2 at Part B rows 3–59 + K41:L46). **Outdated vs the online file**, see §4 |
| `deliverables/b1-b2/` | B1/B2 outputs: tutorial video (20 min, 14 chapters), full report section (.docx, 6 pages), short report section VI + EN (.docx, ~1.4 pages), report text (.md), **Q&A practice** (`QA_practice.md`, 50 questions with model answers) |
| `tools/video/` | Python that generates the tutorial video (spreadsheet mock-up + charts + Piper TTS) |
| `tools/report/` | Node scripts (docx-js) that build the report .docx files from `cells.json` |

## 3. Data facts

- 806 California wines, 2001 vintage. Columns: `ID, Price, Cases (thousands), Score, Age, Napa valley, bay area, sonoma, sierra/foothills, mendocino / lake, Nonvarietal, Pinot noir, Cabernet, Merlot, Shyrah`.
- Reference categories: all region dummies 0 = **Generic California**; all variety dummies 0 = **Zinfandel**.
- Rows are sorted by Price (ascending). Price is heavily right-skewed: mean 29.58, median 22, skewness 7.87, max 440.
- In the **shared online workbook**, sheet `Data` is an Excel Table in A:S (A ID, **B Price**, …, **F Napa valley**, …, P Ln(Price), Q Region dummy, R Variety dummy, S Corresponding region). Columns **U onward hold a teammate's regression output** — do not put helper columns there.

## 4. B1/B2 — status and verified results

Status: **done** in the shared online workbook, sheet **Part B**. All results were verified independently in Python (scipy/statsmodels) and match to full precision.

### Current cell map in the shared online workbook (as of 18 Oct 2026)
A teammate rearranged Part B into a vertical layout on 17–18 Oct. **Re-check before citing cells; it may move again.**

| Block | Cells | Key cells |
|---|---|---|
| B1 · Price = β₀ + ε | A1:B14 | B3 n, B4 mean, B5 s, B6 SE, B7 t*, B8:B9 95% CI, B10 μ₀ = 25, B11 t, B12 p, B13 decision, B14 CONFIDENCE.T |
| B2.1 · Welch t-test | A16:C28 | B18/C18 n (Napa/Other), B19/C19 means, B20/C20 variances, row 21 hypotheses, B22 difference, B23 SE, B24 t, B25 df, B26 p (T.DIST.2T), B27 p (T.TEST type 3), B28 decision |
| B2.2 · Regression on Napa dummy | A30:E43 | LINEST 5×2 spill at B32:C36; coefficient table A38:E40; checks E41:E42 (= 0); fitted-equation text A43 |
| B2.3 · Regression t = pooled t | A45:C53 | C47 pooled var, C48 SE, C49 t, C50 df, C51 p (T.TEST type 2), C52:C53 checks (= 0) |
| B2.4 · rank-based test (note) | A55:A58 | names Wilcoxon rank-sum / Mann–Whitney U |
| B2.4 · computed extra | F45:G58 | G47 mean, G48 median, G49 skewness; G51 rank sum R1, G52 U, G53 E[U], G54 ties Σ(t³−t), G55 SD (tie-corrected), G56 z, G57 p, G58 decision |
| B3 onwards (teammates) | from row 60 | — |

Formulas use no helper columns: `COUNTIF`/`AVERAGEIF` for group n/mean, `SUMPRODUCT((Data!$F$2:$F$807=1)*(Data!$B$2:$B$807-mean)^2)/(n-1)` for group variance, `T.TEST(IF(Data!$F$2:$F$807=1,Data!$B$2:$B$807),IF(Data!$F$2:$F$807=0,Data!$B$2:$B$807),2,3)` for Welch (verified: IF arrays work, FALSE is ignored), `=LINEST(Data!$B$2:$B$807,Data!$F$2:$F$807,TRUE,TRUE)` for the regression.

### Verified numbers

| Quantity | Value |
|---|---|
| n, mean, s, SE | 806 · 29.5806 · 34.1543 · 1.2030 |
| t* (0.05, 805) · 95% CI | 1.9629 · [27.2192 ; 31.9421] |
| One-sample t (μ₀ = 25) · p | 3.8076 · 0.000151 → reject H₀ |
| Napa n / mean / variance | 287 · 43.0592 · 2689.986 |
| Other n / mean / variance | 519 · 22.1272 · 171.312 |
| Welch: diff · SE · t · df · p | 20.9321 · 3.1149 · 6.7199 · 306.29 · 8.886E-11 (T.TEST: 8.874E-11) |
| LINEST: b₁ · b₀ · SE(b₁) · SE(b₀) | 20.9321 · 22.1272 · 2.4031 · 1.4340 |
| LINEST: R² · s(e) · F · df · SSR · SSE | 0.0862 · 32.6689 · 75.87 · 804 · 80972.66 · 858075.60 |
| Slope t · p | 8.7103 · 1.708E-17 (intercept t 15.43, p 3.04E-47) |
| Pooled: s_p² · SE · t · df · p | 1067.258 · 2.4031 · 8.7103 · 804 · 1.708E-17 |
| Mann–Whitney: R1 · U · E[U] · ties · SD · z · p | 154829.5 · 113501.5 · 74476.5 · 869232 · 3162.35 · 12.3405 · 5.48E-35 |

Key interpretations (for the report and the Q&A):
- B1: in Price = β₀ + ε, OLS minimises Σ(Price − β₀)², which is smallest at β₀ = x̄, so b₀ = x̄ = $29.58.
- B2.1: Napa wines cost $20.93 more on average. This is a **raw** gap: score, age, volume and variety are not held constant (that is B5/Part C).
- B2.2: slope = mean(Napa) − mean(Other), intercept = mean(Other) (check cells = 0).
- B2.3: the regression reproduces the **pooled (equal-variance) t-test**, because OLS assumes one common error variance: SSE/(n−2) = 32.67² = 1067.26 = s_p². Welch's SE (3.115) is more trustworthy here because the Napa variance is about 16 times larger. Both reject H₀.
- B2.4: Price is right-skewed, so the safer test is Wilcoxon rank-sum / Mann–Whitney U. The project only requires naming it. The computed extra also rejects H₀.

## 5. Report documents (`deliverables/b1-b2/`)

- `Group_8_B1_B2_report.docx`: detailed (6 pages), with an appendix cell map.
- `Group_8_B1_B2_short_VI.docx` / `_EN.docx`: ~1.4 pages, fits the 8-page team report. The VI version uses decimal commas.
- **Caveat:** the cell references in these documents and in `B1_B2_report_text.md` point to the **local snapshot layout** (Part B rows 3–59, LINEST at K42:L46), not the current online layout. If the team report cites cells, update them using §4. The numbers themselves are unchanged.
- Not yet mentioned in the documents: the computed Mann–Whitney extra (F45:G58). One sentence suffices: "As an extra check, the Mann–Whitney U test (Part B, F45:G58) gives z = 12.34, p ≈ 5.5×10⁻³⁵, also rejecting H₀."

## 6. Working in the shared Excel Online file (lessons learned)

Shared workbook (Excel Online on FTU OneDrive; **anyone with this link can edit**, so keep this repo private):
https://ftueduvn-my.sharepoint.com/:x:/g/personal/k64_2515520031_ftu_edu_vn/IQCUlJ_kZrxySYM_XGMa_ZUqARs2iDOju6qd0R6kGCZyz_U?rtime=85ovv2ks30g

- Several teammates edit at the same time. **Never insert or delete rows or columns in Part B, and never move other people's blocks, without the owner's go-ahead.**
- If the browser isn't signed in, edits are saved as an anonymous guest. Signing in with the owner's FTU account attributes them correctly.
- The grid is drawn on a canvas, so page-text extraction returns nothing. Read it with screenshots and zoom.
- **Do not click cells by screen coordinates while the view may be scrolling.** Once, typed values landed in a teammate's block (they were reverted). Navigate with the keyboard: `Ctrl+Home` → arrow keys to the target cell; `Enter` moves down.
- Clicking the Name Box does not reliably focus it. Typed text can end up in a cell, or Excel tries to create a named range.
- If the first character typed into a cell is non-ASCII (e.g. `μ`), it can be dropped. Start labels with an ASCII character.
- Before typing a font size, confirm the size box text is highlighted; otherwise the typed text goes into the active cell.
- After `Ctrl+B`, follow-up `Shift+Arrow` keystrokes were sometimes lost. Check the Name Box shows the expected range (e.g. `13R x 2C`) before formatting.
- The fill-color Eyedropper did not work. Use **More Colors → Hex `4472C4`**.
- Part B style: titles Times New Roman 11 **bold**; tables Times New Roman 9.5; header rows fill **#4472C4** with white bold text; thin all-borders on tables.
- The `Lưu ý` (notes) tab holds the team's task table and instructions (in Vietnamese). It asks owners of B2/B3/B5 to use the ToolPak in desktop Excel. B1/B2 were done with formulas instead, which give identical results and are allowed by the brief.

## 7. Open items

- [x] Tick "B 1,2" as done in the `Lưu ý` task table (done 18 Oct 2026, cell D28 = TRUE).
- [ ] If the team report cites cells, update the B1/B2 references to the current online layout (§4).
- [ ] Optionally add the Mann–Whitney sentence (§5) to the report.

## 8. Rebuilding outputs

See `tools/README.md`. Summary:
- Video: `pip install -r tools/video/requirements.txt`, download the Piper voice `en_US-lessac-medium` (.onnx + .onnx.json) into `tools/video/voices/`, then `python tools/video/charts.py` and `python tools/video/script.py` (output `tools/video/out/B1_B2_guide.mp4`; `--preview 135 503` renders single frames).
- Reports: `cd tools/report && npm install && python export_cells.py && node build_report.js out.docx && node build_short.js vi short_vi.docx`.
