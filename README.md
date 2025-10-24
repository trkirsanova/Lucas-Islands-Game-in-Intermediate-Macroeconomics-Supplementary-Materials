# Supplementary Materials: Lucas Islands Experiment

This repository provides supplementary materials for the classroom **Lucas Islands Experiment**, including Excel templates for student input and MATLAB scripts for automated data processing.

---

## 1. Overview

These materials illustrate how to use **Dropbox** to distribute and collect Excel-based responses from students, and how to process the resulting files automatically using MATLAB.

---

## 2. Files Provided

### Teaching and Distribution
- **`LucasIslands_ExcelTemplate.xlsx`** — Student input template.  
  Only designated input cells are editable; computed cells (e.g., wages, prices) contain pre-inserted formulas.
- **`LucasIslands_InfoSheet.pdf`** — Instructions for students.
- **`ExperimentSlides2025.tex`** — slides used in the lecture.

### Data Processing
- **`process_files.m`** — MATLAB script that reads all student Excel files, assigns sequential Player IDs, and compiles key data into a single Excel file `combined.xlsx`.
- **`write_matrices.m`** — MATLAB script that appends aggregate and dummy-variable sheets to `combined.xlsx`.

### Example of outotput
- **`LucasIslands_ExcelTemplate (3) Oleg.xlsx`** — Excel file as submitted by a student.
- **`LucasIslands_ExcelTemplate (4) Tatiana.xlsx`** — Excel file as submitted by a student.
- **`LucasIslands_ExcelTemplate (5) Nigar.xlsx`** — Excel file as submitted by a student.
- **`combined.xlsx`** - excel file which will be produced by matlab code.

---

## 3. Distributing Materials via Dropbox

1. Upload `LucasIslands_ExcelTemplate.xlsx` and create and upload `LucasIslands_InfoSheet.pdf` to a Dropbox folder (keep separate from the submissions folder).  
2. For each file, click **Share → Copy link** and ensure “Anyone with the link can view.”  
3. Post both links on Moodle or email them to students.  
   - Students should **download** the Excel template (not edit online).  
   - Fill only the unlocked cells.  
   - Save before uploading.

---

## 4. Collecting Submissions via Dropbox File Request

1. In Dropbox, choose **Create (+) → Send file request**.  
2. Set a title (e.g., *Lucas Islands Excel submission*) and select a submissions folder.  
3. Share the request link on Moodle or by email. Students can drag and drop files without seeing others’ submissions.  
4. **Anonymity:**  
   - In the *Name* field, students can enter their **island letter** (A–J) or simply `anon`.  
   - For the *Email* field, use a placeholder like `anon@example.com` if anonymity is required.  
5. Dropbox automatically appends counters to duplicate filenames, so no submissions are lost.

---

## 5. Practical Tips

- Keep protection on the Excel template; only unlock cells that students must edit.  
- Test the full workflow once using a private or incognito browser:  
  download → enter dummy data → upload via file request → run scripts.

---

## 6. Processing Submissions with MATLAB

### Step 1: Run Scripts
1. Download or sync the Dropbox submissions folder locally.  
2. Run the scripts in order:
   1. `process_files.m`
   2. `write_matrices.m`

### Step 2: Output from `process_files.m`
Creates `combined.xlsx` with the following sheets:
- **`Pe`** — Price forecasts by Player ID (values from `C5:C19` in each submission).  
- **`W`** — Individual wages by Player ID (values from `B5:B19`).  
- **`Is`** — Mapping of ID, island letter (A–J), and numeric code (A=1, …, J=10).  
- **`Mapping`** — Links each Player ID to the original filename.

### Step 3: Output from `write_matrices.m`
Appends additional sheets to `combined.xlsx`:
- **`W_table`** — Wage table by period (*t*) and island number.  
- **`P`** — Aggregate price by period (*t*).  
- **`Nb`, `Nr`** — Island-level demand dummies (booms/reductions).  
- **`Ei`, `Eu`, `Mc`, `Me`** — Time-series dummies:
  - `Ei`: inflation up  
  - `Eu`: unemployment up  
  - `Mc`: contractionary policy  
  - `Me`: expansionary policy

### Step 4: Analysis
Import `combined.xlsx` into **Stata**, **R**, or **Python** for analysis.  
Sequential Player IDs provide consistent keys across all sheets.

---

## 7. Citation

If you use these materials, please cite the corresponding paper:

> Hashimzade, N. et al. (2025). *Lucas Islands Game in Intermediate Macroeconomics*. Journal of Economic Education.

---

## 8. License

Unless otherwise stated, all files are released for educational and research use under a Creative Commons Attribution–NonCommercial license.
