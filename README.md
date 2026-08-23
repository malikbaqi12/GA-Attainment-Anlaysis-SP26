# GA Attainment Automation

Automates entry, analysis and visualization of Graduate Attribute (GA) attainment
data for the FYDP GA Attainment master workbook. Instead of manually finding
the right cell among hundreds of columns for each student/course/CLO, you pick
a few dropdowns and the app finds the cell for you. It also recomputes GA
attainment per student directly from the CLO scores and gives you charts.

## Files in this project

| File | Purpose |
|---|---|
| `app.py` | Streamlit front end: upload the workbook, enter scores, view analysis/visualizations, download the updated file. This is what you deploy to GitHub/Streamlit Cloud. |
| `GA_Attainment_Analysis.ipynb` | **Google Colab** notebook with the same parsing logic: upload the workbook, do guided data entry, run GA attainment analysis and charts, download the updated file. Runs entirely in the browser, no local install needed. |
| `README.md` | This file. |

## How the workbook is read

Each batch sheet (e.g. `AI-SP-26`, `AI-FA-25`, ...) has one row per student and
one block of columns per course, split into CLOs. The header block always
follows the same layout relative to the `Student Name` cell, even though the
exact row/column it starts at differs sheet to sheet:

```
Student Name row      -> "Student Name"
+1                    -> Course name (spans its columns)
+2                    -> CLO label (CLO 1, CLO 2, ...)
+3                    -> Component (Theory / Lab), optional
+4                    -> GA label (GA1, GA2, ...) that CLO feeds into
+5                    -> first row of student data
```

Roll number sits one column left of "Student Name". The app auto-detects
this per sheet, so it works even though `AI-FA-23` starts its header a few
rows lower than the others.

## Running the notebook on Google Colab

1. Go to [colab.research.google.com](https://colab.research.google.com).
2. **File → Upload notebook** → select `GA_Attainment_Analysis.ipynb`.
3. Run the cells top to bottom (**Runtime → Run all**, or step through with
   Shift+Enter). The upload cell (section 1) will pop up a **Choose Files**
   button — pick your `.xlsx` workbook there.
4. For data entry, either:
   - run the **guided interactive entry** cell (section 4a) and answer the
     prompts one score at a time, or
   - edit the **entries list** in section 4b directly in the cell and run it.
5. Section 5 saves the workbook and automatically downloads it back to your
   computer as `GA_Attainment_updated.xlsx`.
6. Sections 6–11 build the analysis dataset and charts (GA attainment per
   batch, a batch × GA heatmap, course-level averages, a below-threshold
   student list). Section 12 downloads the underlying data as CSV.

You can also run this notebook locally with Jupyter if you prefer — it
detects it isn't in Colab and just skips the upload/download popups (set
`FILE_PATH` in section 1 to your file's path instead).

## Running the Streamlit app locally

```bash
pip install streamlit pandas openpyxl plotly
streamlit run app.py
```

Then open the local URL Streamlit prints (usually `http://localhost:8501`).

## Deploying the Streamlit app on Streamlit Community Cloud via GitHub

1. Create a new GitHub repository and push `app.py` (and this README, and the
   notebook if you want it there too).
2. Add a `requirements.txt` in the repo root with:
   ```
   streamlit
   pandas
   openpyxl
   plotly
   ```
3. Go to [share.streamlit.io](https://share.streamlit.io), sign in with
   GitHub, click **New app**, choose the repository/branch, and set the main
   file path to `app.py`.
4. Deploy. Every push to the branch you selected will redeploy the app.

## Using the app

1. **Upload** your `.xlsx` workbook on the main page.
2. **Data Entry tab** — pick Batch, Course, Component, CLO and Student from
   the dropdowns (all populated straight from that sheet's own headers, so
   you can't pick an invalid combination), type the Score, and click
   **Add to queue**. Queue up as many scores as you like, then click
   **Apply all pending entries to the workbook** to write them into the
   in-memory file. Anything that can't be matched (e.g. a roll number not
   found in that batch) is reported and kept in the queue instead of being
   silently dropped.
3. **Analysis & Visualization tab** — shows, computed fresh from the raw CLO
   scores (not the sheet's own formulas, some of which contain pre-existing
   `#REF!`/`#DIV/0!` errors):
   - average GA attainment per batch (bar chart)
   - a batch x GA attainment heatmap
   - course-level average scores for a chosen batch
   - a list of students below a chosen attainment threshold
   - the full tidy dataset, downloadable as CSV
4. **Download tab** — download the workbook with your entries applied.

## Troubleshooting

**Colab: `FileNotFoundError` when running section 3 (loading the workbook)**
This means `FILE_PATH` points to a file that doesn't exist in the Colab
session's storage. It happens when:
- section 1 (the upload cell) hasn't been run yet in this session — run it
  first, then run section 3, or just use **Runtime → Run all**; or
- the Colab runtime disconnected or restarted since you uploaded — Colab's
  uploaded files live only in that session's temporary storage, so a
  disconnect clears them. Re-run section 1 to upload the file again.

The notebook now catches this and other common issues (an invalid/corrupted
file, an unreadable sheet) and prints a plain-English message telling you
which of these it is, instead of a raw traceback. The Streamlit app (`app.py`)
does the same on upload.

**A sheet is missing from the results**
Both `app.py` and the notebook look for a cell that says exactly
"Student Name" (case-insensitive) within the first 10 rows of a sheet to
find its header block. If a sheet is skipped, you'll see a warning naming it
— check that sheet has that header cell.

## Notes

- The app works entirely in memory during your session; nothing is saved
  until you use the download button.
- Formulas already present in the sheet are preserved as formulas on save,
  but Streamlit/openpyxl don't recalculate them — open the downloaded file
  in Excel or LibreOffice once to refresh any formula-driven cells.
- Some GA-average formulas in the original workbook (columns after the CLO
  data) already contained `#REF!` and `#DIV/0!` errors before this tool
  touched them — mostly from blank future-semester columns and a couple of
  broken cross-references. The Analysis tab sidesteps this by recomputing
  attainment directly from the raw CLO scores instead of relying on those
  formulas.
