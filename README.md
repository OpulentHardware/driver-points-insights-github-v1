# Driver Points and Insights

GitHub Pages version of the OPULENT HARDWARE Driver Points and Insights tool.

## Excel file name and location

The workbook must be named:

```text
data/season-results.xlsx
```

The workbook should keep the same structure as the current multi-round Excel file: each competition round is stored as its own worksheet/tab.

## Repository structure

```text
index.html
assets/app.js
assets/style.css
data/season-results.xlsx
.nojekyll
```

## How to update results

1. Export or save the latest Excel workbook as `season-results.xlsx`.
2. Replace the file at `data/season-results.xlsx` in the repository.
3. Commit and push to GitHub.
4. GitHub Pages will serve the updated workbook. The app fetches it on page load.

## GitHub Pages setup

In GitHub:

1. Go to Settings → Pages.
2. Set Source to `Deploy from a branch`.
3. Choose your branch, usually `main`.
4. Choose `/root` if these files are at the repository root.
5. Save.

## Local testing note

Because this version uses `fetch("data/season-results.xlsx")`, opening `index.html` directly from the file system may not auto-load the Excel file in some browsers. Test locally with a small static server, for example:

```bash
python3 -m http.server 8080
```

Then open:

```text
http://localhost:8080
```

The manual override input still works if you drag or choose a workbook locally.
