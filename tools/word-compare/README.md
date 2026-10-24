# Word vs DocLens side-by-side

Renders documents with Microsoft Word (ground truth) and with DocLens, then
writes `out/<name>-<page>.png` images: Word on the left, DocLens on the right.

Requirements: Windows with Microsoft Word, Node.js 18+.

```powershell
# 1. a scratch folder with the documents to test
mkdir D:\wc\docs; copy "C:\path\report.docx" D:\wc\docs\

# 2. Word's rendering -> PDF
powershell -File tools\word-compare\export-word-pdf.ps1 -WorkDir D:\wc

# 3. DocLens rendering + side-by-side images
cd D:\wc
npm i puppeteer pdfjs-dist@4.10.38
$env:WORK_DIR = "D:\wc"; $env:NODE_PATH = "D:\wc\node_modules"
node D:\Krishna\chrome-extension\tools\word-compare\compare.js report
```

`MAX_PAGES` (default 60) limits how many pages are compared.
