# Spark Docs — Document Converter

A simple tool that converts Spark Academy project documents (.docx) into clean, mobile-friendly web pages.

## What This Does

Drop any Word document from the Spark project onto the converter page. It instantly renders a clean, readable version you can scroll through on any phone — and lets you download a standalone HTML file you can add to this repo to give people a permanent link.

---

## First-Time Setup (5 minutes, one time only)

### 1. Fork or clone this repo to your GitHub account

### 2. Enable GitHub Pages
- Go to your repo on GitHub
- Click **Settings** → scroll to **Pages**
- Under "Source," select **Deploy from a branch**
- Choose **main** branch, **/ (root)** folder
- Click **Save**

GitHub will give you a URL like:
```
https://your-username.github.io/spark-docs/
```

That's your converter tool — share it with Michelle or anyone on the project.

---

## How to Convert a Document

1. Go to your GitHub Pages URL
2. Drop a `.docx` file onto the page (or tap to choose a file)
3. Read it right there — it's already mobile-friendly
4. Click **Download HTML** to save the converted file

---

## How to Publish a Document as a Permanent Link

1. Convert the document using the tool above
2. Download the HTML file
3. Create a `docs/` folder in this repo (if it doesn't exist)
4. Add the HTML file there — e.g. `docs/benchmark-brief.html`
5. Commit and push

The document is now live at:
```
https://your-username.github.io/spark-docs/docs/benchmark-brief.html
```

Share that link with Michelle, family, or anyone who needs to review it.

---

## Folder Structure

```
spark-docs/
├── index.html          ← The converter tool (this is what people visit)
├── README.md           ← This file
└── docs/               ← Put converted HTML documents here
    ├── benchmark-brief.html
    ├── research-summary.html
    └── ... (any other converted docs)
```

---

## Notes

- The converter runs entirely in the browser — no files are uploaded to any server
- Documents are never stored anywhere — the conversion is private
- The downloaded HTML files are self-contained — they include all styling and work without an internet connection (except for loading fonts from Google Fonts)
- Any `.docx` file works, not just Spark project files
