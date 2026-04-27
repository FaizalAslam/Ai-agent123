# AI Agent Desktop Automation

A Flask-based desktop automation app for app launching, Office file automation, OCR, PDF tools, voice input, clipboard-triggered commands, and keyboard-triggered commands.

## Features

- App launcher with remembered executable paths
- Excel, Word, and PowerPoint file creation/editing
- OCR through EasyOCR
- PDF reading, creation, merge/split, and editing tools
- Voice, keyboard, and clipboard listeners

## Setup

1. Create and activate a virtual environment:

```bash
python -m venv .venv
.venv\Scripts\activate
```

2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Optional: create `.env` for OpenAI fallback parsing:

```bash
OPENAI_API_KEY=your_key_here
```

4. Run the Flask backend:

```bash
python server.py
```

The legacy Flask UI opens at `http://127.0.0.1:5000`.

## New React Frontend

The professional light-themed UI lives in `frontend/` and talks to Flask through a Next.js rewrite proxy.

1. Keep the Flask backend running:

```bash
python server.py
```

2. In a second terminal, start the Next.js frontend:

```bash
cd frontend
npm install
npm run dev
```

3. Open the Next.js URL shown in the terminal, normally `http://localhost:3000`.

The backend must remain running at `http://127.0.0.1:5000`. To override that, create `frontend/.env.local` from `frontend/.env.local.example`.

## Office Automation

Office document commands can be sent from the Office Agent tab or the command center. Examples:

- `create a new Excel file`
- `make a spreadsheet with 3 columns and 5 rows`
- `create a new Word document`
- `write a Word file about project status`
- `create a new PowerPoint presentation`
- `create a presentation with 3 slides about sales performance`

Generated Office files are saved under `outputs/` unless the request provides a specific file path or filename. Existing files opened with `open_workbook`, `open_document`, or `open_presentation` are loaded from the provided path and saved back to that path unless a save-as action supplies another output path.

## OCR Notes

This project uses `easyocr` from `requirements.txt`. Tesseract and `pytesseract` are not required by the current code.

## Configuration Files

- `app_paths.json`: app launcher paths and aliases, including Office executable paths
- `known_apps.json`: user-selected executable paths remembered by the launcher
- `command_map.json`: cached Office natural-language command mappings
- `.env`: local OpenAI key and other local-only settings

## Smoke Test

Run the regression smoke script after installing dependencies:

```bash
python smoke_test_office_routes.py
```

The script verifies Office routing, generated file existence, structured backend responses, unknown app-launch fallback behavior, and the frontend reliability script hook.

## Frontend Manual Checklist

- Assistant: `create a new Excel file`
- Assistant: `create a new Word document`
- Assistant: `create a new PowerPoint presentation`
- Assistant: `open chrome`
- Documents: screenshot OCR and image file OCR
- Documents: PDF Reader open, pause, resume, and stop
- PDF Tools: merge PDFs and split PDF
- PDF Tools: open, render, edit, and save PDF in the editor
- Outputs & History: confirm UI actions appear
- Backend offline: stop Flask and confirm the frontend shows a backend unavailable error
