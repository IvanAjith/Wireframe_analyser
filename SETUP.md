# SETUP — New Machine Checklist

Follow this once per machine. Takes about 5 minutes.

## 1. Clone the repo

```bash
git clone https://github.com/IvanAjith/Wireframe_analyser.git
cd Wireframe_analyser
```

## 2. Install Python dependencies

```bash
pip install -r requirements.txt
playwright install chromium
```

## 3. Install Tesseract (only needed for image wireframes, not HTML)

- **macOS:** `brew install tesseract`
- **Linux:** `sudo apt install tesseract-ocr`
- **Windows:** install from https://github.com/UB-Mannheim/tesseract/wiki
  and make sure it's added to your PATH.

## 4. Confirm it works

```bash
python3 run.py --help
```

If that prints usage instructions without errors, setup is done.

## Everyday use (after setup)

1. Drop your wireframe file (`.html`, `.png`, `.jpg`) into `source/`.
2. Run:
   ```bash
   python3 run.py your_wireframe.html
   ```
3. Open `output/your_wireframe_design_spec.docx`.

Re-running the same filename never overwrites — it writes `_v2`, `_v3`,
etc. instead. Delete old versions from `output/` yourself when you want
to tidy up.

See `UPDATE.md` for how to pull in a fix/enhancement Claude gives you in
chat. See `README.md` for the full explanation of what the tool does and
every available flag.
