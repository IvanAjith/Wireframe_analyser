# SETUP — How to Use This Repo

Two blocks. Block A you run once, ever, per machine. Block B you run
every single time you want to analyze a wireframe.

---

## Block A — one-time setup (once per machine)

```bash
git clone https://github.com/IvanAjith/Wireframe_analyser.git
cd Wireframe_analyser
pip install -r requirements.txt
playwright install chromium
```

Image wireframes (`.png`/`.jpg`) also need Tesseract installed. HTML
wireframes don't need this step.
- **macOS:** `brew install tesseract`
- **Linux:** `sudo apt install tesseract-ocr`
- **Windows:** install from https://github.com/UB-Mannheim/tesseract/wiki
  and make sure it's on PATH.

Confirm it worked:
```bash
python3 run.py --help
```
If that prints usage instructions with no errors, setup is done — you
will never need to run Block A again on this machine.

---

## Block B — every time you want to analyze a wireframe

**Step 1 (Finder, not Terminal):** drag your HTML/image file into the
`source` folder inside `Wireframe_analyser`.

**Step 2 (Terminal):** copy this, paste it, then edit the filename in
**both** places to match your actual file, then run it:
```bash
cd ~/Downloads/Wireframe_analyser
python3 run.py YOUR_FILE_NAME.html
open output/YOUR_FILE_NAME_design_spec.docx
```

Example, for a file called `homepage.html`:
```bash
cd ~/Downloads/Wireframe_analyser
python3 run.py homepage.html
open output/homepage_design_spec.docx
```

That's the whole workflow. `cd ~/Downloads/Wireframe_analyser` assumes
that's where you cloned it — adjust the path once if yours is different,
it'll be the same every time after that.

Re-running the same filename never overwrites — it writes `_v2`, `_v3`,
etc. instead. Delete old versions from `output/` yourself when you want
to tidy up.

`open` is Mac-only (opens the doc straight into Word). On another OS,
just open the file from `output/` manually — it'll be there either way.

---

See `UPDATE.md` for how to push a fix/enhancement Claude gives you in
chat. See `README.md` for the full explanation of what the tool does and
every available flag.
