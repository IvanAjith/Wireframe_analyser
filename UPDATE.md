# UPDATE — Pushing a Fix From Claude Into GitHub

Claude cannot push to your GitHub itself — no credentials are ever
stored or shared in chat. Every time a change is made to this tool, you
get updated file(s) as a download, plus the block below. You download,
paste, run. That's the entire process, every time.

---

## The block (100% copy-paste, nothing to edit)

1. **Download** every file Claude gives you in that message (click each
   attachment — they land in `~/Downloads` by default).

2. **Paste this into Terminal and run it.** List one `mv` + `git add`
   pair per file Claude gave you — copy the two lines below and repeat
   them for each filename:
```bash
cd ~/Downloads/Wireframe_analyser
mv ~/Downloads/FILENAME .
git add FILENAME
git commit -m "Update from Claude"
git push
```

**Real example** — if Claude gave you `wireframe_analyzer.py` and
`run.py` in one message:
```bash
cd ~/Downloads/Wireframe_analyser
mv ~/Downloads/wireframe_analyzer.py .
git add wireframe_analyzer.py
mv ~/Downloads/run.py .
git add run.py
git commit -m "Update from Claude"
git push
```

Note: if a file listed above doesn't actually exist in `~/Downloads`
(e.g. it saved somewhere else, or under a slightly different name), the
`mv` line for it will print a harmless "No such file or directory" and
do nothing — `git add`/`commit`/`push` still run fine for whichever files
*did* move successfully. Check `git status` afterward if unsure what
actually got committed.

---

## Getting that update onto a different machine

```bash
cd Wireframe_analyser
git pull
```

## If `git push` is rejected

Means the remote has something your local copy doesn't yet:
```bash
git pull --rebase
git push
```
