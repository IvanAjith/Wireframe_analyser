# UPDATE — Applying a Fix/Enhancement From Claude

Claude cannot push to your GitHub directly — no credentials are stored or
shared between chat sessions, by design. Instead, every time a change is
made to this tool in chat, you'll get:

1. An updated file (e.g. `wireframe_analyzer.py`) as a download, **and**
2. A ready-to-run command block like the one below.

Your side of it is: download the file, then copy-paste the block into
your terminal. Nothing else to figure out.

## The standard update block

Downloaded files usually land in `~/Downloads`. Adjust the two paths
below if yours are different, then run it as-is:

```bash
cd ~/Downloads/Wireframe_analyser              # your local repo folder
mv ~/Downloads/wireframe_analyzer.py .          # overwrite with the new version
git add wireframe_analyzer.py
git commit -m "Update from Claude: <short description of the fix>"
git push
```

If `run.py` or another file was also updated, add an extra `mv`/`git add`
line for each one before the commit, e.g.:
```bash
mv ~/Downloads/run.py .
git add run.py
```

## Getting the update on a different machine

Once pushed, any other machine just needs to pull it:
```bash
cd Wireframe_analyser
git pull
```

## If `git push` is rejected

This usually means the remote has a commit your local copy doesn't (e.g.
you updated from a different machine and forgot to `git pull` here
first). Fix:
```bash
git pull --rebase
git push
```
