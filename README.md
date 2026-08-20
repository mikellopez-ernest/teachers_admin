# Teacher Admin Apps Script Workspace

This repository contains Google Apps Script web apps used around teacher administration.

## Structure

Each Apps Script project lives in its own folder. A script folder should contain:

- Apps Script source files, such as `Code.js`, `Index.html`, and `appsscript.json`.
- A local `README.md` explaining that script.
- A `docs/` folder with specification files and implementation notes.

Current scripts:

- `teacher_admin/`: teacher database administration endpoint.

## Local Secrets

Do not commit local Apps Script credentials or deployment configuration.

Files such as `.clasp.json`, `.clasprc.json`, `.env*`, credential JSON files, and local secret files are ignored by the root `.gitignore`.

If a script needs to be deployed with clasp, keep its `.clasp.json` inside that script folder locally, but never publish it.
