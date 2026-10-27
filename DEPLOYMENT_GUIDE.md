# XLSB → XLSX Converter: Deploying on Streamlit Community Cloud (from GitHub)

This guide takes you from "a folder on my PC" to "a live app at `https://<name>.streamlit.app`", including the small details that usually cause failed builds.

---

## 0. Read this first: what changed from your original code

| Topic | Original code | This version | Why |
|---|---|---|---|
| Conversion engine | Microsoft Excel through `pywin32` (COM) | **LibreOffice** (primary on the cloud), Excel (if present, e.g. your Windows PC), `pyxlsb` (last-resort fallback) | Streamlit Cloud is **Linux**. It has no Excel and no `pywin32`, so the original code would stop at "requires Windows" on the server. |
| `requirements.txt` | `pywin32` always | `pywin32` only on Windows (`; sys_platform == "win32"`) | A plain `pywin32` line **fails the build** on Linux. |
| `packages.txt` | none | `libreoffice-calc` | This is how the cloud server gets LibreOffice. |
| Theme | default | Dark theme in `.streamlit/config.toml` plus custom styling | |
| Modes | radio buttons, ZIP results | Three tabs: Single / Multiple / Folder, as you specified | |
| Code layout | one file | `app.py` (UI) + `converter.py` (engines) | Engines can be tested without the UI. |

### How the three modes behave

1. **Single file**: upload one `.xlsb`. The result appears in the page with a **Download XLSX** button and a data preview.
2. **Multiple files**: upload several `.xlsb` files. A results table appears, with a **Download** button per file and **Download all as ZIP**.
3. **Folder**: only `.xlsb` files are converted and everything else is ignored. The output goes into a new **`Outputs`** folder inside that folder. There is one caveat on the cloud:

> **A hosted web app cannot write to folders on your computer.** The server only sees its own disk, and no browser lets a website write into an arbitrary folder.
> - **Run on your own PC:** type the folder path. The app creates `<your folder>\Outputs\` and saves the files there. This is exactly your spec.
> - **Run on Streamlit Cloud:** use **Upload a folder**. The app converts the files and gives you **`Outputs.zip`**, which contains an `Outputs` folder. Unzip it next to your originals. The app picks this option by default when it detects it is hosted.

### Fidelity of each engine

| Engine | Where it runs | Keeps formulas / formatting / charts? |
|---|---|---|
| Microsoft Excel | Your Windows PC with Excel only | Best. It is Excel itself. |
| **LibreOffice** | Streamlit Cloud (and any PC with LibreOffice) | Good, but not identical to Excel. Complex charts, pivot tables and exotic formatting can differ. |
| pyxlsb (fallback) | Anywhere | **Values only.** No formulas, styles, or date formatting (dates appear as serial numbers). |

The default **Auto** setting uses the best available engine and falls back to the next one if a file fails. The sidebar shows which engines are active (green dot).

---

## 1. Prerequisites

- A **GitHub account**. Create one at <https://github.com/signup> if needed.
- A **Streamlit Community Cloud account**. It is free. You sign in with GitHub in step 5.
- (Optional) **Git** installed to push from the command line: <https://git-scm.com/downloads>. You can skip it and use GitHub's web upload instead (step 3, Option A).
- This project folder, unzipped.

---

## 2. Check the project layout

The **repository root** must look exactly like this. `app.py` must not be inside an extra subfolder.

```
your-repo/
├── app.py                  ← UI (entry point)
├── converter.py            ← conversion engines
├── requirements.txt        ← Python packages
├── packages.txt            ← system (apt) packages → installs LibreOffice
├── .gitignore
├── DEPLOYMENT_GUIDE.md     ← this file (optional in the repo)
└── .streamlit/
    └── config.toml         ← dark theme + server settings
```

Details that matter:

- **`packages.txt` must be in the repo root.** Streamlit Cloud only detects it there. Python requirements files can be in the root or next to `app.py`, so keep everything in the root.
- The `.streamlit` folder name starts with a dot and its name is exact. It is a hidden folder on some systems, so make sure it was included when you copied the project.
- Use **one** dependency file only (`requirements.txt`). Do not also add `Pipfile`, `pyproject.toml`, `environment.yml` or `uv.lock`. Streamlit Cloud picks one by priority and ignores the rest, which can silently skip your packages.
- Do **not** commit any `.xlsb` / `.xlsx` data files. The `.gitignore` already blocks them. They would be public if the repo is public.

---

## 3. Put the code on GitHub

### 3.1 Create the repository

1. Sign in at <https://github.com>.
2. Click **+** (top right) → **New repository**.
3. **Repository name:** e.g. `xlsb-to-xlsx-converter`. Use letters, numbers and hyphens.
4. **Visibility:**
   - **Public**: simplest. The app is public by default.
   - **Private**: possible, but a free Streamlit account allows only **one private app at a time**, and you must grant Streamlit access to your private repos (step 5.3).
5. **Leave unticked** "Add a README", ".gitignore" and "license" if you plan to push an existing folder (Option B). They create a commit that conflicts with your first push. For Option A you may tick them.
6. Click **Create repository**.

### 3.2 Upload the files: pick one option

#### Option A: GitHub website (no Git needed)

1. On the new empty repo page click **uploading an existing file** (or **Add file → Upload files**).
2. Open the project folder in File Explorer/Finder. Select **everything inside it** (`app.py`, `converter.py`, `requirements.txt`, `packages.txt`, `.gitignore`, `DEPLOYMENT_GUIDE.md`) and drag them into the browser.
3. The hidden `.streamlit` folder is often missed. Check that `.streamlit/config.toml` appears in the upload list.
   - If it does not: click **Add file → Create new file**, type `.streamlit/config.toml` in the name box (typing `/` creates the folder), paste the contents of `config.toml`, and commit.
4. Scroll down, keep **Commit directly to the main branch**, click **Commit changes**.
5. Verify on the repo's main page that `app.py` is visible **at the top level**, along with `packages.txt` and the `.streamlit` folder.

#### Option B: Git command line

Run these in the project folder (the folder that directly contains `app.py`):

```bash
git init
git add .
git commit -m "Initial commit: XLSB to XLSX converter"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```

- GitHub does **not** accept your account password for pushes. When prompted, sign in through the browser window (Git Credential Manager does this on Windows), or use a **Personal Access Token** (GitHub → Settings → Developer settings → Personal access tokens → *Tokens (classic)* → scope `repo`) as the password.
- If `git commit` says "Please tell me who you are", run:
  ```bash
  git config --global user.name "Your Name"
  git config --global user.email "you@example.com"
  ```
- Check `git ls-files`: it must list `.streamlit/config.toml`, `packages.txt`, and `requirements.txt`.

---

## 4. (Recommended) Test locally before deploying

On Windows with Python 3.11+ installed:

```bat
cd path\to\xlsb-to-xlsx-converter
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

It opens at `http://localhost:8501`. In the sidebar, **Microsoft Excel** should show a green dot if Excel is installed. Folder mode with a typed path writes to `<folder>\Outputs`.

To test the cloud engine locally, install LibreOffice (<https://www.libreoffice.org/download/>) and pick **LibreOffice** in the sidebar's engine list.

---

## 5. Deploy on Streamlit Community Cloud

### 5.1 Sign in

1. Go to <https://share.streamlit.io>.
2. Click **Continue to sign-in** and choose **Continue with GitHub**.
3. Authorize Streamlit to read your GitHub account when asked.

### 5.2 Create the app

1. Click **Create app** (upper-right).
2. When asked if you already have an app, click **Yup, I have an app**.
3. Fill in:
   | Field | Value |
   |---|---|
   | **Repository** | `<your-username>/<your-repo>` |
   | **Branch** | `main` |
   | **Main file path** | `app.py` |
   | **App URL** (optional) | choose a subdomain, e.g. `xlsb-converter` → `https://xlsb-converter.streamlit.app`. Leave blank for an auto-generated URL. |
   
   You can instead click **Paste GitHub URL** and paste the link to `app.py` on GitHub.
4. Click **Advanced settings**:
   - **Python version:** choose **3.12** (the documented default). Use the same version you tested with locally if possible.
   - **Secrets:** leave empty. This app needs none.
   - Click **Save**.
5. Click **Deploy**.

### 5.3 Private repository? Grant access

If your repo is private and it does not appear in the dropdown, click **Connect/Add GitHub access** or **Manage GitHub permissions** and allow Streamlit to access private repositories (or that specific repo). Then reload the page.

### 5.4 Wait for the first build

- The first build is **slower than normal (several minutes)** because the server downloads and installs LibreOffice from `packages.txt`. Do not refresh repeatedly.
- Open **Manage app** (bottom-right of the running app) to see the logs. Look for lines installing `libreoffice-calc` and then the Python packages.
- A progress/"oven" screen is normal during this time.

### 5.5 Verify it works

1. Open the app. In the **sidebar** under *Engines on this machine*, **LibreOffice must show a green dot** and the header chip should say **Engine: LibreOffice**.
2. Go to **Single file**, upload a small `.xlsb`, click **Convert to XLSX**, download the result and open it in Excel.
3. Try **Multiple files** with 2–3 files, then **Folder → Upload a folder**.
4. Compare the converted file against the original. Check formulas, number/date formats, column widths, charts and any named ranges. LibreOffice does not always match Excel exactly (see the fidelity table in section 0).

> **If the chip says "Engine: Python (pyxlsb)"**, LibreOffice did not install. The app still works, but it only copies values. Go to Troubleshooting → *"LibreOffice not found"*.

---

## 6. Updating the app later

- Edit files on GitHub (pencil icon) or `git push` new commits to `main`. The app updates automatically within seconds to a couple of minutes.
- Changing `requirements.txt` or `packages.txt` triggers a **full reinstall**, which takes longer.
- To force a restart (for example after a memory problem): open the app → **Manage app** → ⋮ menu → **Reboot app**.
- Change the URL, Python version, or sharing: go to <https://share.streamlit.io>, find your app, ⋮ → **Settings**.

---

## 7. Access, privacy and limits: please read

**Who can use the app.**
- Apps from **public** repos are **public**. Anyone with the URL can upload files.
- To restrict it: app ⋮ → **Settings → Sharing → "Only specific people can view this app"**, then add email addresses. Invited people sign in with Google or an emailed link.

**Confidential data.** Uploaded files are processed on Streamlit's shared servers (outside your company network). The app keeps them only in temporary storage for your session and does not store them permanently, but if your workbooks are confidential/financial, check your company's policy, and restrict viewers as above, or run the app internally on your own PC instead (section 4). Never put real data files in the GitHub repo.

**Resource limits** (set by Streamlit and subject to change; check the current docs):
- Roughly **0.7 GB to 2.7 GB of RAM** and up to 2 CPU cores per app. Large `.xlsb` files expand a lot in memory when opened by LibreOffice. A 50-100 MB file can exhaust memory. If you see *"This app has gone over its resource limits"*, use smaller files or fewer at a time, then **Reboot app**.
- **Upload size:** `.streamlit/config.toml` sets `maxUploadSize = 200` (MB per file). Raising it does not add RAM.
- Files are converted **one at a time**, so large batches are slow rather than failing. The per-file timeout (default 300 s) is adjustable in the sidebar.
- **Sleep:** apps with no traffic for **12 hours** go to sleep. The next visitor clicks **Yes, get this app back up!** and waits about a minute.
- Converted files live in temporary storage and disappear when the session ends or the app restarts. Download them right after converting.

---

## 8. Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| Build log: *"ERROR: Could not find a version that satisfies the requirement pywin32"* | `pywin32` listed without the Windows marker | Use the provided `requirements.txt` (`pywin32>=306; sys_platform == "win32"`). |
| Sidebar shows **LibreOffice: not found** / chip says pyxlsb | `packages.txt` is not in the repo **root**, has a typo, or the apt install failed | Confirm the file is at the top level, contains exactly `libreoffice-calc` on its own line, then check **Manage app** logs for apt errors. After fixing, commit and wait for the reinstall, or **Reboot app**. |
| apt install fails on `libreoffice-calc` | Package name/availability issue on the server image | Try `libreoffice-core` and `libreoffice-calc` as two lines, or check the logs for the exact apt error. If it cannot be resolved, delete `packages.txt`; the app then runs on the values-only pyxlsb fallback. |
| `ModuleNotFoundError: No module named 'converter'` | `converter.py` is missing or in a subfolder | It must sit **next to `app.py`**. |
| `ModuleNotFoundError: pyxlsb / openpyxl / xlsxwriter` | `requirements.txt` not picked up | Make sure it is in the root and named exactly `requirements.txt`; no second dependency file. |
| App stuck on the "oven" screen | First build is still installing LibreOffice | Wait 5-10 minutes. If it never finishes, open **Manage app** logs. |
| Theme is not dark | `.streamlit/config.toml` missing from the repo, or wrong folder name | Confirm the path is `.streamlit/config.toml` (with the leading dot). Reboot the app. |
| *"Folder picking needs a newer Streamlit"* | Old Streamlit version resolved | Keep `streamlit>=1.50.0` in `requirements.txt`, then reboot. |
| Conversion "did not finish within N seconds" | Large/complex workbook | Raise **Timeout per file** in the sidebar, or convert it alone. |
| *"LibreOffice could not convert this file…"* | Corrupt, password-protected or not a real XLSB | Open it in Excel first. Remove the password and re-save. |
| *"This app has gone over its resource limits"* | Too much memory | Convert fewer/smaller files; **Reboot app**. |
| Upload shows *"File must be 200.0MB or smaller"* | Over `maxUploadSize` | Split the batch; very large files should be converted locally. |
| Folder mode path says folder "does not exist" on the cloud | The path refers to the **server's** disk | Use **Upload a folder** on the cloud. Typed paths work only when running on your own PC. |
| `git push` rejected / "failed to push some refs" | You created the repo with a README/license | `git pull origin main --allow-unrelated-histories`, then `git push`. |
| Changes not showing | Browser cache or build still running | Hard refresh (Ctrl+F5), check **Manage app** logs, then **Reboot app**. |

---

## 9. Known limitations

- `.xlsx` cannot contain **VBA macros**. Macro content in an `.xlsb` is dropped during conversion; keep the original (or use `.xlsm`) if you need macros.
- Password-protected workbooks are not supported by the LibreOffice and pyxlsb engines.
- Folder mode converts only the `.xlsb` files directly inside the chosen folder when using a typed path (sub-folders are not scanned); a folder **upload** keeps its sub-folder structure inside `Outputs/`. Excel's temporary lock files (`~$name.xlsb`) are skipped.
- Existing files with the same name inside `Outputs` are overwritten in typed-path mode.
- The Excel (COM) engine cannot be interrupted, so the timeout setting does not apply to it.

---

## 10. Quick reference

```text
Repo root must contain : app.py, converter.py, requirements.txt, packages.txt, .streamlit/config.toml
Main file path         : app.py
Python version         : 3.12 (Advanced settings)
Secrets                : none
First build            : slow (installs LibreOffice), normal
Check after deploy     : sidebar → LibreOffice = green dot
```

### Official documentation used

- [App dependencies for Community Cloud](https://docs.streamlit.io/streamlit-community-cloud/deploy-your-app/app-dependencies) (`requirements.txt`, `packages.txt`, Python version)
- [Deploy your app on Community Cloud](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy)
- [Manage your app](https://docs.streamlit.io/deploy/streamlit-community-cloud/manage-your-app) (sleep, reboot, logs, resource limits)
- [Share your app](https://docs.streamlit.io/deploy/streamlit-community-cloud/share-your-app) (private apps, viewers)
