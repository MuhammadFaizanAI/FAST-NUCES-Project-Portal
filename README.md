# FAST-NUCES CFD — Project Group & Partner Portal
**Managed By Aoun Malik • FAST NUCES CFD**

## Folder Structure (Single Netlify + GitHub Deployment)
```text
FAST-NUCES-Project-Portal/
├── config.js          # Single configuration file (Google Apps Script URL & Branding)
├── index.html         # Student Portal  -> https://appname.netlify.app/
├── admin/
│   └── index.html     # CR Admin Panel  -> https://appname.netlify.app/admin
├── netlify.toml       # Clean routing & instant cache refresh on GitHub commits
├── Code.gs            # Google Apps Script backend (auto-creates sheets & backups)
└── README.md          # Setup & deployment guide
```

## Step 1: Create Google Spreadsheet & Apps Script Backend
1. Go to [sheets.google.com](https://sheets.google.com) and create a new blank spreadsheet (e.g., `FAST_NUCES_Project_Portal_DB`).
2. Click **Extensions → Apps Script** in the top menu.
3. Delete any default code in `Code.gs`, paste the entire contents of `Code.gs` from this folder, and press **Ctrl + S**.
4. Select `setupSheets` in the top toolbar and click **Run** (accept permissions once). This automatically creates and formats the 5 tabs: `Courses`, `Students`, `Groups`, `Requests`, and `Backups`.
5. Optional: Select `installDailyBackupTrigger` and click **Run** to enable automatic nightly backups at 1:00 AM.
6. Click **Deploy → New deployment → Select type: Web app**:
   - **Execute as:** `Me`
   - **Who has access:** `Anyone`
   - Click **Deploy** and copy the **Web app URL** (`https://script.google.com/macros/s/.../exec`).

## Step 2: Connect URL & Push to GitHub + Netlify
1. Open `config.js` and paste your Web App URL into `SCRIPT_URL`.
2. Push this entire folder to a GitHub repository.
3. Log in to [Netlify](https://app.netlify.com), click **Add new site → Import an existing project → GitHub**, select your repo, and click **Deploy site**.
   - **Student Portal:** `https://appname.netlify.app`
   - **Admin Panel:** `https://appname.netlify.app/admin` (Default PIN: `1234`)
4. Whenever you edit any file on GitHub and commit, Netlify automatically redeploys in seconds.
