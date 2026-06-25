# GeometryHome WFMS — Complete Railway Deployment Guide

---

## What This Guide Covers

- Migrating your existing SQLite data to PostgreSQL (one-time, done locally)
- Deploying the app to Railway so it runs 24/7
- Verifying everything works
- Troubleshooting common issues

---

## Prerequisites — Install These First

### On Your Local Machine (Windows):

1. **Python 3.10+** — https://python.org/downloads
   - During install, tick **"Add Python to PATH"**
   - Verify: open Command Prompt and run `python --version`

2. **Git** — https://git-scm.com/download/win
   - Install with all defaults
   - Verify: `git --version`

3. **psycopg2** (Python PostgreSQL driver) — open Command Prompt and run:
   ```
   pip install psycopg2-binary
   ```

---

## PART 1 — Set Up Railway (5 minutes)

### Step 1.1 — Create a Railway Account
1. Go to https://railway.app
2. Click **"Start a New Project"**
3. Sign up with GitHub (recommended) or email

### Step 1.2 — Create a New Project
1. After login, click **"New Project"**
2. Select **"Empty Project"**
3. Railway creates a project dashboard — leave this tab open

### Step 1.3 — Add a PostgreSQL Database
1. Inside your Railway project, click **"+ New"**
2. Select **"Database"** → **"Add PostgreSQL"**
3. Railway spins up a managed PostgreSQL instance (takes ~30 seconds)
4. Once it appears, click on the **Postgres** card
5. Go to the **"Connect"** tab
6. Find **"DATABASE_URL"** — it looks like:
   ```
   postgresql://postgres:AbCdEfGh@containers-us-west-1.railway.app:6543/railway
   ```
7. **Copy this entire URL** — you need it in Part 2

---

## PART 2 — Migrate Your Existing Data (5 minutes)

This runs on your local Windows PC. It reads your old SQLite file and
copies all 88 employees, settings, prices, and reports into PostgreSQL.

### Step 2.1 — Open Command Prompt in the App Folder
1. Extract the `WorkforceMealSystem_v52.zip` somewhere, e.g. `C:\WFMS\`
2. Open Command Prompt
3. Navigate to the folder:
   ```
   cd C:\WFMS\workforce_app
   ```

### Step 2.2 — Place Your SQLite Database File
Copy your existing `workforce.db` file into `C:\WFMS\workforce_app\`
(the same folder as `app.py`).

### Step 2.3 — Run the Migration Script
Replace the DATABASE_URL below with the one you copied from Railway:

```
python migrate.py --sqlite workforce.db --pg "postgresql://postgres:AbCdEfGh@containers-us-west-1.railway.app:6543/railway"
```

**Expected output:**
```
===================================================
  GeometryHome WFMS — Database Migration
  Source : workforce.db
  Target : postgresql://postgres:AbCd...
===================================================

▶ Creating tables (if not exist)...
  ✅ Tables ready

  ✅ settings: 4/4 rows migrated
  ✅ locations: 5/5 rows migrated
  ✅ food_preferences: 5/5 rows migrated
  ✅ users: 1/1 rows migrated
  ✅ employees: 88/88 rows migrated
  ✅ attendance: 9/9 rows migrated
  ✅ suspensions: 1/1 rows migrated
  ✅ meal_prices: 20/20 rows migrated
  ✅ report_history: 9/9 rows migrated
  ...

===================================================
  Migration complete! 133 rows migrated.
===================================================
```

If you see any ⚠️  errors, see the Troubleshooting section at the bottom.

---

## PART 3 — Push Code to GitHub (3 minutes)

Railway deploys from GitHub. You need to put the app code into a repo.

### Step 3.1 — Create a GitHub Account (if you don't have one)
Go to https://github.com and sign up.

### Step 3.2 — Create a New Private Repository
1. Click **"+"** → **"New repository"**
2. Name it: `gh-meal-system` (or anything you like)
3. Set to **Private**
4. **Do NOT** tick "Add README" or any other options
5. Click **"Create repository"**
6. GitHub shows a page with instructions — leave it open

### Step 3.3 — Push Your Code
Open Command Prompt in `C:\WFMS\workforce_app\` and run these one by one:

```
git init
git add .
git commit -m "Initial deployment"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/gh-meal-system.git
git push -u origin main
```

Replace `YOUR_USERNAME` with your GitHub username.

When prompted, enter your GitHub username and password
(or personal access token if you have 2FA enabled).

---

## PART 4 — Deploy on Railway (5 minutes)

### Step 4.1 — Connect GitHub to Railway
1. Back in Railway, click **"+ New"** → **"GitHub Repo"**
2. Click **"Configure GitHub App"** if asked
3. Authorise Railway to access your repos
4. Select the `gh-meal-system` repo you just created
5. Railway automatically detects it's a Python app and starts building

### Step 4.2 — Set Environment Variables
1. Click on your newly created **service** (not the Postgres one)
2. Go to the **"Variables"** tab
3. Click **"+ New Variable"** and add these one by one:

| Variable Name | Value |
|---|---|
| `DATABASE_URL` | *(click "Add Reference" → select your Postgres plugin — Railway fills this automatically)* |
| `SECRET_KEY` | Any long random string, e.g. `gh-wfms-2024-xK9mP3nQ7rL2wA8vB5` |

**For DATABASE_URL specifically:**
- Click **"+ New Variable"**
- In the name field type: `DATABASE_URL`
- In the value field, click **"Add Reference"** button
- Select your PostgreSQL plugin from the dropdown
- Railway links them automatically — this is the correct way

### Step 4.3 — Trigger a Deployment
1. Go to the **"Deployments"** tab
2. Click **"Deploy Now"** (or it may have already started)
3. Click on the deployment to see live logs
4. Wait for it to show:
   ```
   ======================================================
     GeometryHome WFMS v5.2  |  DB: PostgreSQL ✅
     Listening on port 8080
   ======================================================
   ```

If you see `SQLite ⚠️` instead of `PostgreSQL ✅`, the DATABASE_URL
variable is not set correctly — go back to Step 4.2.

### Step 4.4 — Get Your Public URL
1. Go to the **"Settings"** tab of your service
2. Under **"Networking"** → **"Public Networking"**, click **"Generate Domain"**
3. Railway gives you a URL like `gh-meal.up.railway.app`
4. Open it in your browser — you should see the login page

---

## PART 5 — Verify Everything Works

### Step 5.1 — Check the Health Endpoint
Open this URL in your browser (replace with your domain):
```
https://gh-meal.up.railway.app/api/health
```

You should see:
```json
{"status": "ok", "database": "postgresql", "version": "5.2"}
```

If `"database"` says `"sqlite"` — the DATABASE_URL is not set. Recheck Part 4.2.

### Step 5.2 — Log In
1. Open your Railway URL
2. Log in with your existing username and password (migrated from old DB)
3. Check Dashboard — your 88 employees should all be there
4. Check Supplier Report, Billing prices — all data intact

### Step 5.3 — Test Data Persistence
1. Add a test employee or mark someone absent
2. Go to Railway → your service → **"Restart"**
3. After it restarts, check that the test data is still there
4. ✅ If yes — PostgreSQL is working correctly, data is permanent

---

## PART 6 — Future Deployments (after code changes)

Whenever you get an updated `index.html` or `app.py`:

1. Replace the file in `C:\WFMS\workforce_app\`
2. Open Command Prompt in that folder and run:
   ```
   git add .
   git commit -m "Update: description of what changed"
   git push
   ```
3. Railway detects the push and automatically redeploys
4. No data is lost — the PostgreSQL database is completely separate from the code

---

## Troubleshooting

### "psycopg2 not found" during migration
```
pip install psycopg2-binary
```

### Migration error: "could not connect to server"
- Double-check the DATABASE_URL — copy it fresh from Railway
- Make sure you're wrapping the URL in quotes in the command
- Check that your internet connection is working

### Migration error: "SSL connection required"
Add `?sslmode=require` to the end of your DATABASE_URL:
```
python migrate.py --sqlite workforce.db --pg "postgresql://...railway.app:5432/railway?sslmode=require"
```

### Railway build fails: "No module named fastapi"
- Check `requirements.txt` is in the `workforce_app` folder
- Make sure you committed it: `git add requirements.txt && git commit -m "add requirements" && git push`

### App shows white page after deployment
- Go to Railway → Deployments → click the latest one → read the logs
- Look for any Python error in red
- Most common cause: DATABASE_URL not set (app can't connect to DB on startup)

### Login says "Invalid credentials" after migration
Your password is the same as before. If you've forgotten it:
1. Open Command Prompt and run:
   ```
   python -c "import hashlib; print(hashlib.sha256(b'your_password_here').hexdigest())"
   ```
2. Then connect to Railway's PostgreSQL and run:
   ```sql
   UPDATE users SET password_hash='<hash from above>' WHERE username='your_username';
   ```
   You can connect via Railway's built-in query editor under the Postgres plugin → **Data** tab.

### "Data is gone after Railway restarted"
This means the app was still using SQLite at the time. Check `/api/health` to confirm it says `"postgresql"`. If it says `"sqlite"`, the `DATABASE_URL` environment variable was not set when the app started — set it in Railway Variables and redeploy.

---

## Architecture Overview

```
Your Browser
     │
     ▼
Railway (gh-meal.up.railway.app)
     │  FastAPI app (app.py)
     │  Reads DATABASE_URL env var
     │
     ▼
Railway PostgreSQL (managed, persistent)
     │  All data stored here permanently
     │  Survives restarts, redeployments, crashes
     └─ employees, settings, reports, prices...
```

Local Windows PC (START.bat) still uses SQLite — this is fine for
offline use. Railway always uses PostgreSQL. The two are independent.

---

*GeometryHome Workforce & Meal Management System v5.2*
*Created by Denil Joseph*
