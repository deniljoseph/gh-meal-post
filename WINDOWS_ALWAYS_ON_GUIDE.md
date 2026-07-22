# Running GH Meal Management System 24/7 on a Windows Laptop

This guide sets up your laptop so the app:
- Starts automatically when the laptop turns on (even before anyone logs in)
- Restarts automatically if it crashes
- Keeps running even if you close the lid / lock the screen / log out
- Never goes to sleep and interrupts the server
- Is reachable from other devices on your office WiFi (phones, other PCs)

Follow every section in order. Total time: ~20 minutes.

---

## Part 1 — Install Python (5 min)

1. Go to https://python.org/downloads and download the latest **Python 3.11 or 3.12** installer (do not use 3.13+ yet, some packages lag behind).
2. Run the installer.
3. **Important:** On the first screen, tick **"Add python.exe to PATH"** at the bottom before clicking Install.
4. Click **Install Now**.
5. When done, open Command Prompt (press `Win`, type `cmd`, press Enter) and run:
   ```
   python --version
   ```
   You should see something like `Python 3.12.x`. If you see an error, restart your laptop and try again (PATH changes need a restart to take effect sometimes).

---

## Part 2 — Set Up the App Folder (3 min)

1. Extract the `workforce_app` folder from the zip to a permanent location, e.g.:
   ```
   C:\GHMealSystem\workforce_app
   ```
   Avoid Desktop or Downloads — use a folder that won't accidentally get deleted or moved.

2. Open Command Prompt and navigate there:
   ```
   cd C:\GHMealSystem\workforce_app
   ```

3. Install the required packages once:
   ```
   pip install -r requirements.txt
   ```
   This takes 1–2 minutes on first run.

4. Test it works by running:
   ```
   python app.py
   ```
   You should see:
   ```
   ======================================================
     GeometryHome WFMS v5.2  |  DB: SQLite ⚠️  (local only)
     Listening on port 8000
   ======================================================
   ```
   Open a browser and go to `http://localhost:8000` — you should see the login/setup page.
   Press `Ctrl+C` in Command Prompt to stop it for now — we'll set it up to run automatically next.

---

## Part 3 — Prevent the Laptop From Sleeping (3 min)

If the laptop sleeps, the server stops responding until someone wakes it. Disable sleep completely since this machine's only job is running the server.

1. Press `Win`, type **"Power & sleep settings"**, press Enter.
2. Under **Screen**, set both dropdowns to whatever you like (screen can turn off, that's fine and saves the display).
3. Under **Sleep**, set **both** "When plugged in, PC goes to sleep after" and "On battery" (if a laptop) to **Never**.
4. Also open **Control Panel → Power Options → Choose when to turn off display**:
   - Click **"Change advanced power settings"**
   - Expand **Sleep** → **Sleep after** → set to **Never** for both "On battery" and "Plugged in"
   - Expand **Sleep** → **Allow hybrid sleep** → **Off**
   - Click **Apply → OK**
5. **Also disable "Allow the computer to sleep" in the laptop's lid-close behavior** if applicable:
   - Control Panel → Power Options → **"Choose what closing the lid does"**
   - Set **"When I close the lid"** (both On battery and Plugged in) to **"Do nothing"**
   - Click **Save changes**

   This lets you close the laptop lid without it sleeping — important if it needs to sit closed in a cabinet or drawer.

6. Make sure the laptop stays plugged into power permanently (not running on battery).

---

## Part 4 — Set Up PostgreSQL Locally (Recommended) or Keep SQLite

**Option A — Simple (SQLite, good for a single-laptop setup):**
Nothing to do — the app already uses SQLite (`workforce.db` in the app folder) when no `DATABASE_URL` is set. This is fine for one laptop that's always on, since the data lives on that machine's disk. Just make sure you use the built-in Backup feature (Settings → Backup & Restore) regularly, or rely on the automatic daily 2 AM backup already built into the app.

**Option B — More robust (local PostgreSQL):**
If you want database-grade reliability even on this single machine, install PostgreSQL locally:
1. Download from https://www.postgresql.org/download/windows/
2. During install, set a password for the `postgres` user and remember it.
3. After install, open **pgAdmin** (installed alongside) or use Command Prompt to create a database:
   ```
   "C:\Program Files\PostgreSQL\16\bin\createdb.exe" -U postgres gh_meal
   ```
4. Set an environment variable so the app uses it (see Part 5, add this alongside `SECRET_KEY`):
   ```
   DATABASE_URL=postgresql://postgres:YOUR_PASSWORD@localhost:5432/gh_meal
   ```

Most people running a single always-on laptop can safely stick with **Option A (SQLite)** — it's simpler and the automatic backups already protect you.

---

## Part 5 — Set Environment Variables (2 min)

1. Press `Win`, type **"environment variables"**, choose **"Edit the system environment variables"**.
2. Click **"Environment Variables..."** button.
3. Under **"System variables"**, click **New...** and add:
   | Variable name | Value |
   |---|---|
   | `SECRET_KEY` | any long random text, e.g. `gh-meal-2026-xK9mP3nQ7rL2wA8vB5` |
   | `PORT` | `8000` |

   (Skip `DATABASE_URL` entirely if you're using Option A / SQLite above.)

4. Click **OK** on all windows to save.

---

## Part 6 — Run the App Automatically at Startup, Even Before Login (10 min)

We'll use **Windows Task Scheduler** (built into Windows, no extra downloads needed) to run the app as a background task that:
- Starts the moment Windows boots (before anyone logs in)
- Restarts automatically if it ever crashes
- Runs silently with no visible window

### Step 6.1 — Create a launcher script

Create a new file at `C:\GHMealSystem\workforce_app\run_silent.vbs` with this exact content (use Notepad):

```vbscript
Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = "C:\GHMealSystem\workforce_app"
objShell.Run "cmd /c python app.py >> app_log.txt 2>&1", 0, False
```

This runs the app with **no visible console window** and logs all output to `app_log.txt` in the same folder — useful for troubleshooting later.

### Step 6.2 — Open Task Scheduler

1. Press `Win`, type **"Task Scheduler"**, press Enter.
2. Click **"Create Task..."** (not "Create Basic Task" — we need the advanced options).

### Step 6.3 — General tab
- **Name:** `GH Meal System`
- **Description:** `Runs the GeometryHome Meal Management server`
- Select **"Run whether user is logged on or not"**
- Tick **"Run with highest privileges"**
- Under "Configure for:", choose your Windows version (e.g. Windows 10)

### Step 6.4 — Triggers tab
1. Click **New...**
2. Begin the task: **"At startup"**
3. Tick **"Delayed task for:"** → **30 seconds** (gives Windows time to get network ready)
4. Click **OK**

### Step 6.5 — Actions tab
1. Click **New...**
2. Action: **"Start a program"**
3. Program/script: `wscript.exe`
4. Add arguments: `"C:\GHMealSystem\workforce_app\run_silent.vbs"`
5. Click **OK**

### Step 6.6 — Conditions tab
- **Untick** "Start the task only if the computer is on AC power" (or leave it ticked if the laptop is always plugged in anyway — but unticking is safer)
- **Untick** "Stop if the computer switches to battery power"

### Step 6.7 — Settings tab
- Tick **"Allow task to be run on demand"**
- Tick **"If the task fails, restart every:"** → set to **1 minute**, attempt restart **999 times** (effectively forever)
- Tick **"If the running task does not end when requested, force it to stop"**
- Under "If the task is already running...", choose **"Do not start a new instance"**

### Step 6.8 — Save
Click **OK**. You'll be asked for the Windows account password (needed for "run whether user is logged on or not"). Enter it and confirm.

### Step 6.9 — Test it
1. Right-click the **"GH Meal System"** task in the list → **Run**
2. Wait 5 seconds, then open a browser and go to `http://localhost:8000`
3. You should see the app load. If not, check `app_log.txt` in the app folder for error messages.
4. **Restart the laptop fully** to confirm it comes back up automatically without anyone logging in — this is the real test.

---

## Part 7 — Make It Reachable From Other Devices on the Network (5 min)

Right now the app only responds to `localhost` on the laptop itself. To let phones/other PCs on the same WiFi access it:

### Step 7.1 — Find the laptop's local IP address
1. Open Command Prompt, run:
   ```
   ipconfig
   ```
2. Look for **"IPv4 Address"** under your active network adapter (Wi-Fi or Ethernet), e.g. `192.168.1.45`.

### Step 7.2 — Allow the app through Windows Firewall
1. Press `Win`, type **"Windows Defender Firewall with Advanced Security"**, press Enter.
2. Click **"Inbound Rules"** (left panel) → **"New Rule..."** (right panel).
3. Rule Type: **Port** → Next
4. **TCP**, Specific local ports: `8000` → Next
5. **Allow the connection** → Next
6. Tick all three (Domain, Private, Public) → Next
7. Name: `GH Meal System Port 8000` → Finish

### Step 7.3 — Access from another device
On any phone or PC connected to the **same WiFi network**, open a browser and go to:
```
http://192.168.1.45:8000
```
(replace with your laptop's actual IP from Step 7.1)

### Step 7.4 — (Optional) Give the laptop a fixed IP so it never changes
By default, your router might assign a different IP to the laptop each time it reconnects, breaking the link above. To fix this:
1. Log into your WiFi router's admin page (usually `192.168.1.1` or `192.168.0.1` in a browser)
2. Find **"DHCP reservation"** or **"Static IP binding"** (varies by router brand)
3. Reserve the laptop's current IP permanently based on its MAC address (shown in `ipconfig /all` as "Physical Address")

Ask whoever manages your office WiFi if you're not sure how to access the router.

---

## Part 8 — Ongoing Maintenance

### Checking if it's running
- Task Scheduler → find "GH Meal System" in the list → check the **"Last Run Result"** column shows `(0x0)` (success)
- Or just visit `http://localhost:8000/api/health` in a browser — should show:
  ```json
  {"status": "ok", "database": "sqlite", "version": "5.2"}
  ```

### Viewing logs
Open `C:\GHMealSystem\workforce_app\app_log.txt` in Notepad to see server output and any errors.

### Restarting manually
Task Scheduler → right-click "GH Meal System" → **End** then **Run** again. Or just restart the laptop.

### Updating the app (new files from Claude)
1. Stop the task: Task Scheduler → right-click "GH Meal System" → **End**
2. Replace `app.py` and/or `static/index.html` with the new versions in `C:\GHMealSystem\workforce_app\`
3. Start the task again: right-click → **Run**

### Backups
The app automatically backs up every day at 2 AM (server time) — no action needed. You can also manually trigger one anytime from **Settings → Automatic Backups → Run a Backup Now**, and download any of the last 14 backups from there. Since this laptop uses SQLite, it's also worth occasionally copying the `workforce.db` file itself to a USB drive or cloud folder (Google Drive/OneDrive) as an extra safety net, in case the laptop's hard drive ever fails.

### If the laptop needs a Windows Update restart
Windows Updates sometimes force a restart. Because we set the task to **"Run at startup, whether user is logged on or not"**, the app will come back automatically once Windows finishes rebooting — no action needed, though it's worth checking `http://localhost:8000/api/health` afterward just to confirm.

---

## Troubleshooting

**"python is not recognized"** — Python wasn't added to PATH during install. Reinstall Python and tick "Add python.exe to PATH", or manually add `C:\Users\YOUR_USERNAME\AppData\Local\Programs\Python\Python312\` to your PATH environment variable.

**Task runs but app_log.txt shows "ModuleNotFoundError"** — the packages weren't installed for the account the task runs as. Re-run `pip install -r requirements.txt` from an Administrator Command Prompt.

**Can't reach it from other devices** — double check the firewall rule (Part 7.2) and that both devices are on the *same* WiFi network (not one on WiFi and one on mobile data).

**App works but browser shows "This site can't be reached" after a Windows Update** — the scheduled task may need re-enabling if Windows reset some settings. Open Task Scheduler and confirm "GH Meal System" still shows as Ready/Enabled.

**Data seems to have disappeared** — check you're looking at the same `workforce.db` file. If someone accidentally copied the app folder to a new location and ran it from there, it'll start with a fresh empty database. Always run from the one canonical folder path.

---

*GeometryHome Workforce & Meal Management System v5.2*
*Created by Denil Joseph*
