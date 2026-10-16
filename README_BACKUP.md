# Daily backup to Google Drive (free, no Google Cloud, no card)

Every day at 11:00 PM Dubai time the app uploads a full Excel backup to a private Drive folder
and deletes the oldest ones beyond the number you choose to keep.
This is completely separate from any other app: it has its own Drive folder, its own Apps Script
project, its own web app URL and its own secret token.

1. In Google Drive create a **new private** folder, e.g. "GH Meal Backups". Copy the id from its link
   (`https://drive.google.com/drive/folders/<THIS-PART>`).
2. Go to https://script.google.com > **New project**. Replace the code with `apps_script/Code.gs`.
   Set `FOLDER_ID` to the folder id above and set `TOKEN` to a long random secret
   (e.g. 40 random characters, different from any other app's).
3. Click **Deploy > New deployment > Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**   (the secret token protects it; the folder itself stays private)
   Authorize when asked, then copy the **Web app URL** (ends with `/exec`).
4. Give the app the same secret as an environment variable named `GDRIVE_WEBAPP_TOKEN`:
   - **Railway:** Variables tab > add `GDRIVE_WEBAPP_TOKEN`.
   - **Windows laptop:** add `set GDRIVE_WEBAPP_TOKEN=your-secret` on a line above `python app.py` in
     `START.bat` (or in whatever script/Task Scheduler action launches the app).
   Restart the app afterwards.
5. In the app: **Settings > Daily Backup to Google Drive** > paste the Web app URL, tick the switch,
   press **Save**, then press **Upload a test backup now**. A file named
   `GH_MealBackup_<date>_<time>.xlsx` should appear in the Drive folder.

If you edit the script later, use Deploy > Manage deployments > Edit > New version
(the URL stays the same).

The existing in-app backups (2 AM, last 14 kept in Settings) and optional email delivery are unchanged.
The Drive upload is an additional off-site copy.
