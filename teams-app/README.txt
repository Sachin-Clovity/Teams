TEAMS APP PACKAGE — SETUP STEPS
================================

1. GET WEBTRIGGER URL
   Run: forge webtrigger -e development
   Copy the URL for "teams-bot-endpoint"

2. AZURE BOT SETUP
   - Go to portal.azure.com
   - Create resource → "Azure Bot"
   - Bot handle: jira-teams-connector
   - Use existing app: 8529f84e-4daa-48df-bd92-0f62fd47651e
   - Messaging endpoint: paste the webtrigger URL from step 1
   - Add channel: Microsoft Teams
   - Save

3. UPDATE manifest.json
   Replace all 3 occurrences of REPLACE_WITH_AZURE_BOT_APP_ID
   with: 8529f84e-4daa-48df-bd92-0f62fd47651e

4. ADD ICONS
   - color.png  → 192x192 px (Jira blue logo)
   - outline.png → 32x32 px (white/transparent logo)

5. PACKAGE THE APP
   Zip these 3 files (NOT the folder, just the files):
   - manifest.json
   - color.png
   - outline.png
   Name the zip: jira-teams-connector.zip

6. UPLOAD TO TEAMS
   - Teams → Apps → Manage your apps → Upload an app
   - Upload jira-teams-connector.zip
   - Install to a team

7. TEST
   Paste any Jira issue URL in Teams chat
   → should unfurl into a rich card
