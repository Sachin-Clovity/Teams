// Shown when someone adds the Jira tab to a channel, team, or meeting — lets them pick
// which issues the tab should show. TAB_CONTENT_URL is filled in post-deploy, same as teamsTabContent.js.
const TAB_CONTENT_URL = 'https://1e22dbd0-9cce-44ac-b3b7-03d8530fb34e.hello.atlassian-dev.net/x1/eeGtmdcuSsFXrMtrmSOo1SRJKtw';

export async function teamsTabConfig(req) {
  const html = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Configure Jira Tab</title>
<script src="https://res.cdn.office.net/teams-js/2.19.0/js/MicrosoftTeams.min.js"></script>
<style>
  body { font-family: -apple-system, "Segoe UI", sans-serif; padding: 20px; color: #172B4D; }
  label { font-size: 13px; font-weight: 600; display: block; margin-bottom: 6px; }
  textarea { width: 100%; min-height: 70px; font-family: ui-monospace, monospace; font-size: 13px; padding: 8px; border: 1px solid #DFE1E6; border-radius: 4px; box-sizing: border-box; }
  .presets { margin-top: 10px; display: flex; gap: 6px; flex-wrap: wrap; }
  .preset { font-size: 12px; padding: 4px 10px; border: 1px solid #DFE1E6; border-radius: 12px; background: #fafbfc; cursor: pointer; }
  .preset:hover { border-color: #0052CC; color: #0052CC; }
  .hint { font-size: 12px; color: #6B778C; margin-top: 8px; }
</style>
</head>
<body>
  <label for="jql">Jira filter (JQL)</label>
  <textarea id="jql" placeholder="project = PROJ AND statusCategory != Done ORDER BY updated DESC"></textarea>
  <div class="presets">
    <span class="preset" data-jql="statusCategory != Done ORDER BY updated DESC">Open issues</span>
    <span class="preset" data-jql="priority in (High, Highest) AND statusCategory != Done ORDER BY updated DESC">High priority</span>
    <span class="preset" data-jql="created >= -7d ORDER BY created DESC">Created this week</span>
  </div>
  <p class="hint">Shown as a tab in this channel or meeting. Leave blank to show all recently updated issues.</p>
  <script>
    document.querySelectorAll('.preset').forEach(function (el) {
      el.addEventListener('click', function () { document.getElementById('jql').value = el.getAttribute('data-jql'); });
    });

    microsoftTeams.app.initialize().then(function () {
      microsoftTeams.pages.config.setValidityState(true);
      microsoftTeams.pages.config.registerOnSaveHandler(function (saveEvent) {
        var jql = document.getElementById('jql').value.trim();
        var contentUrl = '${TAB_CONTENT_URL}' + (jql ? ('?jql=' + encodeURIComponent(jql)) : '');
        microsoftTeams.pages.config.setConfig({
          entityId: 'jira-issues-tab',
          contentUrl: contentUrl,
          websiteUrl: contentUrl,
          suggestedDisplayName: 'Jira Issues',
        });
        saveEvent.notifySuccess();
      });
    });
  </script>
</body>
</html>`;

  return {
    statusCode: 200,
    headers: {
      'Content-Type': ['text/html; charset=utf-8'],
      'Content-Security-Policy': ["frame-ancestors https://teams.microsoft.com https://*.teams.microsoft.com https://teams.live.com https://*.skype.com https://*.cloud.microsoft;"],
    },
    body: html,
  };
}
