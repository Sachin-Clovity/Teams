// Serves the actual tab UI shown inside Teams (channel/team tab, meeting side panel, and
// personal tab all reuse this same page — they differ only in the query string Teams passes it).
// NOTE: TAB_DATA_URL is a placeholder — filled in with the real deployed webtrigger URL
// after first deploy (Forge webtrigger URLs aren't known until the app is deployed once).
const TAB_DATA_URL = 'https://1e22dbd0-9cce-44ac-b3b7-03d8530fb34e.hello.atlassian-dev.net/x1/zEKcEvBfFUDGAkIC_4lyhMICqsE';

export async function teamsTabContent(req) {
  const html = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Jira Issues</title>
<script src="https://res.cdn.office.net/teams-js/2.19.0/js/MicrosoftTeams.min.js"></script>
<style>
  body { font-family: -apple-system, "Segoe UI", sans-serif; margin: 0; padding: 16px; background: #fafbfc; color: #172B4D; }
  h1 { font-size: 15px; margin: 0 0 12px; font-weight: 600; }
  .issue { display: flex; align-items: center; padding: 10px 12px; border: 1px solid #DFE1E6; border-radius: 6px; margin-bottom: 8px; background: white; text-decoration: none; color: inherit; }
  .issue:hover { border-color: #0052CC; }
  .key { font-weight: 600; color: #0052CC; margin-right: 10px; white-space: nowrap; }
  .summary { flex: 1; font-size: 13px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .badge { font-size: 11px; padding: 2px 8px; border-radius: 10px; background: #DFE1E6; color: #42526E; margin-left: 6px; white-space: nowrap; }
  .empty, .error, .loading { color: #6B778C; font-size: 13px; padding: 24px 0; text-align: center; }
</style>
</head>
<body>
  <h1>Jira Issues</h1>
  <div id="root" class="loading">Loading…</div>
  <script>
    function qs(name) { return new URLSearchParams(window.location.search).get(name); }

    function render(issues, error) {
      var root = document.getElementById('root');
      if (error) { root.className = 'error'; root.textContent = 'Error: ' + error; return; }
      if (!issues.length) { root.className = 'empty'; root.textContent = 'No issues found.'; return; }
      root.className = '';
      root.innerHTML = issues.map(function (i) {
        return '<a class="issue" href="' + i.url + '" target="_blank" rel="noopener">' +
          '<span class="key">' + i.key + '</span>' +
          '<span class="summary">' + i.summary + '</span>' +
          '<span class="badge">' + i.status + '</span>' +
          '<span class="badge">' + i.priority + '</span>' +
          '</a>';
      }).join('');
    }

    function fetchAndRender(url) {
      fetch(url)
        .then(function (r) { return r.json(); })
        .then(function (data) { render(data.issues || [], data.error); })
        .catch(function (e) { render([], e.message); });
    }

    function start() {
      var jql = qs('jql');
      if (jql) { fetchAndRender('${TAB_DATA_URL}?jql=' + encodeURIComponent(jql)); return; }

      if (window.microsoftTeams) {
        microsoftTeams.app.initialize()
          .then(function () { return microsoftTeams.app.getContext(); })
          .then(function (ctx) {
            var email = ctx.user && ctx.user.userPrincipalName;
            fetchAndRender('${TAB_DATA_URL}' + (email ? ('?email=' + encodeURIComponent(email)) : ''));
          })
          .catch(function () { fetchAndRender('${TAB_DATA_URL}'); });
      } else {
        fetchAndRender('${TAB_DATA_URL}');
      }
    }

    start();
  </script>
</body>
</html>`;

  return {
    statusCode: 200,
    headers: {
      'Content-Type': ['text/html; charset=utf-8'],
      // Without this, browsers refuse to let Teams iframe this page at all ("refused to connect").
      // No X-Frame-Options here — it can't express multiple allowed origins, and CSP overrides it anyway.
      'Content-Security-Policy': ["frame-ancestors https://teams.microsoft.com https://*.teams.microsoft.com https://teams.live.com https://*.skype.com https://*.cloud.microsoft;"],
    },
    body: html,
  };
}
