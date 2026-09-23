import React, { useEffect, useState } from 'react';
import { invoke } from '@forge/bridge';
import { SectionCard, Button, useToast, LogTable, formatError } from '../common';

// ── Send Custom Message — posts free text to the default channel. Fully self-contained: no
// props, own loading/text state, uses the useToast hook directly rather than having it
// threaded down from the page. ──
function SendCustomMessageCard() {
  const showToast = useToast();
  const [customMsg, setCustomMsg] = useState('');
  const [sending,   setSending]   = useState(false);

  async function send() {
    setSending(true);
    try {
      const res = await invoke('sendCustomMessage', { message: customMsg });
      res.success ? showToast('Message sent to Teams!') : showToast(formatError({ message: res.error }, 'Send failed.'), 'error');
    } catch (e) { showToast(formatError(e, 'Send failed.'), 'error'); }
    finally { setSending(false); }
  }

  return (
    <SectionCard title="Send Custom Message" description="Posts to the default channel.">
      <textarea
        className="form-textarea w-full mb-2 h-20"
        placeholder="Type your message here..."
        value={customMsg}
        onChange={e => setCustomMsg(e.target.value)}
      />
      <Button className="w-full" onClick={send} loading={sending} disabled={!customMsg.trim()}>Send to Teams</Button>
    </SectionCard>
  );
}

// ── Automation Webhook URL — static info card, fetches its own URL once. ──
function AutomationWebhookCard() {
  const [url, setUrl] = useState('');

  useEffect(() => {
    invoke('getWebhookUrl').then(res => setUrl(res.url)).catch(() => {});
  }, []);

  return (
    <SectionCard title="Automation Webhook URL" description="Use this URL in Jira Automation rules.">
      <div className="card">
        <div className="text-xs font-mono text-jira-dark break-all bg-white border border-jira-border rounded p-2">{url}</div>
      </div>
      <p className="text-xs text-jira-grey">
        In Jira Automation: Add action → Send web request → POST to this URL<br />
        Body: <code className="bg-jira-light px-1 rounded break-all">{'{"title":"{{issue.summary}}","issueKey":"{{issue.key}}","issueUrl":"{{baseUrl}}/browse/{{issue.key}}"}'}</code>
        <br />
        This works from any Jira site — no app install needed there. Include <code className="bg-jira-light px-1 rounded">issueUrl</code> using the <code className="bg-jira-light px-1 rounded">{'{{baseUrl}}'}</code> smart value so the link points at the right site.
      </p>
    </SectionCard>
  );
}

// ── Debug Tools — MS auth check + bot activity console, own log buffer. ──
function DebugToolsCard() {
  const [consoleLogs, setConsoleLogs] = useState([]);
  const [showConsole, setShowConsole] = useState(false);

  function pushLog(lines) {
    setConsoleLogs(prev => [{ time: new Date().toLocaleTimeString(), lines }, ...prev].slice(0, 20));
    setShowConsole(true);
  }

  async function checkAuthStatus() {
    try {
      const res = await invoke('getAuthStatus');
      pushLog(res.authenticated
        ? [{ key: 'authenticated', val: 'true', isErr: false }, { key: 'org', val: res.org || '—', isErr: false }, { key: 'tenantId', val: res.tenantId || '(default — no bot traffic seen yet)', isErr: false }]
        : [{ key: 'authenticated', val: 'false', isErr: true }, { key: 'error', val: res.error || 'Token request failed', isErr: true }]);
    } catch (e) { pushLog([{ key: 'error', val: formatError(e), isErr: true }]); }
  }

  async function checkBotDebug() {
    try {
      const res  = await invoke('getBotDebug');
      const data = res.debug;
      pushLog(typeof data === 'string'
        ? [{ key: 'info', val: data, isErr: false }]
        : Object.entries(data).map(([k, v]) => ({ key: k, val: typeof v === 'boolean' ? String(v) : String(v ?? '—'), isErr: k === 'error' || k === 'detail' })));
    } catch (e) { pushLog([{ key: 'error', val: formatError(e), isErr: true }]); }
  }

  // Same idea as Bot Console, but for the jira-issue-trigger function — the only way to confirm
  // that trigger is firing (and why a notification was or wasn't sent) on sites where `forge
  // logs` CLI access is restricted (e.g. AGC).
  async function checkSyncDebug() {
    try {
      const res  = await invoke('getSyncDebug');
      const data = res.debug;
      pushLog(typeof data === 'string'
        ? [{ key: 'info', val: data, isErr: false }]
        : Object.entries(data).map(([k, v]) => ({ key: k, val: String(v ?? '—'), isErr: k === 'result' && String(v).startsWith('error:') })));
    } catch (e) { pushLog([{ key: 'error', val: formatError(e), isErr: true }]); }
  }

  return (
    <SectionCard title="Debug Tools">
      <div className="flex gap-2 flex-wrap mb-3">
        <Button className="text-xs" onClick={checkAuthStatus}>Check MS Auth</Button>
        <Button variant="dark" className="text-xs" onClick={checkBotDebug}>Bot Console</Button>
        <Button variant="dark" className="text-xs" onClick={checkSyncDebug}>Sync Console</Button>
        <Button variant="ghost" className="text-xs" onClick={() => { setConsoleLogs([]); setShowConsole(false); }}>Clear</Button>
      </div>
      {showConsole && <LogTable entries={consoleLogs} />}
    </SectionCard>
  );
}

export default function DebugSidebar() {
  return (
    <div>
      <SendCustomMessageCard />
      <AutomationWebhookCard />
      <DebugToolsCard />
    </div>
  );
}
