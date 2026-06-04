import React, { useEffect, useState } from 'react';
import { invoke } from '@forge/bridge';

const s = {
  wrap:      { fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif', padding: '24px', maxWidth: '540px' },
  h1:        { fontSize: '18px', fontWeight: '700', color: '#172B4D', marginBottom: '4px' },
  sub:       { fontSize: '13px', color: '#6B778C', marginBottom: '24px' },
  label:     { display: 'block', fontSize: '11px', fontWeight: '600', color: '#344563', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: '6px' },
  select:    { width: '100%', padding: '8px 10px', border: '1px solid #DFE1E6', borderRadius: '4px', fontSize: '14px', marginBottom: '14px', background: '#fff' },
  btnBlue:   { background: '#0052CC', color: '#fff', border: 'none', borderRadius: '4px', padding: '8px 18px', fontSize: '14px', fontWeight: '500', cursor: 'pointer', marginRight: '8px' },
  btnGreen:  { background: '#00875A', color: '#fff', border: 'none', borderRadius: '4px', padding: '8px 18px', fontSize: '14px', fontWeight: '500', cursor: 'pointer', marginRight: '8px' },
  btnRed:    { background: '#DE350B', color: '#fff', border: 'none', borderRadius: '4px', padding: '8px 18px', fontSize: '14px', fontWeight: '500', cursor: 'pointer' },
  btnGrey:   { background: '#DFE1E6', color: '#97A0AF', border: 'none', borderRadius: '4px', padding: '8px 18px', fontSize: '14px', cursor: 'not-allowed', marginRight: '8px' },
  btnDark:   { background: '#172B4D', color: '#fff', border: 'none', borderRadius: '4px', padding: '8px 18px', fontSize: '14px', fontWeight: '500', cursor: 'pointer', marginRight: '8px' },
  card:      { background: '#F4F5F7', borderRadius: '6px', padding: '14px 16px', marginBottom: '16px' },
  cardTitle: { fontSize: '11px', fontWeight: '600', color: '#6B778C', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: '2px' },
  cardVal:   { fontSize: '15px', fontWeight: '600', color: '#172B4D' },
  err:       { background: '#FFEBE6', border: '1px solid #FF8F73', borderRadius: '4px', padding: '10px 14px', color: '#BF2600', fontSize: '13px', marginBottom: '14px' },
  ok:        { background: '#E3FCEF', border: '1px solid #79F2C0', borderRadius: '4px', padding: '10px 14px', color: '#006644', fontSize: '13px', marginBottom: '14px' },
  divider:   { borderTop: '1px solid #DFE1E6', margin: '20px 0' },
  step:      { fontSize: '11px', fontWeight: '700', color: '#0052CC', textTransform: 'uppercase', letterSpacing: '0.5px', marginBottom: '10px' },
  console:   { background: '#1E2333', borderRadius: '6px', padding: '12px', fontFamily: 'monospace', fontSize: '11px', color: '#A8B5D1', maxHeight: '220px', overflowY: 'auto', marginTop: '10px' },
  logLine:   { marginBottom: '4px', lineHeight: '1.5', borderBottom: '1px solid #2D3550', paddingBottom: '4px' },
  logKey:    { color: '#79C0FF' },
  logVal:    { color: '#A5D6A7' },
  logErr:    { color: '#FF7B72' },
  logTime:   { color: '#6B778C', marginRight: '8px' },
};

export default function App() {
  const [connected,   setConnected]   = useState(false);
  const [profile,     setProfile]     = useState(null);
  const [teams,       setTeams]       = useState([]);
  const [channels,    setChannels]    = useState([]);
  const [teamId,      setTeamId]      = useState('');
  const [channelId,   setChannelId]   = useState('');
  const [webhookUrl,  setWebhookUrlInput] = useState('');
  const [savedConfig, setSavedConfig] = useState(null);
  const [automationWebhookUrl, setAutomationWebhookUrl] = useState('');
  const [loading,     setLoading]     = useState(false);
  const [err,         setErr]         = useState('');
  const [msg,         setMsg]         = useState('');
  const [consoleLogs,   setConsoleLogs]   = useState([]);
  const [showConsole,   setShowConsole]   = useState(false);
  const [customMsg,     setCustomMsg]     = useState('');

  useEffect(() => {
    invoke('getConfig').then(({ config }) => { if (config) setSavedConfig(config); }).catch(() => {});
    invoke('getWebhookUrl').then(({ url }) => setAutomationWebhookUrl(url)).catch(() => {});
  }, []);

  const clear = () => { setErr(''); setMsg(''); };

  async function connect() {
    clear(); setLoading(true);
    try {
      const res = await invoke('getTeams');
      if (!res.teams?.length) { setErr('Connected but no Teams found. Check Azure Application permissions.'); return; }
      setTeams(res.teams);
      setConnected(true);
      setMsg('Connected. Select a Team below.');
    } catch (e) { setErr(e.message || 'Connection failed. Check Forge logs.'); }
    finally { setLoading(false); }
  }

  async function onTeamChange(e) {
    const id = e.target.value;
    setTeamId(id); setChannelId(''); setChannels([]);
    if (!id) return;
    try {
      const res = await invoke('getChannels', { teamId: id });
      if (res.requiresAuth) { setErr('Session expired. Reconnect.'); return; }
      setChannels(res.channels || []);
    } catch (e) { setErr(e.message || 'Failed to load channels.'); }
  }

  async function save() {
    clear();
    if (!teamId || !channelId) { setErr('Select both a Team and a Channel.'); return; }
    setLoading(true);
    try {
      const team    = teams.find(t => t.id === teamId);
      const channel = channels.find(c => c.id === channelId);
      await invoke('saveConfig', {
        teamId, channelId,
        teamName: team?.displayName || teamId,
        channelName: channel?.displayName || channelId,
        webhookUrl,
      });
      setSavedConfig({ teamId, channelId, teamName: team?.displayName || teamId, channelName: channel?.displayName || channelId });
      setMsg('Saved. Jira issue events will now post to this channel.');
    } catch (e) { setErr(e.message || 'Save failed.'); }
    finally { setLoading(false); }
  }

  async function test() {
    clear(); setLoading(true);
    try {
      const res = await invoke('testConnection');
      res.success ? setMsg('Test message sent. Check your Teams channel.') : setErr(`Test failed: ${res.error}`);
    } catch (e) { setErr(e.message || 'Test failed.'); }
    finally { setLoading(false); }
  }

  async function sendCustomMessage() {
    clear(); setLoading(true);
    try {
      const res = await invoke('sendCustomMessage', { message: customMsg });
      res.success ? setMsg('Message sent to Teams!') : setErr(`Failed: ${res.error}`);
    } catch (e) { setErr(e.message || 'Send failed.'); }
    finally { setLoading(false); }
  }

  async function checkBotDebug() {
    clear();
    try {
      const res = await invoke('getBotDebug');
      const data = res.debug;
      const lines = typeof data === 'string'
        ? [{ key: 'info', val: data, isErr: false }]
        : Object.entries(data).map(([k, v]) => ({
            key: k,
            val: typeof v === 'boolean' ? (v ? 'true' : 'false') : String(v ?? '—'),
            isErr: k === 'error' || k === 'detail',
          }));
      setConsoleLogs(prev => [
        { time: new Date().toLocaleTimeString(), lines },
        ...prev,
      ].slice(0, 20));
      setShowConsole(true);
    } catch (e) {
      setConsoleLogs(prev => [{ time: new Date().toLocaleTimeString(), lines: [{ key: 'error', val: e.message, isErr: true }] }, ...prev].slice(0, 20));
      setShowConsole(true);
    }
  }

  async function disconnect() {
    clear();
    await invoke('clearConfig').catch(() => {});
    setConnected(false); setProfile(null); setTeams([]); setChannels([]);
    setTeamId(''); setChannelId(''); setSavedConfig(null);
    setMsg('Disconnected and configuration cleared.');
  }

  return (
    <div style={s.wrap}>
      <div style={s.h1}>Microsoft Teams Connector</div>
      <div style={s.sub}> Jira issue events to a Teams channel via Microsoft Graph API.</div>

      {err && <div style={s.err}>{err}</div>}
      {msg && <div style={s.ok}>{msg}</div>}

      <div style={s.step}>Step 1 — Connect to Microsoft Teams</div>
      {!connected ? (
        <button style={loading ? s.btnGrey : s.btnBlue} onClick={connect} disabled={loading}>
          {loading ? 'Connecting…' : 'Connect Microsoft Teams'}
        </button>
      ) : (
        <>
          <div style={s.card}>
            <div style={s.cardTitle}>Status</div>
            <div style={s.cardVal}>Connected</div>
          </div>

          <div style={s.divider} />
          <div style={s.step}>Step 2 — Select Destination</div>

          <label style={s.label}>Team</label>
          <select style={s.select} value={teamId} onChange={onTeamChange}>
            <option value=''>— choose a team —</option>
            {teams.map(t => <option key={t.id} value={t.id}>{t.displayName}</option>)}
          </select>

          {channels.length > 0 && (
            <>
              <label style={s.label}>Channel</label>
              <select style={s.select} value={channelId} onChange={e => setChannelId(e.target.value)}>
                <option value=''>— choose a channel —</option>
                {channels.map(c => <option key={c.id} value={c.id}>{c.displayName}</option>)}
              </select>
            </>
          )}

          <label style={s.label}>Teams Incoming Webhook URL</label>
          <input
            style={{ ...s.select, marginBottom: '14px', fontSize: '12px' }}
            type='text'
            placeholder='https://outlook.office.com/webhook/...'
            value={webhookUrl}
            onChange={e => setWebhookUrlInput(e.target.value)}
          />
          <div style={{ fontSize: '12px', color: '#6B778C', marginBottom: '12px' }}>
            Teams channel → ··· → Connectors → Incoming Webhook → Configure → Copy URL
          </div>

          <button
            style={(!teamId || !channelId || !webhookUrl || loading) ? s.btnGrey : s.btnBlue}
            onClick={save} disabled={!teamId || !channelId || !webhookUrl || loading}
          >
            {loading ? 'Saving…' : 'Save Configuration'}
          </button>

          <button style={{ ...s.btnRed, marginTop: '8px' }} onClick={disconnect}>Disconnect</button>
        </>
      )}

      {savedConfig && (
        <>
          <div style={s.divider} />
          <div style={s.step}>Step 3 — Active Configuration</div>
          <div style={s.card}>
            <div style={{ marginBottom: '10px' }}>
              <div style={s.cardTitle}>Team</div>
              <div style={s.cardVal}>{savedConfig.teamName}</div>
            </div>
            <div>
              <div style={s.cardTitle}>Channel</div>
              <div style={s.cardVal}>{savedConfig.channelName}</div>
            </div>
          </div>
          <button style={loading ? s.btnGrey : s.btnGreen} onClick={test} disabled={loading}>
            {loading ? 'Sending…' : 'Send Test Message to Teams'}
          </button>
          <div style={s.divider} />
          <div style={s.step}>Send Custom Message to Teams</div>
          <textarea
            style={{ ...s.select, height: '72px', resize: 'vertical', fontFamily: 'inherit' }}
            placeholder='Type your message here...'
            value={customMsg}
            onChange={e => setCustomMsg(e.target.value)}
          />
          <button
            style={(!customMsg.trim() || loading) ? s.btnGrey : s.btnBlue}
            onClick={sendCustomMessage}
            disabled={!customMsg.trim() || loading}
          >
            {loading ? 'Sending…' : 'Send to Teams'}
          </button>

          <div style={s.divider} />
          <button style={{ ...s.btnDark, marginTop: '8px' }} onClick={checkBotDebug}>
            Refresh Bot Console
          </button>
          <button style={{ ...s.btnGrey, marginTop: '8px', cursor: 'pointer', color: '#344563' }} onClick={() => { setConsoleLogs([]); setShowConsole(false); }}>
            Clear
          </button>

          {webhookUrl && (
            <>
              <div style={s.divider} />
              <div style={s.step}>Feature 6 — Automation Webhook URL</div>
              <div style={s.card}>
                <div style={s.cardTitle}>Use this URL in Jira Automation rules</div>
                <div style={{ fontSize: '12px', color: '#172B4D', wordBreak: 'break-all', marginTop: '6px', fontFamily: 'monospace' }}>
                  {automationWebhookUrl}
                </div>
              </div>
              <div style={{ fontSize: '12px', color: '#6B778C' }}>
                In Jira Automation: Add action → Send web request → POST to this URL<br/>
                Body: <code>{`{"title":"{{issue.summary}}","issueKey":"{{issue.key}}"}`}</code>
              </div>
            </>
          )}
        </>
      )}
      {showConsole && (
        <>
          <div style={s.divider} />
          <div style={s.step}>Bot Debug Console</div>
          <div style={s.console}>
            {consoleLogs.length === 0 && <div style={{ color: '#6B778C' }}>No activity yet. Send a message to the bot first.</div>}
            {consoleLogs.map((entry, i) => (
              <div key={i} style={{ marginBottom: '10px' }}>
                <div style={s.logTime}>── {entry.time} ──────────────</div>
                {entry.lines.map((line, j) => (
                  <div key={j} style={s.logLine}>
                    <span style={line.isErr ? s.logErr : s.logKey}>{line.key}: </span>
                    <span style={line.isErr ? s.logErr : s.logVal}>{line.val}</span>
                  </div>
                ))}
              </div>
            ))}
          </div>
        </>
      )}
    </div>
  );
}
