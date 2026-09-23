import React, { useEffect, useState } from 'react';
import { invoke } from '@forge/bridge';
import { Skeleton, formatError } from '../common';
import TeamsTab from './TeamsTab';
import ActionsTab from './ActionsTab';
import ActivityTab from './ActivityTab';

const STATUS_BADGE_CLASS = { 'In Progress': 'bg-jira-blue text-white', 'Done': 'bg-jira-green text-white', 'In Review': 'bg-jira-purple text-white' };
const statusBadgeClass = status => `text-xs font-semibold px-2 py-0.5 rounded ${STATUS_BADGE_CLASS[status] || 'bg-jira-border text-jira-dark'}`;
const tabClass = active => `tab-btn ${active ? 'border-jira-blue text-jira-blue' : 'border-transparent text-jira-grey hover:text-jira-dark'}`;

// ── Issue Context Panel ───────────────────────────────────────────────────────
// Loads the issue + current user once, then hands off to whichever tab is active. Each tab
// (see the sibling files in this folder) owns its own mode/form state — switching tabs
// unmounts the inactive one, which naturally discards its state instead of needing an
// explicit reset() call threaded through every action.
export default function IssuePanel({ issueKey }) {
  const [issue,       setIssue]       = useState(null);
  const [currentUser, setCurrentUser] = useState(null);
  const [tab,         setTab]         = useState('teams');
  const [loading,     setLoading]     = useState(true);
  const [err,         setErr]         = useState('');

  useEffect(() => {
    if (!issueKey) return;
    Promise.all([
      invoke('getIssueDetails', { issueKey }),
      invoke('getCurrentUser'),
    ]).then(([issueData, userData]) => {
      setIssue(issueData);
      setCurrentUser(userData);
      setLoading(false);
    }).catch(e => { setErr(formatError(e, 'Failed to load issue.')); setLoading(false); });
  }, [issueKey]);

  if (loading) {
    return (
      <div className="p-3" aria-busy="true" aria-label="Loading">
        <div className="card mb-3">
          <Skeleton className="h-4 w-20 mb-2" />
          <Skeleton className="h-3 w-full mb-1" />
          <Skeleton className="h-3 w-2/3 mb-2" />
          <Skeleton className="h-5 w-16" />
        </div>
        <div className="flex gap-3 mb-3">
          <Skeleton className="h-6 w-14" /><Skeleton className="h-6 w-14" /><Skeleton className="h-6 w-14" />
        </div>
        <Skeleton className="h-8 w-full mb-2" />
        <Skeleton className="h-8 w-full mb-2" />
        <Skeleton className="h-8 w-full" />
      </div>
    );
  }

  return (
    <div className="p-3 font-sans text-sm">
      {issue && (
        <div className="card mb-3">
          <div className="font-bold text-jira-dark text-sm">{issue.key}</div>
          <div className="text-jira-dark mt-1 mb-2 leading-snug text-xs">{issue.summary}</div>
          <div className="flex flex-wrap gap-2 items-center">
            <span className={statusBadgeClass(issue.status)}>{issue.status}</span>
            {issue.priority && <span className="text-xs text-jira-grey">P: {issue.priority}</span>}
            {issue.issueType && <span className="text-xs text-jira-grey">{issue.issueType}</span>}
          </div>
          {issue.assignee && (
            <div className="mt-2 text-xs text-jira-grey">
              Assignee: <span className="font-semibold text-jira-dark">{issue.assignee.name}</span>
            </div>
          )}
        </div>
      )}

      <div className="flex border-b border-jira-border mb-3">
        <button className={tabClass(tab === 'teams')}    onClick={() => setTab('teams')}>Teams</button>
        <button className={tabClass(tab === 'actions')}  onClick={() => setTab('actions')}>Actions</button>
        <button className={tabClass(tab === 'activity')} onClick={() => setTab('activity')}>Activity</button>
      </div>

      {err && <div className="alert-err">{err}</div>}

      {tab === 'teams'    && <TeamsTab    issueKey={issueKey} issue={issue} currentUser={currentUser} />}
      {tab === 'actions'  && <ActionsTab  issueKey={issueKey} issue={issue} onIssueChange={setIssue} />}
      {tab === 'activity' && <ActivityTab issueKey={issueKey} />}
    </div>
  );
}
