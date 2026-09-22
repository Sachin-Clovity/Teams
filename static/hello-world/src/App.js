import React, { useEffect, useState } from 'react';
import { view } from '@forge/bridge';
import AdminPage from './components/AdminPage';
import ProjectSettingsPage from './components/ProjectSettingsPage';
import PersonalSettingsPage from './components/PersonalSettingsPage';
import IssuePanel from './components/IssuePanel';
import { ToastProvider, PageSkeleton } from './components/common';

// ── Root — routes to the right screen based on which Forge module rendered it ──
export default function App() {
  const [moduleKey,  setModuleKey]  = useState(null);
  const [issueKey,   setIssueKey]   = useState(null);
  const [projectKey, setProjectKey] = useState(null);

  useEffect(() => {
    view.getContext()
      .then(ctx => {
        setModuleKey(ctx.moduleKey || 'admin');
        setIssueKey(ctx.extension?.issue?.key || null);
        setProjectKey(ctx.extension?.project?.key || null);
      })
      .catch(() => setModuleKey('admin'));
  }, []);

  return (
    <ToastProvider>
      {!moduleKey ? <PageSkeleton /> : moduleKey === 'teams-issue-context' ? <IssuePanel issueKey={issueKey} />
        : moduleKey === 'teams-project-settings' ? <ProjectSettingsPage projectKey={projectKey} />
        : moduleKey === 'teams-personal-settings' ? <PersonalSettingsPage />
        : <AdminPage />}
    </ToastProvider>
  );
}
