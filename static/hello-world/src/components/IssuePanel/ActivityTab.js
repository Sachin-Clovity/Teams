import React, { useEffect, useState } from 'react';
import { invoke } from '@forge/bridge';
import { Button, Skeleton, EmptyState } from '../common';

export default function ActivityTab({ issueKey }) {
  const [activity, setActivity] = useState(null);

  function load() {
    setActivity(null);
    invoke('getIssueActivity', { issueKey }).then(res => setActivity(res.comments || []));
  }

  useEffect(load, [issueKey]);

  return (
    <div>
      {activity === null && (
        <>
          <Skeleton className="h-10 w-full mb-2" />
          <Skeleton className="h-10 w-full" />
        </>
      )}
      {activity !== null && activity.length === 0 && <EmptyState icon="💬" title="No comments yet" />}
      {activity !== null && activity.map(c => (
        <div key={c.id} className="border-b border-jira-border pb-3 mb-3 last:border-0">
          <div className="flex justify-between items-center mb-1">
            <span className="text-xs font-semibold text-jira-dark">{c.author}</span>
            <span className="text-xs text-jira-grey">{new Date(c.created).toLocaleDateString()}</span>
          </div>
          <div className="text-xs text-jira-dark leading-relaxed">{c.text}</div>
        </div>
      ))}
      {activity !== null && activity.length > 0 && (
        <Button variant="ghost" className="text-xs mt-1" onClick={load}>Refresh</Button>
      )}
    </div>
  );
}
