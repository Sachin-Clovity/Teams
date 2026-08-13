export const DEFAULT_FIELDS = ['status', 'priority', 'issueType', 'assignee', 'reporter'];

export const FIELD_DEFS = {
  status:    { label: 'Status',   get: f => f.status?.name || 'Unknown' },
  priority:  { label: 'Priority', get: f => f.priority?.name || 'None' },
  issueType: { label: 'Type',     get: f => f.issuetype?.name || 'Issue' },
  assignee:  { label: 'Assignee', get: f => f.assignee?.displayName || 'Unassigned' },
  reporter:  { label: 'Reporter', get: f => f.reporter?.displayName || 'Unknown' },
  dueDate:   { label: 'Due',      get: f => f.duedate || 'None' },
  labels:    { label: 'Labels',   get: f => (f.labels || []).join(', ') || 'None' },
};
