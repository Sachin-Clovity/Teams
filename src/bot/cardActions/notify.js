import { getIssueNotifySub, saveIssueNotifySub } from '../../storage/kvsStore.js';
import { taskContinue, taskMessage } from './helpers.js';

export async function fetchNotify(issueKey) {
  const sub = (await getIssueNotifySub(issueKey)) || { notifyOnUpdate: true, notifyOnComment: true };
  return taskContinue('Channel Notifications', {
    type: 'AdaptiveCard', version: '1.2',
    body: [
      { type: 'TextBlock', text: 'Channel Notifications', weight: 'Bolder', size: 'Medium' },
      { type: 'TextBlock', text: 'Customize the type of notifications this issue sends to your configured Teams channel.', isSubtle: true, wrap: true },
      { type: 'TextBlock', text: `For work items that match: issueKey = ${issueKey}`, wrap: true, spacing: 'Medium' },
      { type: 'TextBlock', text: 'Send a message to the channel when:', weight: 'Bolder', spacing: 'Medium' },
      { type: 'Input.Toggle', id: 'notifyOnUpdate', title: 'Work item is updated', value: sub.notifyOnUpdate ? 'true' : 'false' },
      { type: 'Input.Toggle', id: 'notifyOnComment', title: 'Comment is added', value: sub.notifyOnComment ? 'true' : 'false' },
    ],
    actions: [{ type: 'Action.Submit', title: 'Save notifications', data: { cardAction: 'notify', issueKey } }],
  });
}

export async function submitNotify(issueKey, data) {
  await saveIssueNotifySub(issueKey, {
    notifyOnUpdate: data.notifyOnUpdate === 'true',
    notifyOnComment: data.notifyOnComment === 'true',
  });
  return taskMessage(`🔔 Notification settings saved for ${issueKey}`);
}
