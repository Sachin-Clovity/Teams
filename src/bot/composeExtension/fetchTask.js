import { buildCreateIssueSiteCard, buildCommentCard, buildLogTimeCard } from '../cards.js';

// Message actions: fetchTask — route by commandId, return the adaptive-card form.
export async function handleFetchTask(body) {
  const commandId  = body.value?.commandId || 'createJiraIssue';
  const rawContent = body.value?.messagePayload?.body?.content || '';
  const messageText = rawContent.replace(/<[^>]+>/g, '').trim().slice(0, 500);
  console.log('[composeExtension] fetchTask commandId:', commandId, '| text preview:', messageText.slice(0, 60));

  let taskTitle, cardContent;
  if (commandId === 'commentInJira') {
    taskTitle   = 'Comment in Jira';
    cardContent = buildCommentCard(messageText);
  } else if (commandId === 'logTimeInJira') {
    taskTitle   = 'Log Time in Jira';
    cardContent = buildLogTimeCard(messageText);
  } else {
    // Step 1 of the create-issue wizard — Site → Project → Type/Summary/Description
    taskTitle   = 'Create a work item';
    const siteId   = process.env.JIRA_BASE_URL || 'site';
    const siteName = (process.env.JIRA_BASE_URL || '').replace(/^https?:\/\//, '').split('.')[0] || 'Jira site';
    cardContent = buildCreateIssueSiteCard(messageText, siteId, siteName);
  }

  return {
    statusCode: 200,
    headers: { 'Content-Type': ['application/json'] },
    body: JSON.stringify({
      task: {
        type: 'continue',
        value: {
          title: taskTitle,
          height: 'medium',
          width: 'medium',
          card: {
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: cardContent,
          },
        },
      },
    }),
  };
}
