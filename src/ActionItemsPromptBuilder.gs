/**
 * ActionItemsPromptBuilder.gs - Prompt builders for action item extraction.
 */

/**
 * Builds the action item extraction prompt for Claude.
 *
 * @param {string} emailBody - The plain text email body
 * @param {Object[]} collaborators - Array of collaborator objects
 * @returns {string} Prompt text
 */
function buildActionItemsExtractionPrompt(emailBody, collaborators) {
  const collaboratorsJson = JSON.stringify(collaborators.map(c => ({
    id: c.id,
    name: c.full_name || c.name,
    email: c.email
  })));

  const today = new Date().toISOString().split('T')[0];

  return `You are a specialized data processing tool designed to extract action items from meeting summary emails.

Here is the meeting summary email:
---
${emailBody}
---

Here are the project collaborators who can be assigned tasks:
${collaboratorsJson}

Today's date is: ${today}

### Your Task:
1. Find all action items mentioned in the email (usually in a numbered list or "Action Items" section)
2. For each action item, extract:
   - title: A concise title (max 100 chars)
   - description: The full action item text
   - assignee_id: Match the assignee name to a collaborator ID, or null if no match
   - assignee_name: The name mentioned in the action item, or null
   - due_date: In YYYY-MM-DD format. Use context clues like "next Monday", "by Friday". If no date specified, set to one week from today.

### Output Format:
Return ONLY valid JSON (no markdown, no explanation):
{
  "tasks": [
    {
      "title": "...",
      "description": "...",
      "assignee_id": "...",
      "assignee_name": "...",
      "due_date": "YYYY-MM-DD"
    }
  ]
}

If no action items found, return: {"tasks": []}`;
}
