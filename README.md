# Client Management Automation System

## Overview

This repository defines a Google Apps Script based automation system that manages client meetings, email organization, agenda generation, and daily and weekly outlook reporting. This document is a specification. It is intended to be used as a baseline for automated code generation. All behaviors described below are required unless explicitly marked as Future.

The content below is derived directly from the original system plan and preserves all functional requirements, data structures, triggers, integrations, and processing steps. Structure has been normalized into a conventional README order without adding new concepts or interpretive language.

---

## Core Capabilities

### Meeting Automation

When a meeting ends, the system processes meeting data from Fathom, drafts a meeting summary email, and later creates tasks and updates records after the email is sent.

Required behaviors:

* Fathom sends a webhook containing meeting data when a meeting ends
* The system identifies the client from meeting participants
* A draft email is created containing the meeting summary and action items
* After the draft is manually sent:

  * Action items are extracted and created as Todoist tasks in a client specific project
  * Meeting notes are appended to the client’s running Google Doc

---

### Email Sorting

Email organization is handled using Gmail labels and filters that are automatically created and maintained for each client.

Required behaviors:

* Gmail filters are generated for all clients listed in the Client_Registry sheet
* Incoming client emails are automatically labeled by client
* Meeting summaries, meeting agendas, and internal briefing emails are labeled consistently

---

### Meeting Agenda Generation

The system generates agendas for upcoming meetings that do not already have one.

Required behaviors:

* The system checks for upcoming meetings on an hourly schedule during business hours
* If a meeting does not already have a generated agenda, one is created
* Agenda generation uses:

  * Outstanding Todoist tasks
  * Recent client emails
  * Previous meeting notes
  * Action items discussed previously but not yet captured in Todoist
* The generated agenda is emailed to the user and appended to the client’s Google Doc

---

### Daily and Weekly Outlooks

The system produces daily and weekly outlook reports summarizing meetings, tasks, and conflicts.

Required behaviors:

* A daily outlook is generated every morning
* A weekly outlook is generated every Monday morning
* Reports are organized by client
* Schedule conflicts are detected
* Missing agendas and overdue prerequisite tasks are identified

---

## System Architecture

The system runs entirely on Google Apps Script using Google Workspace services and external APIs. Processing runs serverless through Google’s infrastructure, with all structured data stored in Google Sheets.

Technologies and required integrations:

* Google Apps Script
* Google Sheets (all data storage)
* Gmail API (draft creation, sending, search, labeling, and filter and label management)
* Google Calendar API (fetch events)
* Google Docs API (append agendas and notes to running docs)
* Todoist REST API (create and query tasks)
* Anthropic Claude API (generate meeting agendas)
* Fathom webhooks (meeting end event ingestion via Apps Script Web App endpoint)

---

## Data Storage

All persistent data is stored in a single Google Sheets file. Each dataset is stored in a dedicated tab.

### Client_Registry

Purpose: Central registry for client identification and routing.

Columns:

* client_id
* client_name
* email_domains (comma separated)
* contact_emails (comma separated)
* google_doc_url
* todoist_project_id

---

### Generated_Agendas

Purpose: Prevent duplicate agenda generation.

Columns:

* event_id
* event_title
* client_id
* generated_timestamp

---

### Processing_Log

Purpose: Audit trail of all automated actions.

Columns:

* timestamp
* action_type
* client_id
* details
* status

---

### Unmatched

Purpose: Track meetings or emails that could not be matched to a client.

Columns:

* timestamp
* item_type (meeting or email)
* item_details
* participant_emails
* manually_resolved

---

## Client Identification

Client identification is shared across all system modules.

Inputs:

* Meeting participant email addresses
* Email sender and recipient addresses
* Calendar event guest lists

Matching order:

1. Exact match against contact_emails
2. Domain match against email_domains

The first match is used. If no match is found, the item is logged to the Unmatched sheet and processing stops.

Example:

* Meeting participants: [john@acmecorp.com](mailto:john@acmecorp.com), [sarah@gmail.com](mailto:sarah@gmail.com)
* Client A email_domains: acmecorp.com
* Client B contact_emails: [sarah@gmail.com](mailto:sarah@gmail.com)
* Result: Client A is selected

---

## Meeting Automation Details

### Trigger

Fathom sends an HTTP POST request to the Apps Script Web App endpoint when a meeting ends.

### Webhook Payload

Provided fields:

* meeting_title
* meeting_date
* transcript
* summary
* action_items (description, assignee, due_date)
* participants (name and email)

### Processing Steps

1. Receive webhook and parse JSON payload
2. Identify client from participant emails
3. If no client is found, log to Unmatched and exit
4. Create a Gmail draft

   * Subject: Meeting Summary - [meeting_title] - [date]
   * Body includes meeting title, date, summary, and numbered action items
5. Monitor sent mail for meeting summary emails
6. When sent:

   * Extract action items
   * Create Todoist tasks in the client project
   * Append meeting notes to the client’s running Google Doc

---

## Email Labels and Filters

### Client Labels

For each client:

* Client: [client_name]
* Client: [client_name]/Meeting Summaries
* Client: [client_name]/Meeting Agendas

### Briefing Labels

* Brief: Daily
* Brief: Weekly

### Filter Behavior

Filters are generated daily based on Client_Registry, and updated if email domains or contacts change.

#### Client Filter

For each client, create a Gmail filter that labels incoming messages from the client.

* Criteria (built from Client_Registry):

  * from:(*@domain1 OR *@domain2 OR [contact1@email.com](mailto:contact1@email.com) OR [contact2@email.com](mailto:contact2@email.com))
* Action:

  * Apply label: Client: [client_name]

#### Meeting Summary Sub Filter

For each client, create a Gmail filter that labels meeting summary emails related to that client.

* Criteria:

  * from:me
  * subject:'Meeting Summary'
  * to:(*@domain1 OR *@domain2 OR [contact1@email.com](mailto:contact1@email.com) OR [contact2@email.com](mailto:contact2@email.com))
* Action:

  * Apply label: Client: [client_name]/Meeting Summaries

#### Meeting Agenda Sub Filter

For each client, create a Gmail filter that labels agendas sent to self for that client.

* Criteria:

  * from:me
  * to:me
  * subject:'Agenda: [client_name]'
* Action:

  * Apply label: Client: [client_name]/Meeting Agendas

#### Global Internal Filters

Create the following Gmail filters to label internal briefing emails.

Daily Outlook Filter:

* Criteria: from:me to:me subject:'Daily Outlook'
* Action: Apply label 'Brief: Daily'

Weekly Outlook Filter:

* Criteria: from:me to:me subject:'Weekly Outlook'
* Action: Apply label 'Brief: Weekly'

#### Post Send Label Application for Meeting Summaries

After a meeting summary draft is manually sent, the system must also apply the sub label to the sent email using the Gmail API:

* Apply label: Client: [client_name]/Meeting Summaries

This ensures the sub label is applied even before Gmail filter processing.

---

## Meeting Agenda Generation Details

### Trigger

Time based trigger runs every hour from 8:00 AM to 6:00 PM.

### Processing Steps

1. Fetch today’s calendar events between now and end of day (11:59:59 PM)
2. For each event:

   * If event_id exists in Generated_Agendas, skip
   * Otherwise continue
3. Identify client from event guest list using Client_Registry matching

   * If no match, log to Unmatched and skip
4. Gather context

Todoist tasks:

* Call Todoist API to fetch tasks for the client’s project_id
* Include tasks due today or overdue

Recent emails:

* Build Gmail search query from client email_domains and contact_emails
* Add newer_than:7d
* Limit to 20 threads
* For each thread, extract subject, sender, date, and the first 500 characters of the body

Meeting history:

* Fetch the client’s running Google Doc content
* Extract the most recent "Meeting Notes - [date]" section
* Compare action oriented items in that section against current Todoist tasks
* Identify "Action Items from Last Meeting Not in Todoist"

5. Generate agenda with Claude

* Call Anthropic API POST /v1/messages
* Model: claude-sonnet-4-20250514
* Max tokens: 1000
* Prompt includes:

  * Meeting title
  * Client name
  * Meeting date and time
  * Outstanding Todoist tasks
  * Recent email activity
  * Previous meeting notes context
  * Action items from last meeting not in Todoist
  * Instructions to produce a concise agenda with time allocations

6. Send agenda email to self

* Recipient: the user’s own email address
* Subject: Agenda: [event_title]
* Body:

  * Meeting: [event_title]
  * Date/Time: [formatted_datetime]
  * Blank line
  * Agenda content from Claude
  * Blank line
  * "Please review and let me know if there are any additional topics you'd like to discuss."

7. Append agenda to the client’s running Google Doc

* Append to end of doc:

  * "Meeting Agenda - [formatted_date]"
  * Agenda content
  * Blank paragraph

8. Record generation

* Append to Generated_Agendas:

  * event_id
  * event_title
  * client_id
  * current timestamp

---

## Daily Outlook

### Trigger

Runs daily at 7:00 AM.

### Processing Steps

* Fetch today’s calendar events
* Identify clients per event
* Fetch Todoist tasks due today or earlier
* Organize meetings and tasks by client
* Detect schedule conflicts
* Identify missing agendas and prerequisite tasks
* Create or send an email titled: Daily Outlook - [date]

---

## Weekly Outlook

### Trigger

Runs every Monday at 7:00 AM.

### Processing Steps

* Fetch calendar events for the week
* Fetch Todoist tasks due this week or earlier
* Organize by client and by day
* Detect conflicts
* Identify missing agendas
* Create or send an email titled: Weekly Outlook - Week of [date]

---

## Configuration

### Script Properties

API keys and tokens must be stored in Apps Script Properties Service:

* TODOIST_API_TOKEN
* CLAUDE_API_KEY

### Web App Deployment

Apps Script must be deployed as a Web App to receive Fathom webhooks:

* Execute as: User accessing the app
* Access: Anyone

The deployment URL is used for Fathom webhook configuration.

---

## Triggers

Manual trigger creation is required for:

* Label and filter creation: daily at 6:00 AM
* Sent meeting summary monitor: every 10 minutes
* Agenda generation: every 1 hour, limited to 8:00 AM to 6:00 PM
* Daily outlook: daily at 7:00 AM
* Weekly outlook: weekly on Monday at 7:00 AM

---

## Code Organization

Source files:

* Code.gs: Entry points and trigger handlers
* ClientIdentification.gs: Client matching logic
* MeetingAutomation.gs: Fathom webhook handling and post send processing
* EmailSorting.gs: Gmail label and filter management
* AgendaGeneration.gs: Agenda creation logic
* OutlookReports.gs: Daily and weekly outlook logic
* Utilities.gs: Shared helpers

---

## Future Enhancements

* Extract action items from ongoing email threads
* Confidence scoring for extracted action items
* Manual review queue for low confidence extractions

---

## Setup Sequence

1. Create the Google Sheets file with required tabs
2. Create a standalone Apps Script project
3. Add all source files
4. Configure Script Properties
5. Deploy as Web App
6. Configure Fathom webhook
7. Create required triggers
8. Populate Client_Registry
9. Run label and filter creation once manually
10. Monitor Processing_Log and Unmatched sheets
