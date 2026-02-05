# Blue Dot Consulting - Client Management Automation System

A comprehensive Google Apps Script automation system that streamlines client relationship management through intelligent meeting workflows, AI-powered agenda generation, automated email organization, and strategic daily/weekly briefings.

## Table of Contents

- [Overview](#overview)
- [Key Features](#key-features)
- [System Architecture](#system-architecture)
- [Module Breakdown](#module-breakdown)
- [Setup Instructions](#setup-instructions)
- [Configuration](#configuration)
- [User Workflows](#user-workflows)
- [Data Structure](#data-structure)
- [Automation Schedule](#automation-schedule)
- [External Integrations](#external-integrations)
- [Troubleshooting](#troubleshooting)
- [Advanced Features](#advanced-features)

---

## Overview

This system automates the complete client management lifecycle for consulting teams:

1. **Meeting Processing**: Fathom records meetings → System creates draft summaries → User sends → Tasks automatically created in Todoist → Notes appended to client docs
2. **Email Organization**: All client emails automatically labeled and filtered by client
3. **Agenda Generation**: AI generates context-aware agendas for upcoming meetings using previous notes, tasks, and recent emails
4. **Strategic Briefings**: Daily and weekly AI-generated outlook reports with schedule analysis, conflict detection, and priority recommendations

**Platform**: Google Apps Script (serverless, runs on Google infrastructure)
**Data Storage**: Google Sheets (single file with multiple tabs)
**AI Engine**: Claude 3.5 (Anthropic)

---

## Key Features

### ✅ Core Capabilities

- **Automatic Meeting Summaries**: Fathom webhooks → Draft creation → Todoist task generation → Doc updates
- **AI-Powered Agendas**: Context-aware meeting preparation using Claude 3.5
- **Smart Email Organization**: Automatic Gmail labels and filters by client
- **Task Automation**: Extract action items and create Todoist tasks with assignees
- **Running Meeting Notes**: Structured Google Docs with searchable delimiters
- **Daily Strategic Briefing**: AI-generated morning report with schedule overview and priority actions
- **Weekly Outlook**: Project dashboard with cross-client analysis and risk detection
- **Conflict Detection**: Automatic identification of overlapping meetings
- **Client Onboarding**: Automatic Google Doc and Todoist project creation for new clients
- **Drive Folder Sync**: Keep folder dropdowns in sync with Google Drive structure
- **Weather Insights**: Optional National Weather Service (US) or OpenWeather forecasts in daily outlooks

### 🎯 Advanced Features

- **Context-Aware Agendas**: Incorporates recent emails, previous notes, and outstanding tasks
- **Intelligent Action Extraction**: Claude AI identifies action items even from edited drafts
- **Smart Client Matching**: Exact contact match + domain match + internal-only client support
- **Customizable AI Prompts**: Edit system prompts and select model tiers (Haiku vs Sonnet)
- **Model Caching**: 24-hour cache of available Claude models
- **Error Recovery**: Comprehensive logging with graceful degradation
- **Fallback Templates**: System works without AI if Claude API unavailable
- **Business Hours Enforcement**: Agendas only generated 8 AM - 6 PM
- **Webhook Security**: HMAC SHA-256 signature verification for Fathom webhooks
- **Filter Safety**: System only manages its own labels (prevents accidental modifications)

---

## System Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                    GOOGLE APPS SCRIPT                        │
│                  (Serverless Execution)                      │
└─────────────────────────────────────────────────────────────┘
                              │
                              │
        ┌─────────────────────┼─────────────────────┐
        │                     │                     │
   ┌────▼────┐          ┌─────▼──────┐      ┌──────▼──────┐
   │ Webhooks│          │  Triggers  │      │ Manual Runs │
   │ (Fathom)│          │(Scheduled) │      │  (Testing)  │
   └────┬────┘          └─────┬──────┘      └──────┬──────┘
        │                     │                     │
        └─────────────────────┼─────────────────────┘
                              │
                    ┌─────────▼─────────┐
                    │   Core Modules    │
                    └─────────┬─────────┘
                              │
        ┌─────────────────────┼─────────────────────┐
        │                     │                     │
   ┌────▼────────┐    ┌───────▼────────┐    ┌──────▼──────┐
   │   Meeting   │    │  Email & Tasks │    │   Reports   │
   │  Automation │    │   Automation   │    │  & Agendas  │
   └────┬────────┘    └───────┬────────┘    └──────┬──────┘
        │                     │                     │
        └─────────────────────┼─────────────────────┘
                              │
            ┌─────────────────┼─────────────────┐
            │                 │                 │
     ┌──────▼──────┐   ┌──────▼──────┐   ┌─────▼─────┐
     │   Gmail     │   │   Sheets    │   │  Claude   │
     │   Todoist   │   │   Calendar  │   │   API     │
     │   Docs      │   │   Drive     │   │           │
     └─────────────┘   └─────────────┘   └───────────┘
```

### Data Flow: Meeting Processing

```
Fathom Meeting End
         │
         ▼
doPost(e) - Webhook Receiver
         │
         ▼
Verify HMAC Signature
         │
         ▼
Identify Client (participants)
         │
         ├─── Match Found ────────┐
         │                        │
         ▼                        ▼
Create Gmail Draft          Log to Unmatched
(with summary & actions)         Sheet
         │
         ▼
User Reviews & Sends Draft
         │
         ▼
Monitor Sent (every 10 min)
         │
         ▼
Detect Sent Summary Email
         │
         ├─── Extract Action Items (Claude AI)
         ├─── Create Todoist Tasks (with assignees)
         └─── Append to Google Doc (with delimiters)
         │
         ▼
Log to Processing_Log
```

### Data Flow: Agenda Generation

```
Hourly Trigger (8 AM - 6 PM)
         │
         ▼
Get Today's Remaining Events
         │
         ▼
For Each Event:
  │
  ├─── Check Generated_Agendas (skip if exists)
  │
  ├─── Identify Client (guest list)
  │
  ├─── Gather Context:
  │    ├─ Todoist tasks (due today/overdue)
  │    ├─ Recent emails (last 7 days)
  │    ├─ Previous meeting notes (LAST only)
  │    └─ Unmatched action items (fuzzy match)
  │
  ├─── Generate Agenda (Claude API)
  │    └─ Prompt: Meeting details + context
  │
  ├─── Send Email to Self
  │    └─ Apply label immediately
  │
  ├─── Append to Client Doc
  │    └─ Convert HTML to plain text
  │
  └─── Record in Generated_Agendas
```

---

## Module Breakdown

### **Code.gs** (Main Entry Point)

**Purpose**: Orchestration layer for all triggers and webhooks

**Key Functions**:
- `doGet(e)` - Health check endpoint
- `doPost(e)` - Fathom webhook receiver with signature verification
- `setupTriggers()` - Creates all scheduled triggers (idempotent)
- Trigger handlers for all automated processes

**Configuration Constants**:
```javascript
CONFIG = {
  SPREADSHEET_ID: // From Script Properties
  SHEETS: {
    CLIENT_REGISTRY: 'Client_Registry',
    GENERATED_AGENDAS: 'Generated_Agendas',
    PROCESSING_LOG: 'Processing_Log',
    UNMATCHED: 'Unmatched',
    FOLDERS: 'Folders'
  },
  BUSINESS_HOURS: { START: 8, END: 18 }
}
```

### **ClientIdentification.gs** (Client Matching Engine)

**Purpose**: Universal client identification across all modules

**Matching Logic**:
1. Exact match against `contact_emails` (case-insensitive)
2. Special handling for "Internal" clients (requires ALL participants to be internal)
3. Returns first match

**Key Functions**:
- `identifyClient(emails)` - Core matching function
- `identifyClientFromCalendarEvent(event)` - For agenda generation
- `identifyClientFromEmail(message)` - For email processing
- `identifyClientFromParticipants(participants)` - For Fathom webhooks
- `getClientRegistry()` - Loads all active clients
- `logUnmatched()` - Tracks unidentified items

### **MeetingAutomation.gs** (Webhook & Post-Send Processing)

**Purpose**: Complete meeting workflow automation

**Workflow**:
1. **Webhook Processing** (`processFathomWebhook`):
   - Receives Fathom payload
   - Identifies client
   - Creates Gmail draft with formatted summary
   - Stores draft metadata in cache

2. **Sent Email Monitoring** (`monitorSentMeetingSummaries`):
   - Runs every 10 minutes
   - Scans Meeting Summaries labels
   - Processes only first message in thread
   - Time-limited (last hour only)

3. **Post-Send Actions** (`processSentMeetingSummary`):
   - Extracts action items with Claude AI (handles edited drafts)
   - Creates Todoist tasks with assignees
   - Appends meeting notes to client Google Doc
   - Uses structured delimiters for parsing

**Draft Format**:
- Subject: `Meeting Summary: {client_name} - {meeting_title} ({date})`
- Body: HTML with summary, action items, and hidden metadata
- Metadata div: Draft ID, client ID, meeting date

### **EmailSorting.gs** (Gmail Label & Filter Management)

**Purpose**: Automatic email organization by client

**Label Structure**:
```
Client: [ClientName]
├── Meeting Summaries
└── Meeting Agendas
Brief: Daily
Brief: Weekly
```

**Filter Creation** (runs daily 6 AM):

For each client with `setup_complete = TRUE`:

1. **Incoming Client Emails**: `from:{contact@email.com OR contact2@email.com}` → `Client: [Name]`
2. **Outgoing Client Emails**: `to:{contact@email.com OR contact2@email.com}` → `Client: [Name]`
3. **Meeting Summaries**: `from:me subject:"Meeting Summary: [Name]" to:{contacts}` → `Client: [Name]/Meeting Summaries`
4. **Agendas**: `from:me to:me subject:"Agenda: [Name]"` → `Client: [Name]/Meeting Agendas`

**Safety Features**:
- Only manages system-created labels (starts with `Client:` or `Brief:`)
- Checks for duplicate filters before creating
- Detects orphaned labels (doesn't auto-delete)

**Key Functions**:
- `syncLabelsAndFilters()` - Daily sync
- `syncClientLabels(client)` - Per-client setup
- `createGmailApiFilter(criteria, labelName)` - Gmail API integration
- `removeOrphanedLabels(clients)` - Cleanup detection

### **AgendaGeneration.gs** (AI-Powered Meeting Prep)

**Purpose**: Generate context-aware agendas for upcoming meetings

**Trigger**: Hourly, 8 AM - 6 PM only

**Context Gathering**:
- **Todoist Tasks**: Due today or overdue from client project
- **Recent Emails**: Last 7 days, max 20 threads
- **Previous Notes**: LAST meeting notes only (not full history)
- **Unmatched Actions**: Action items from previous meeting not in Todoist (fuzzy matching)

**AI Generation** (Claude 3.5):
- Prompt: `AGENDA_CLAUDE_PROMPT` from Prompts sheet
- Max tokens: 1000
- Model: Dynamic (Haiku or Sonnet based on config)
- Output: Structured HTML agenda

**Email & Doc Updates**:
- Email to self: Subject `Agenda: {client_name} - {meeting_title} ({date})`
- Label applied immediately from Client Registry
- Google Doc: HTML converted to plain text with structured delimiters

**Tracking**:
- `Generated_Agendas` sheet prevents duplicates
- Calendar doc attachment (if Advanced Calendar API enabled)

**Key Functions**:
- `generateAgendas()` - Main trigger handler
- `gatherAgendaContext()` - Context collection
- `generateAgendaWithClaude()` - AI call
- `sendAgendaEmail()` - Email with labeling
- `appendAgendaToDoc()` - Doc update with HTML→text conversion
- `htmlToPlainText()` - HTML stripping for readable docs

### **OutlookReports.gs** (Strategic Briefings)

**Purpose**: AI-generated daily and weekly outlook reports

**Daily Outlook** (7 AM):
- Today's meetings by client
- Todoist tasks due today
- Schedule conflicts
- Missing agendas
- Unread emails (optional)

**Weekly Outlook** (Monday 7 AM):
- Week's meetings grouped by day and client
- Weekly task dashboard
- Multi-day scheduling analysis
- Overdue prerequisite detection

**AI Generation**:
- Prompts: `DAILY_BRIEFING_CLAUDE_PROMPT`, `WEEKLY_BRIEFING_CLAUDE_PROMPT`
- Max tokens: 4000
- Fallback to HTML templates if API fails

**Email Format**:
- Subject: `Daily Outlook - {date}` or `Weekly Outlook - Week of {date}`
- Recipient: Self
- Labels: `Brief: Daily` or `Brief: Weekly`

**Key Functions**:
- `generateDailyOutlook()` / `generateWeeklyOutlook()` - Main handlers
- `compileDailyData()` / `compileWeeklyData()` - Data gathering
- `generateDailyOutlookWithClaude()` - AI generation
- `detectScheduleConflicts()` - Overlap detection
- `sendOutlookEmail()` - Email with auto-labeling

### **ClientOnboarding.gs** (Automatic Resource Creation)

**Purpose**: Automated setup for new clients

**Workflow** (runs daily 6:30 AM):
1. Scan Client_Registry for clients without `setup_complete = TRUE`
2. For each client:
   - Create Google Doc if `google_doc_url` empty
   - Create Todoist project if `todoist_project_id` empty
   - Update Client_Registry with URLs/IDs
   - Mark `setup_complete = TRUE` only if BOTH succeed

**Google Doc Creation**:
- Template: `DOC_NAME_TEMPLATE` = "Client Notes - {client_name}"
- Initial content from `DOC_TEMPLATE` (customizable)
- Moves to specified folder if `docs_folder_path` provided

**Todoist Project Creation**:
- Uses Todoist REST API
- Project name = client name
- Returns project ID for storage

**Key Functions**:
- `processNewClients()` - Main automation
- `createClientDoc()` - Doc creation with template
- `createTodoistProject()` - API call
- `validateClientResources()` - Status check

### **FolderSync.gs** (Drive Folder Management)

**Purpose**: Keep Client_Registry folder dropdown in sync with Drive

**Workflow** (runs daily 5:30 AM):
1. Scan "My Drive" recursively (max 10 levels)
2. Scan shared folders
3. Store paths in hidden "Folders" sheet
4. Update data validation for `docs_folder_path` column

**Key Functions**:
- `syncDriveFolders()` - Main sync
- `collectFolders()` - Recursive scanner
- `updateClientRegistryFolderValidation()` - Dropdown setup
- `getFolderIdByPath()` / `getFolderPathById()` - Lookup helpers

### **PromptManager.gs** (AI Prompt Customization)

**Purpose**: Manage AI prompts and email templates

**Stored Prompts** (hidden "Prompts" sheet):
- `WEEKLY_BRIEFING_CLAUDE_PROMPT` - Strategic weekly overview
- `DAILY_BRIEFING_CLAUDE_PROMPT` - Daily focus briefing
- `AGENDA_CLAUDE_PROMPT` - Meeting agenda generation
- `AGENDA_EMAIL_TEMPLATE` - Email wrapper for agendas
- Legacy templates for non-AI fallback

**Features**:
- Variable substitution: `{variable_name}` replaced at runtime
- Model selection per prompt (Haiku vs Sonnet)
- Prompt compression with AI (optimize token usage)
- Web UI for editing via `showPromptsEditor()`

**Claude Model Management**:
- Fetches available models from Anthropic API
- 24-hour cache in Script Properties
- Fallback to hardcoded models if API fails
- Pretty model name formatting

**Key Functions**:
- `getPrompt(key)` / `setPrompt(key, value, model)` - CRUD operations
- `getAllPromptsForEditor()` - For UI
- `fetchAvailableModelsFromAPI()` - Model discovery
- `compressPromptWithAI()` - Token optimization
- `applyTemplate(template, variables)` - Variable substitution

### **Utilities.gs** (Shared Helper Functions)

**Purpose**: Common utilities used across all modules

**Categories**:

1. **Date/Time**: `formatDate()`, `formatDateTime()`, `getStartOfToday()`, etc.
2. **Logging**: `logProcessing()` - Main audit trail function
3. **String Utils**: `truncate()`, `escapeHtml()`, `stripHtml()`, `toTitleCase()`
4. **Array Utils**: `uniqueArray()`, `groupBy()`, `sortBy()`
5. **Error Handling**: `withErrorHandling()`, `retryWithBackoff()`
6. **Spreadsheet**: `getSheetData()`, `appendToSheet()`, `findRowByColumn()`
7. **Todoist**: `fetchTodoistTasksDueToday()`, `createTodoistTask()`, `createTodoistTasksWithAssignees()`
8. **Gmail**: `markMessageProcessed()`, `isMessageProcessed()`
9. **Documents**: `appendMeetingNotesToDoc()`, `appendAgendaToDoc()`, `extractDocIdFromUrl()`
10. **Markdown**: `markdownToHtml()` - Fathom summary conversion
11. **AI**: `extractActionItemsWithAI()` - Claude-powered action extraction

---

## Setup Instructions

### Prerequisites

- Google Workspace account
- Fathom account (for meeting recordings)
- Todoist account
- Claude API key (Anthropic)

### Step 1: Create Google Sheets

1. Create a new Google Sheets file: "BDC Automation Data"
2. Create the following sheets (tabs):
   - `Client_Registry` - Client master data
   - `Generated_Agendas` - Agenda tracking
   - `Processing_Log` - Audit trail
   - `Unmatched` - Unidentified items
   - `Folders` - Drive folder cache (hidden)
   - `Prompts` - AI prompts (hidden)

### Step 2: Set Up Apps Script

1. In your Google Sheets, go to **Extensions → Apps Script**
2. Copy all `.gs` files from this repo to the script editor:
   - Code.gs
   - ClientIdentification.gs
   - MeetingAutomation.gs
   - EmailSorting.gs
   - AgendaGeneration.gs
   - OutlookReports.gs
   - ClientOnboarding.gs
   - FolderSync.gs
   - PromptManager.gs
   - Utilities.gs
3. Copy HTML files:
   - SettingsEditor.html
   - PromptsEditor.html
   - MigrationWizard.html

> Setup note: `SETUP_RUN_THIS_FIRST()` now detects existing sheets and will preserve them, creating only missing sheets after a confirmation prompt.

### Step 3: Enable Advanced Services

In Apps Script editor:
1. Click **Project Settings** (gear icon)
2. Scroll to **Google Services**
3. Enable **Gmail API** (required for filters)
4. Enable **Google Calendar API** (optional, for doc attachments)

### Step 4: Configure Script Properties

1. In Apps Script editor: **Project Settings → Script Properties**
2. Add the following properties:

| Property | Value | Example |
|----------|-------|---------|
| `SPREADSHEET_ID` | Your Sheets file ID | `1ABC...xyz` |
| `TODOIST_API_TOKEN` | Your Todoist API token | From Todoist Settings → Integrations |
| `CLAUDE_API_KEY` | Your Anthropic API key | `sk-ant-...` |
| `FATHOM_WEBHOOK_SECRET` | Fathom webhook secret (optional) | For signature verification |
| `USER_NAME` | Your name | For email signatures |
| `DOC_NAME_TEMPLATE` | Doc naming template | `Client Notes - {client_name}` |
| `MEETING_SUBJECT_TEMPLATE` | Email subject template | `Meeting Summary: {client_name} - {meeting_title} ({date})` |
| `NWS_WEATHER_ENABLED` | Enable NWS forecasts | `true` |
| `NWS_WEATHER_STATE` | Primary weather state (US) | `WA` |
| `NWS_WEATHER_CITY` | Primary weather city | `Seattle` |
| `NWS_WEATHER_LAT` | Primary weather latitude | `47.6062` |
| `NWS_WEATHER_LON` | Primary weather longitude | `-122.3321` |
| `WEATHER_API_KEY` | OpenWeather API key (optional) | `your-key` |

> Tip: You can also run `SETUP_RUN_THIS_FIRST()` from Apps Script to answer prompts (including National Weather Service setup) and then fine-tune the location in the Settings editor.

### Step 5: Deploy as Web App

1. In Apps Script editor: **Deploy → New deployment**
2. Type: **Web app**
3. Description: "BDC Automation Webhook"
4. Execute as: **Me**
5. Who has access: **Anyone**
6. Click **Deploy**
7. Copy the **Web App URL** (needed for Fathom)

### Step 6: Create Triggers

1. In Apps Script editor: **Triggers** (clock icon)
2. Option A: Run `setupTriggers()` manually from Code.gs
3. Option B: Create manually:
   - `runFolderSync` - Time-based, Day timer, 5:00-6:00 AM
   - `runLabelAndFilterCreation` - Time-based, Day timer, 6:00-7:00 AM
   - `runClientOnboarding` - Time-based, Day timer, 6:00-7:00 AM
   - `runSentMeetingSummaryMonitor` - Time-based, Minutes timer, Every 10 minutes
   - `runAgendaGeneration` - Time-based, Hour timer, Every hour
   - `runDailyOutlook` - Time-based, Day timer, 7:00-8:00 AM
   - `runWeeklyOutlook` - Time-based, Week timer, Monday, 7:00-8:00 AM

### Step 7: Configure Fathom

1. Go to Fathom Settings → Integrations → Webhooks
2. Add new webhook:
   - URL: Your Web App URL from Step 5
   - Events: `recording.ended`
   - Secret: (optional, for signature verification)

### Step 8: Populate Client Registry

Add your first client to `Client_Registry`:

| Column | Example | Required |
|--------|---------|----------|
| `client_name` | ACME Corp | ✅ |
| `contact_emails` | john@acme.com, sarah@acme.com | ✅ |
| `docs_folder_path` | Clients/ACME Corp | Optional |
| `setup_complete` | FALSE | Auto-set |
| `google_doc_url` | (auto-created) | Auto-set |
| `todoist_project_id` | (auto-created) | Auto-set |
| `gmail_label` | Client: ACME Corp | Auto-set |
| `meeting_summaries_label` | Client: ACME Corp/Meeting Summaries | Auto-set |
| `meeting_agendas_label` | Client: ACME Corp/Meeting Agendas | Auto-set |

### Step 9: Test the System

1. **Test Folder Sync**: Run `runFolderSync()` manually → Check `Folders` sheet populated
2. **Test Client Onboarding**: Wait for 6:30 AM trigger or run `runClientOnboarding()` → Check `google_doc_url` and `todoist_project_id` populated
3. **Test Labels/Filters**: Wait for 6 AM trigger or run `runLabelAndFilterCreation()` → Check Gmail labels created
4. **Test Fathom Webhook**: Record a test meeting with Fathom → Check Gmail draft created
5. **Test Agenda Generation**: Add calendar event → Wait for hourly trigger or run `runAgendaGeneration()` → Check for agenda email

### Step 10: Monitor Logs

- **Processing_Log** sheet: All automated actions
- **Unmatched** sheet: Items that couldn't be matched to clients
- Apps Script **Executions** page: Trigger history and errors

---

## Configuration

### Script Properties Reference

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `SPREADSHEET_ID` | String | Required | Google Sheets file ID |
| `TODOIST_API_TOKEN` | String | Required | Todoist REST API token |
| `CLAUDE_API_KEY` | String | Required | Anthropic API key |
| `FATHOM_WEBHOOK_SECRET` | String | Optional | HMAC verification secret |
| `USER_NAME` | String | Required | Your name for signatures |
| `CLIENT_DOCS_FOLDER_ID` | String | Optional | Default folder for client docs |
| `DOC_NAME_TEMPLATE` | String | `Client Notes - {client_name}` | Doc naming pattern |
| `DOC_TEMPLATE` | String | (see Prompts sheet) | Initial doc content |
| `MEETING_SUBJECT_TEMPLATE` | String | `Meeting Summary: {client_name} - {meeting_title} ({date})` | Email subject |
| `MEETING_SIGNATURE` | String | `Best regards,\n{user_name}` | Email signature |
| `AGENDA_SUBJECT_TEMPLATE` | String | `Agenda: {client_name} - {meeting_title} ({date})` | Agenda subject |
| `DAILY_BRIEFING_LABEL` | String | `Brief: Daily` | Gmail label for daily outlooks |
| `WEEKLY_BRIEFING_LABEL` | String | `Brief: Weekly` | Gmail label for weekly outlooks |
| `INCLUDE_UNREAD_EMAILS` | Boolean | `false` | Include unread count in outlooks |
| `AUTO_MARK_READ_AFTER_DAYS` | Number | `0` | Auto-mark old emails read (0=disabled) |
| `BUSINESS_HOURS_START` | Number | `8` | Agenda generation start hour |
| `BUSINESS_HOURS_END` | Number | `18` | Agenda generation end hour |
| `NWS_WEATHER_ENABLED` | Boolean | `false` | Enable National Weather Service forecasts |
| `NWS_WEATHER_STATE` | String | (blank) | Primary weather state (US) |
| `NWS_WEATHER_CITY` | String | (blank) | Primary weather city |
| `NWS_WEATHER_LAT` | Number | (blank) | Primary weather latitude |
| `NWS_WEATHER_LON` | Number | (blank) | Primary weather longitude |
| `WEATHER_API_KEY` | String | Optional | OpenWeather fallback API key |
| `TRAVEL_WEATHER_ENABLED` | Boolean | `false` | Reserve secondary weather (travel) support |

### Client Registry Columns

| Column | Type | Description | Auto-Populated |
|--------|------|-------------|----------------|
| `client_name` | String | Client display name | Manual |
| `contact_emails` | String | Comma-separated email addresses | Manual |
| `docs_folder_path` | String | Google Drive folder path | Manual |
| `setup_complete` | Boolean | Whether resources are created | Auto |
| `google_doc_url` | String | Running meeting notes doc | Auto |
| `todoist_project_id` | String | Todoist project ID | Auto |
| `gmail_label` | String | Base Gmail label name | Auto |
| `meeting_summaries_label` | String | Sub-label for summaries | Auto |
| `meeting_agendas_label` | String | Sub-label for agendas | Auto |

### AI Prompt Customization

Access via spreadsheet menu: **Adjust Prompts**

Edit prompts in `Prompts` sheet (hidden) or use the web UI:
- `WEEKLY_BRIEFING_CLAUDE_PROMPT` - Weekly strategic overview
- `DAILY_BRIEFING_CLAUDE_PROMPT` - Daily focus briefing
- `AGENDA_CLAUDE_PROMPT` - Meeting agenda generation
- `AGENDA_EMAIL_TEMPLATE` - Email wrapper for agendas

**Model Selection**: Choose Haiku (fast/cheap) or Sonnet (better quality) per prompt

---

## User Workflows

### Adding a New Client

1. Open `Client_Registry` sheet
2. Add new row:
   - `client_name`: Client display name
   - `contact_emails`: All client contact emails (comma-separated)
   - `docs_folder_path`: (optional) Folder path in Drive
3. Save the sheet
4. System automatically (within 24 hours):
   - Creates Google Doc
   - Creates Todoist project
   - Creates Gmail labels
   - Sets up Gmail filters
   - Marks `setup_complete = TRUE`

### Processing a Meeting

1. **Record meeting** in Fathom (with client participant)
2. **Meeting ends** → Fathom sends webhook
3. **Check Gmail** → Draft created with summary
4. **Review draft** → Edit if needed
5. **Send to client** → Manually send the draft
6. **Wait 10 minutes** → System detects sent email
7. **Automatic actions**:
   - Action items extracted (Claude AI)
   - Todoist tasks created
   - Notes appended to client doc
8. **Verify** → Check Processing_Log for success

### Reviewing Daily Outlook

1. **7 AM**: Check email for "Daily Outlook - [date]"
2. Review sections:
   - Today's Focus (AI-generated priority)
   - Schedule Overview (meetings by client)
   - Priority Tasks (due today)
   - Alerts (overdue, conflicts, missing agendas)
   - Communication Queue (action items)
3. Click calendar links to join meetings
4. Check if agendas are ready (automatically generated)

### Managing Unmatched Items

1. Open `Unmatched` sheet
2. Review items that couldn't be matched
3. For each item:
   - Check if client exists in `Client_Registry`
   - If not, add client with matching contact emails
   - If client exists, add missing contact email
   - Mark `manually_resolved = TRUE` once fixed
4. Re-run meeting processing if needed

### Customizing AI Prompts

1. **Spreadsheet menu**: Adjust Prompts → Edit Prompts
2. Select prompt to edit
3. Modify prompt text (preserve variable placeholders like `{client_name}`)
4. Select model: Haiku (fast) or Sonnet (quality)
5. Optional: Use "Compress with AI" to optimize token usage
6. Save changes → Takes effect immediately

---

## Data Structure

### Client_Registry Sheet

**Purpose**: Master client database

| Column | Example | Notes |
|--------|---------|-------|
| client_name | ACME Corp | Display name |
| contact_emails | john@acme.com, sarah@acme.com | Comma-separated |
| docs_folder_path | Clients/ACME Corp | Drive folder path |
| setup_complete | TRUE | Auto-populated |
| google_doc_url | https://docs.google.com/... | Auto-created |
| todoist_project_id | 2301234567 | Auto-created |
| gmail_label | Client: ACME Corp | Auto-created |
| meeting_summaries_label | Client: ACME Corp/Meeting Summaries | Auto-created |
| meeting_agendas_label | Client: ACME Corp/Meeting Agendas | Auto-created |

### Generated_Agendas Sheet

**Purpose**: Prevent duplicate agenda generation

| Column | Example | Notes |
|--------|---------|-------|
| event_id | 7adk82j... | Calendar event ID |
| event_title | Project Kickoff | Event name |
| client_name | ACME Corp | Matched client |
| generated_timestamp | 2025-01-19 10:23:45 | When agenda created |

### Processing_Log Sheet

**Purpose**: Audit trail of all automated actions

| Column | Example | Notes |
|--------|---------|-------|
| timestamp | 2025-01-19 10:23:45 | When action occurred |
| action_type | AGENDA_GENERATED | Action category |
| client_name | ACME Corp | Related client |
| details | Generated agenda for: Project Kickoff | Description |
| status | success | success/error/warning/info |

**Action Types**:
- `WEBHOOK_RECEIVED` - Fathom webhook
- `DRAFT_CREATED` - Meeting summary draft
- `SUMMARY_SENT` - Detected sent summary
- `TASKS_CREATED` - Todoist tasks
- `DOC_UPDATED` - Google Doc appended
- `AGENDA_GENERATED` - Agenda created
- `LABEL_SYNC` - Gmail labels synced
- `CLIENT_ONBOARDED` - New client setup
- `DAILY_OUTLOOK` - Daily briefing
- `WEEKLY_OUTLOOK` - Weekly briefing

### Unmatched Sheet

**Purpose**: Track items without client match

| Column | Example | Notes |
|--------|---------|-------|
| timestamp | 2025-01-19 10:23:45 | When logged |
| item_type | meeting | "meeting" or "email" |
| item_details | Calendar Event: Team Sync at 2025-01-19 2:00 PM | Description |
| participant_emails | unknown@company.com | Emails found |
| manually_resolved | FALSE | User marks TRUE after fixing |

### Folders Sheet (Hidden)

**Purpose**: Cache of Drive folder structure

| Column | Example | Notes |
|--------|---------|-------|
| folder_path | Clients/ACME Corp | Full path |
| folder_id | 1ABC...xyz | Drive folder ID |
| folder_url | https://drive.google.com/... | Direct link |

### Prompts Sheet (Hidden)

**Purpose**: AI prompt storage

| Column | Example | Notes |
|--------|---------|-------|
| prompt_key | AGENDA_CLAUDE_PROMPT | Unique identifier |
| prompt_value | You are an executive assistant... | Full prompt text |
| model_preference | sonnet | "haiku" or "sonnet" |

---

## Automation Schedule

| Time | Trigger | Function | Purpose |
|------|---------|----------|---------|
| **5:30 AM** | Daily | `runFolderSync()` | Update Drive folder cache |
| **6:00 AM** | Daily | `runLabelAndFilterCreation()` | Sync Gmail labels/filters |
| **6:30 AM** | Daily | `runClientOnboarding()` | Create resources for new clients |
| **7:00 AM** | Daily | `runDailyOutlook()` | Send daily briefing |
| **7:00 AM** | Monday | `runWeeklyOutlook()` | Send weekly briefing |
| **Every 10 min** | All day | `runSentMeetingSummaryMonitor()` | Detect sent summaries |
| **Every hour** | 8 AM-6 PM | `runAgendaGeneration()` | Generate meeting agendas |
| **On demand** | Webhook | `doPost(e)` | Receive Fathom webhooks |

**Business Hours Enforcement**:
- Agenda generation only runs 8 AM - 6 PM (configurable)
- Other triggers run all day

**Execution Limits**:
- Apps Script: 6 hours/day (consumer), 90 hours/day (Workspace)
- Trigger frequency: Max 1 trigger per minute per function
- Email quota: 100/day (consumer), 1500/day (Workspace)

---

## External Integrations

### Fathom (Meeting Recordings)

**Webhook Configuration**:
- URL: Your Apps Script Web App URL
- Event: `recording.ended`
- Payload includes: title, date, participants, summary, action items

**Security**:
- Optional HMAC SHA-256 signature verification
- Set `FATHOM_WEBHOOK_SECRET` in Script Properties

**Payload Structure**:
```json
{
  "meeting_title": "Project Kickoff",
  "meeting_date": "2025-01-19T14:00:00Z",
  "participants": [
    {"name": "John Doe", "email": "john@acme.com"}
  ],
  "summary": "Discussed project goals...",
  "action_items": [
    {
      "description": "Send proposal",
      "assignee": "John Doe",
      "due_date": "2025-01-26"
    }
  ]
}
```

### Todoist (Task Management)

**API**: REST API v2
**Authentication**: Bearer token in `Authorization` header

**Endpoints Used**:
- `POST /rest/v2/projects` - Create client project
- `GET /rest/v2/projects` - List all projects
- `POST /rest/v2/tasks` - Create task
- `GET /rest/v2/tasks` - Query tasks by project

**Task Creation**:
- Content: Action item description
- Project ID: From `Client_Registry.todoist_project_id`
- Due date: From Fathom or default
- Assignee: Extracted from action item

### Claude API (Anthropic)

**API**: Claude Messages API v1
**Authentication**: `x-api-key` header

**Models Used**:
- **Claude 3.5 Sonnet** (primary): Better quality, higher cost
- **Claude 3 Haiku** (fallback): Faster, lower cost

**Endpoints**:
- `POST /v1/messages` - Generate completions
- `GET /v1/models` - List available models (cached 24h)

**Use Cases**:
1. **Agenda Generation**: Context-aware meeting preparation
2. **Action Extraction**: Intelligent parsing of edited drafts
3. **Daily Briefing**: Strategic morning overview
4. **Weekly Briefing**: Multi-client project dashboard

**Token Usage** (approximate):
- Agenda: 500-1000 tokens
- Daily outlook: 2000-4000 tokens
- Weekly outlook: 3000-5000 tokens
- Action extraction: 300-500 tokens

### National Weather Service (US) + OpenWeather (Fallback)

**NWS API**: Public NOAA/NWS endpoints (no key required)  
**OpenWeather API**: Optional fallback when NWS is disabled or unavailable

**Settings**:
- `NWS_WEATHER_ENABLED`: Toggle NWS usage on/off
- `NWS_WEATHER_STATE`, `NWS_WEATHER_CITY`, `NWS_WEATHER_LAT`, `NWS_WEATHER_LON`: Primary location selection (via Settings editor dropdowns)
- `WEATHER_API_KEY`: OpenWeather API key

**Behavior**:
- Daily outlooks prioritize NWS when enabled and a location is selected.
- OpenWeather is used only if NWS is disabled or missing location data.

### Gmail (Email & Organization)

**APIs Used**:
- **GmailApp** (native): Draft creation, email sending
- **Gmail Advanced Service** (API): Label management, filter creation

**Label Structure**:
```
Client: [ClientName]
├── Meeting Summaries
└── Meeting Agendas
Brief: Daily
Brief: Weekly
```

**Filter Behavior**:
- Created via Gmail API (permanent filters)
- Applied automatically by Gmail (no Apps Script execution needed)
- System only manages its own labels (safety check)

### Google Calendar (Events & Scheduling)

**API**: CalendarApp (native)

**Operations**:
- Fetch events for date range
- Get event attendees (including organizer)
- Attach Google Docs to events (requires Advanced Calendar API)

**Conflict Detection**:
- Compares event start/end times
- Identifies overlapping meetings
- Reports conflicts in daily/weekly outlooks

### Google Docs (Meeting Notes)

**API**: DocumentApp (native)

**Operations**:
- Create new docs from template
- Append paragraphs with formatting
- Apply heading styles
- Extract text for parsing

**Document Structure**:
```
═══════════════════════════════════════
MEETING NOTES - 2025-01-19
───────────────────────────────────────
[Meeting summary and action items]
───────────────────────────────────────
END OF MEETING NOTES - 2025-01-19
═══════════════════════════════════════

═══════════════════════════════════════
MEETING AGENDA - 2025-01-26
───────────────────────────────────────
[Agenda content]
───────────────────────────────────────
END OF MEETING AGENDA - 2025-01-26
═══════════════════════════════════════
```

**Delimiters**: Used for parsing previous notes and extracting action items

### Google Drive (Folder Management)

**API**: DriveApp (native)

**Operations**:
- Scan folder structure (recursive)
- Create folders programmatically
- Move docs to folders
- Share docs with users

**Folder Sync**:
- Daily scan of "My Drive" (max 10 levels deep)
- Populates `Folders` sheet for dropdown validation
- Updates Client_Registry data validation

---

## Troubleshooting

### Common Issues

#### 1. Fathom Webhook Not Working

**Symptoms**: No Gmail draft created after meeting

**Checks**:
1. Verify Web App URL in Fathom settings
2. Check webhook secret matches Script Properties
3. Review Apps Script **Executions** page for errors
4. Check if meeting had client participants
5. Verify `doPost(e)` function exists

**Fix**:
- Re-deploy Web App (Deploy → Manage deployments → Edit → Deploy)
- Check `Processing_Log` for `WEBHOOK_RECEIVED` entries
- Look in `Unmatched` sheet for unidentified meetings

#### 2. No Gmail Labels Created

**Symptoms**: Client emails not labeled

**Checks**:
1. Verify Gmail Advanced Service enabled (Project Settings → Google Services)
2. Check `setup_complete = TRUE` in Client_Registry
3. Run `runLabelAndFilterCreation()` manually
4. Check Apps Script permissions granted

**Fix**:
- Enable Gmail Advanced Service
- Wait for 6 AM daily sync or run manually
- Check `Processing_Log` for `LABEL_SYNC` entries
- Verify contact emails populated

#### 3. Agendas Not Generated

**Symptoms**: No agenda email before meeting

**Checks**:
1. Verify time is within business hours (8 AM - 6 PM)
2. Check `Generated_Agendas` sheet (already exists?)
3. Verify client matched to calendar event
4. Check Claude API key in Script Properties
5. Review `Processing_Log` for `AGENDA_ERROR` entries

**Fix**:
- Add client contact emails to Client_Registry
- Check calendar event has client as attendee
- Verify Claude API quota not exceeded
- Run `runAgendaGeneration()` manually for testing

#### 4. Todoist Tasks Not Created

**Symptoms**: No tasks after sending meeting summary

**Checks**:
1. Verify Todoist API token in Script Properties
2. Check `todoist_project_id` populated in Client_Registry
3. Verify sent email monitored (check label applied)
4. Review `Processing_Log` for `TASKS_CREATED` entries

**Fix**:
- Test Todoist API token: `getTodoistProjects()`
- Run `createTodoistProject(clientName)` if missing
- Check if meeting summary label exists
- Review action item extraction (Claude AI)

#### 5. Meeting Notes Not Appended to Doc

**Symptoms**: Google Doc not updated after sending summary

**Checks**:
1. Verify `google_doc_url` populated in Client_Registry
2. Check doc permissions (script has edit access)
3. Review `Processing_Log` for `DOC_UPDATED` entries
4. Verify sent email was detected

**Fix**:
- Re-run client onboarding if doc missing
- Share doc with script's service account
- Check `appendMeetingNotesToDoc()` for errors

#### 6. Agenda Email Body Empty

**Symptoms**: Email shows only "```html"

**Fix Applied**: System now strips markdown code fences from Claude responses

**If still occurring**:
- Check `AGENDA_EMAIL_TEMPLATE` in Prompts sheet
- Verify Claude API returns valid HTML
- Review `Processing_Log` for `AGENDA_GENERATED` details

#### 7. Raw HTML in Google Docs

**Symptoms**: Doc shows HTML tags instead of formatted text

**Fix Applied**: System now converts HTML to plain text before appending

**If still occurring**:
- Verify `htmlToPlainText()` function exists
- Check if old agendas (pre-fix) need manual cleanup
- Review `appendAgendaToDoc()` code

### Error Logs

**Check Processing_Log Sheet**:
- Filter by `status = error`
- Review `details` column for error messages
- Match `timestamp` to Apps Script **Executions** page

**Common Error Types**:
- `WEBHOOK_ERROR` - Fathom webhook processing failed
- `AGENDA_ERROR` - Agenda generation failed
- `TODOIST_ERROR` - Task creation failed
- `DOC_ERROR` - Document update failed
- `LABEL_ERROR` - Label/filter sync failed

**Apps Script Logs**:
1. Apps Script editor → **Executions**
2. Filter by function name and status
3. Click execution to see full stack trace
4. Check quota usage (6h/day limit)

### Performance Optimization

**Reduce Execution Time**:
- Limit `fetchRecentClientEmails()` to 7 days
- Use `fetchPreviousMeetingNotes()` LAST section only (not full doc)
- Enable Claude model caching (24h cache)
- Batch Todoist task creation

**Reduce API Costs**:
- Use Haiku model for agendas (cheaper than Sonnet)
- Compress prompts with `compressPromptWithAI()`
- Reduce agenda generation frequency (every 2 hours instead of hourly)

**Reduce Email Quota**:
- Disable `INCLUDE_UNREAD_EMAILS` in outlooks
- Limit unread email fetching to 3 days
- Combine multiple agendas into one email (future)

---

## Advanced Features

### Custom Label Names

**Default**: `Client: [Name]`, `Client: [Name]/Meeting Summaries`, `Client: [Name]/Meeting Agendas`

**Customization**:
1. Edit `Client_Registry` directly:
   - `gmail_label` - Base label
   - `meeting_summaries_label` - Summaries sub-label
   - `meeting_agendas_label` - Agendas sub-label
2. Labels will be used as-is (no auto-generation)
3. Filters will be created with custom labels

**Use Case**: Match existing label structure or use custom naming conventions

### AI Model Selection

**Per-Prompt Model Choice**:
- Edit `Prompts` sheet → `model_preference` column
- Options: `haiku` (fast/cheap), `sonnet` (quality)
- Takes effect immediately

**When to Use Haiku**:
- High-frequency operations (agendas every hour)
- Budget constraints
- Simple content generation

**When to Use Sonnet**:
- Strategic reports (daily/weekly outlooks)
- Complex reasoning (action item extraction)
- High-quality prose

### Webhook Signature Verification

**Enable**:
1. Set `FATHOM_WEBHOOK_SECRET` in Script Properties
2. Fathom will send `X-Fathom-Signature` header
3. System verifies HMAC SHA-256 signature

**Security Benefits**:
- Prevents webhook spoofing
- Ensures requests from Fathom only
- Protects against replay attacks

### Internal-Only Clients

**Use Case**: Meetings with only internal team members

**Setup**:
1. Add client to Client_Registry:
   - `client_name`: "Internal"
   - `contact_emails`: Your team's emails (comma-separated)
2. System requires **ALL** participants to match internal emails
3. If any external participant, won't match "Internal"

**Behavior**: Special matching logic in `identifyClient()`

### Prompt Compression

**Feature**: Optimize prompts for token efficiency

**Usage**:
1. Spreadsheet menu: Adjust Prompts → Edit Prompts
2. Select prompt
3. Click "Compress with AI"
4. Review compressed version
5. Save if token savings acceptable

**Algorithm**: Uses Claude Sonnet to rewrite prompt conservatively

### Custom Email Signatures

**Default**: `Best regards,\n{user_name}`

**Customization**:
- Edit `MEETING_SIGNATURE` in Script Properties
- Variables: `{user_name}`, `{date}`, `{client_name}`
- Supports HTML formatting

### Auto-Mark Old Emails Read

**Feature**: Automatically mark old unread emails as read

**Enable**:
- Set `UNREAD_AUTO_MARK_DAYS` in Script Properties
- Example: `7` = emails older than 7 days auto-marked read
- Set to `0` to disable

**Use Case**: Inbox zero automation

### Folder-Based Doc Organization

**Feature**: Organize client docs in Drive folders

**Setup**:
1. Create folder structure in Drive (e.g., "Clients/[ClientName]")
2. Run `syncDriveFolders()` to populate Folders sheet
3. Set `docs_folder_path` in Client_Registry
4. New docs automatically created in specified folder

**Benefits**:
- Better organization
- Easier sharing with team
- Matches existing folder structure

### Migration Wizard

**Feature**: Bulk update labels for existing clients

**Usage**:
1. Spreadsheet menu: Adjust Prompts → Migration Wizard
2. Select clients to update
3. Choose new label structure
4. Option to delete old labels
5. System updates Client_Registry + Gmail

**Use Case**: Restructure labels for all clients at once

---

## License

Copyright © 2025 Blue Dot Consulting. All rights reserved.

---

## Support

For issues or questions:
1. Check `Processing_Log` and `Unmatched` sheets
2. Review Apps Script **Executions** page
3. Consult this README's Troubleshooting section
4. Contact your system administrator

---

## Version History

- **v1.0** (2025-01-19) - Initial release with full automation suite
  - Meeting automation (Fathom webhooks)
  - Email organization (Gmail labels/filters)
  - Agenda generation (Claude AI)
  - Daily/weekly outlooks (strategic briefings)
  - Client onboarding (automatic resource creation)
  - Drive folder sync
  - AI prompt management
  - Comprehensive logging and error handling
