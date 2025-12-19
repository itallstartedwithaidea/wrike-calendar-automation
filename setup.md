# Setup Guide

## Step-by-Step Installation

### 1. Create Apps Script Project

1. Navigate to [script.google.com](https://script.google.com)
2. Click "New Project"
3. Name it "Wrike Calendar Automation"

### 2. Install the Code

1. Delete default `myFunction()`
2. Copy entire contents of `src/Code.gs`
3. Paste into Apps Script editor
4. Save (Ctrl+S or Cmd+S)

### 3. Configure Settings

Update the `CONFIG` object:
```javascript
const CONFIG = {
  assigneeName: "Your Name",  // Change this
  timezone: "America/Phoenix", // Your timezone
  alerts: {
    criticalClientsNeedEarlier: ["CLIENT1", "CLIENT2"], // Your VIP clients
  }
};
```

### 4. Run Initial Setup

1. Select `setupAllTriggers` from function dropdown
2. Click "Run" ▶️
3. Grant permissions when prompted:
   - Read, compose, send emails
   - Manage calendars
   - Create/edit labels

### 5. Verify Installation

Run test functions:
```javascript
testDateExtraction()  // Should extract "Jan 5, 2026"
debugLastEmail()      // Check last Wrike email
```

### 6. Gmail Labels

The script auto-creates:
- `Wrike/Processed`
- `Wrike/PastDue`
- `Wrike/Urgent`

## Triggers Schedule

- **Email Processing**: Every hour
- **Daily Summary**: 8:00 PM daily
- **Past Due Check**: Hourly with email processing

## Customization Options

### Lead Times by Task Type
```javascript
taskPatterns: {
  "Final Approval": 72,    // 3 days
  "QA Review": 48,        // 2 days
  "Upload": 24,           // 1 day
}
```

### Color Schemes
```javascript
colors: {
  normal: "5",   // Yellow
  late: "11",    // Red
  urgent: "6",   // Orange
}
```

## First Run Expectations

1. All unread Wrike emails will be processed
2. Calendar events created for future tasks
3. Past due tasks consolidated into one event
4. Emails marked as read and labeled

## Monitoring

Check logs: View → Execution log

Look for:
- "Found X unread Wrike emails"
- "Created calendar event:"
- "Processed X tasks (Y past due)"
