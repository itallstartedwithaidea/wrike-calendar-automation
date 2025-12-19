# Wrike to Google Calendar Automation

Automatically sync Wrike task assignments to Google Calendar with intelligent scheduling, past due tracking, and daily summaries.

## 🌟 Features

- **Automatic Calendar Events**: Creates calendar reminders 48-72 hours before due dates
- **Past Due Consolidation**: Groups all overdue tasks into a single daily event
- **Smart Date Detection**: Accurately identifies due dates vs email dates
- **Daily Summary Emails**: Sends comprehensive task summary at 8 PM daily
- **No Duplicates**: Prevents duplicate calendar entries
- **Color-Coded Priority**: Yellow (normal), Red (past due), Orange (urgent), Green (complete)
- **Weekend/Holiday Awareness**: Automatically adjusts for non-working days

## 🚀 Quick Start

1. **Open Google Apps Script**
   - Go to [script.google.com](https://script.google.com)
   - Create a new project

2. **Copy the Script**
   - Copy all code from `src/Code.gs`
   - Paste into your Apps Script project

3. **Run Setup**
```javascript
   setupAllTriggers()
```

4. **Authorize Permissions**
   - Gmail read/write access
   - Calendar management
   - Email sending

## 📋 Requirements

- Google Account with Gmail and Calendar
- Wrike account with email notifications enabled
- Google Apps Script access

## ⚙️ Configuration

Edit the `CONFIG` object to customize:
```javascript
const CONFIG = {
  timezone: "America/Phoenix",
  defaultLeadTimeHours: 72,  // Days before due date
  dailySummaryTime: 20,      // 8 PM
  // ... more settings
};
```

## 📧 How It Works

1. **Hourly Processing**: Checks for new Wrike emails from Frankie
2. **Date Extraction**: Identifies due dates (e.g., "Jan 5, 2026")
3. **Event Creation**: Schedules calendar reminder 72 hours before due
4. **Past Due Tracking**: Consolidates overdue tasks into one red event
5. **Daily Summary**: Sends email recap at 8 PM

## 📅 Calendar Event Format
```
Title: CHF0225 - Children's Hospital: Review PPC Spend [Due Jan 5]
Color: Yellow (Normal) / Red (Past Due) / Orange (Urgent)
Duration: 15 minutes
Reminder: 24 hours before
```

## 🔧 Manual Functions
```javascript
manualRun()           // Process emails now
testDailySummary()    // Send test summary
markTaskComplete()    // Mark task as done
debugLastEmail()      // Troubleshoot extraction
```

## 📊 Daily Summary Email

Sent at 8 PM daily containing:
- Past due tasks (red alert)
- Today's tasks with times
- Tomorrow's preview
- Weekly outlook
- Client workload distribution

## 🐛 Troubleshooting

See [TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) for common issues.

## 📝 License

MIT License - see [LICENSE](LICENSE)

## 👤 Author

John Williams - Senior Paid Media Specialist

## 🤝 Contributing

Pull requests welcome! Please read contributing guidelines first.
