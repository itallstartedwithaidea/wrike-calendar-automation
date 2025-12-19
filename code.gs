/**
 * Wrike to Google Calendar Automation
 * Version: 1.0.0
 * Author: John Williams
 * 
 * Main entry point for the automation script
 */

// ====================
// WRIKE TO GOOGLE CALENDAR AUTOMATION - COMPLETE SCRIPT
// ====================

// ====================
// CONFIGURATION
// ====================
const CONFIG = {
  triggerEmail: "wrike@app-us2.wrike.com",
  triggerPerson: "Frankie",  // ANY Frankie notification
  assigneeName: "First Last Name of Who Task Is Assigned",
  
  // Set your timezone
  timezone: "America/Phoenix",  // Arizona doesn't observe DST
  
  // Default lead time in hours before due date
  defaultLeadTimeHours: 72,  // 3 days before due date
  
  // Task type patterns with lead times (in hours)
  taskPatterns: {
    "QA All Shape Budgets": 72,
    "Final Approval Shape": 72,
    "Review PPC Spend File": 72,
    "QA All Budgets": 72,
    "Upload Budgets": 48,
    "Review Budget Changes": 48,
    "Approve Campaign Changes": 72,
    "Budget Pacing Review": 48,
    "Month-End Reconciliation": 96,
    "Client Reporting": 48,
    "Weekly QA of Budget Pacing": 72,
    "Monthly Budget Maintenance": 72,
    "Monthly Spend Maintenance": 72,
    "Campaign Refresh": 72,
    "Final Review": 72
  },
  
  // Priority keywords that need faster turnaround
  urgentKeywords: ["URGENT", "ASAP", "Critical", "Rush", "Priority", "Escalated", "overdue"],
  
  // Calendar settings - Using string IDs for colors
  calendar: {
    colors: {
      normal: "5",  // YELLOW
      late: "11",   // RED
      urgent: "6",  // ORANGE
      completed: "2", // GREEN
      blocked: "8"  // GRAY
    },
    eventDuration: 15,  // minutes
    reminderTime: 1440, // 24 hours in minutes
    urgentReminderTime: 120, // 2 hours for urgent items
    workHoursStart: 8,
    workHoursEnd: 17,
    preferredSlots: ["09:00", "11:00", "14:00", "16:00"],
    consolidatePastDue: true  // Consolidate all past due into one event
  },
  
  // Notification thresholds
  alerts: {
    criticalClientsNeedEarlier: ["PHXART", "CHF", "DKI", "WECU"],
    holidayBuffer: 2, // Extra days before holidays
    maxDailyTasks: 8,  // Alert if more than 8 tasks in one day
    dailySummaryTime: 20,  // 8 PM
    sendIndividualPastDueAlerts: false  // DISABLE individual past due emails
  }
};

// Global variable to collect past due tasks
let globalPastDueTasks = [];

// ====================
// MANUAL RUN FUNCTION
// ====================
function manualRun() {
  Logger.log('Manual run triggered at ' + new Date().toLocaleString());
  processWrikeEmails();
  return "Processing complete - check your calendar and Gmail";
}

// ====================
// TIMEZONE VERIFICATION
// ====================
function verifyAndSetTimezone() {
  const scriptTimeZone = Session.getScriptTimeZone();
  const calendarTimeZone = CalendarApp.getDefaultCalendar().getTimeZone();
  
  Logger.log(`Script Timezone: ${scriptTimeZone}`);
  Logger.log(`Calendar Timezone: ${calendarTimeZone}`);
  Logger.log(`Configured Timezone: ${CONFIG.timezone}`);
  
  if (scriptTimeZone !== CONFIG.timezone || calendarTimeZone !== CONFIG.timezone) {
    Logger.log(`Warning: Timezone mismatch detected. Using ${CONFIG.timezone} for calculations.`);
  }
  
  return CONFIG.timezone;
}

// ====================
// MAIN FUNCTION - RUN HOURLY
// ====================
function processWrikeEmails() {
  try {
    // Reset past due collection
    globalPastDueTasks = [];
    
    // Verify timezone
    verifyAndSetTimezone();
    
    // First, check and update any existing calendar events for past due status
    updatePastDueEvents();
    
    // Search for ALL unread Wrike emails from Frankie
    const searchQuery = `from:${CONFIG.triggerEmail} is:unread`;
    const threads = GmailApp.search(searchQuery);
    
    Logger.log(`Found ${threads.length} unread Wrike emails to process`);
    
    const dailyTaskCount = {};  // Track tasks per day for overload warning
    const processedTasks = [];  // Track for summary
    
    threads.forEach(thread => {
      const messages = thread.getMessages();
      
      messages.forEach(message => {
        if (message.isUnread()) {
          const result = processWrikeEmail(message);
          
          if (result) {
            processedTasks.push(result);
            
            // Track daily load for non-past-due tasks
            if (!result.isPastDue && result.eventDate) {
              const dateKey = result.eventDate.toDateString();
              dailyTaskCount[dateKey] = (dailyTaskCount[dateKey] || 0) + 1;
            }
          }
        }
      });
    });
    
    // Create or update consolidated past due event
    if (globalPastDueTasks.length > 0) {
      createConsolidatedPastDueEvent(globalPastDueTasks);
    }
    
    // Check for overload
    checkForOverload(dailyTaskCount);
    
    // Log processed tasks for daily summary
    if (processedTasks.length > 0) {
      storeProcessedTasks(processedTasks);
    }
    
    Logger.log(`Completed processing: ${processedTasks.length} tasks (${globalPastDueTasks.length} past due)`);
    
  } catch (error) {
    console.error('Error in main processing:', error);
    sendErrorNotification(error);
  }
}

// ====================
// EMAIL PROCESSOR - WITH EMAIL DATE TRACKING
// ====================
function processWrikeEmail(message) {
  try {
    const body = message.getPlainBody();
    const subject = message.getSubject();
    const emailDate = message.getDate(); // Capture when email was received
    
    const taskInfo = extractTaskInfo(body, subject, emailDate);
    
    if (!taskInfo.taskType) {
      Logger.log('Could not extract task type from email');
      return null;
    }
    
    // If no due date found, set a reasonable default (7 days from email date)
    if (!taskInfo.dueDate) {
      taskInfo.dueDate = new Date(emailDate);
      taskInfo.dueDate.setDate(taskInfo.dueDate.getDate() + 7);
      Logger.log(`No due date found, defaulting to 7 days from email date: ${taskInfo.dueDate}`);
    }
    
    const now = new Date();
    now.setHours(0, 0, 0, 0); // Compare dates only, not times
    
    const dueDate = new Date(taskInfo.dueDate);
    dueDate.setHours(0, 0, 0, 0);
    
    // Task is past due ONLY if due date is before today OR explicitly marked as overdue
    const isPastDue = dueDate < now || taskInfo.isExplicitlyOverdue;
    const isUrgent = checkIfUrgent(taskInfo, body);
    const isBlocked = body.toLowerCase().includes('blocked');
    
    Logger.log(`Task: ${taskInfo.taskType}, Due: ${taskInfo.dueDate}, Past Due: ${isPastDue}`);
    
    // Mark email as read first
    message.markRead();
    
    // Add label
    const labelName = isPastDue ? 'Wrike/PastDue' : 
                     isUrgent ? 'Wrike/Urgent' : 
                     'Wrike/Processed';
    
    const label = GmailApp.getUserLabelByName(labelName) || 
                  GmailApp.createLabel(labelName);
    message.getThread().addLabel(label);
    
    if (isPastDue) {
      // Add to past due collection for consolidation
      globalPastDueTasks.push(taskInfo);
      Logger.log(`Added to past due list: ${taskInfo.clientCode}: ${taskInfo.taskType}`);
      return { taskInfo, isPastDue: true, isUrgent };
      
    } else {
      // Calculate when to schedule the reminder task (48-72 hours before due)
      const leadTimeHours = CONFIG.taskPatterns[taskInfo.taskType] || CONFIG.defaultLeadTimeHours;
      const msBeforeDue = leadTimeHours * 60 * 60 * 1000; // Convert hours to milliseconds
      
      // Calculate the reminder date
      let reminderDate = new Date(taskInfo.dueDate.getTime() - msBeforeDue);
      
      // If reminder date would be in the past, schedule for tomorrow
      if (reminderDate < now) {
        reminderDate = new Date(now);
        reminderDate.setDate(reminderDate.getDate() + 1);
        reminderDate.setHours(9, 0, 0, 0); // 9 AM tomorrow
        Logger.log(`Reminder date would be in past, scheduling for tomorrow: ${reminderDate}`);
      }
      
      // Skip weekends
      const dayOfWeek = reminderDate.getDay();
      if (dayOfWeek === 0) { // Sunday -> Monday
        reminderDate.setDate(reminderDate.getDate() + 1);
      } else if (dayOfWeek === 6) { // Saturday -> Monday
        reminderDate.setDate(reminderDate.getDate() + 2);
      }
      
      // Check if event already exists
      const existingEvent = findExistingEvent(taskInfo);
      
      let eventColor = isUrgent ? CONFIG.calendar.colors.urgent : 
                      isBlocked ? CONFIG.calendar.colors.blocked :
                      CONFIG.calendar.colors.normal;
      
      if (existingEvent) {
        Logger.log(`Event already exists for ${taskInfo.clientCode}: ${taskInfo.taskType}`);
        updateExistingEvent(existingEvent, taskInfo, eventColor, false);
      } else {
        createCalendarEvent(taskInfo, reminderDate, eventColor, false, isUrgent);
        Logger.log(`Created new event for ${taskInfo.clientCode}: ${taskInfo.taskType} on ${reminderDate}`);
      }
      
      return { eventDate: reminderDate, taskInfo, isPastDue: false, isUrgent };
    }
    
  } catch (error) {
    console.error('Error processing email:', error);
    return null;
  }
}

// ====================
// EXTRACT TASK INFO - WITH EMAIL DATE AND IMPROVED DUE DATE DETECTION
// ====================
function extractTaskInfo(body, subject, emailDate) {
  const info = {
    clientCode: null,
    clientName: null,
    taskType: null,
    campaignName: null,
    dueDate: null,
    emailDate: emailDate, // Store when email was received
    status: null,
    assignedTo: ["Your Name"],
    isExplicitlyOverdue: false
  };
  
  // Clean up the body text
  body = body.replace(/\+/g, ' ');
  try {
    body = decodeURIComponent(body);
  } catch (e) {
    // If decode fails, use original
  }
  
  // Check if explicitly marked as overdue
  info.isExplicitlyOverdue = body.toLowerCase().includes('overdue') || 
                             body.toLowerCase().includes('past due') ||
                             body.toLowerCase().includes('is overdue');
  
  // Extract client code AND full name from subject
  if (subject) {
    const subjectClientMatch = subject.match(/\[([A-Z]{2,6}\d{3,4})(?:\s*[-–]\s*([^\]]+))?\]/);
    if (subjectClientMatch) {
      info.clientCode = subjectClientMatch[1];
      info.clientName = subjectClientMatch[2] ? subjectClientMatch[2].trim() : subjectClientMatch[1];
    }
  }
  
  // If client name is cut off in subject, try to get from body
  if (info.clientName && info.clientName.includes('…')) {
    const bodyClientMatch = body.match(new RegExp(info.clientCode + '\\s*[-–]\\s*([^<\n]+)'));
    if (bodyClientMatch) {
      info.clientName = bodyClientMatch[1].trim();
    }
  }
  
  // If no client code found, use default
  if (!info.clientCode) {
    info.clientCode = "WRIKE";
    info.clientName = "Wrike Task";
  }
  
  // Extract task type from subject first
  if (subject) {
    let taskFromSubject = subject
      .replace(/\[.*?\]\s*/, '')
      .trim();
    
    if (taskFromSubject.includes('|')) {
      const parts = taskFromSubject.split('|').map(p => p.trim());
      info.taskType = parts[0];
      if (parts.length > 1) {
        info.campaignName = parts.slice(1).join(' | ');
      }
    } else {
      info.taskType = taskFromSubject;
    }
  }
  
  // Clean up task type
  if (info.taskType) {
    info.taskType = info.taskType
      .replace(/<[^>]+>/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }
  
  // IMPROVED DATE EXTRACTION
  const dateMatches = body.match(/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2}(?:,?\s+\d{4})?/gi);
  
  if (dateMatches && dateMatches.length > 0) {
    Logger.log(`Found ${dateMatches.length} dates in email: ${dateMatches.join(', ')}`);
    Logger.log(`Email was received on: ${emailDate.toDateString()}`);
    
    let selectedDate = null;
    let selectedDateStr = null;
    
    // First priority: Dates with future years (2026+)
    for (const dateStr of dateMatches) {
      if (dateStr.match(/202[6-9]/)) {
        selectedDate = new Date(dateStr);
        selectedDateStr = dateStr;
        Logger.log(`Selected future year date as due date: ${dateStr}`);
        break;
      }
    }
    
    // Second priority: Dates with current year that are in the future
    if (!selectedDate) {
      const currentYear = new Date().getFullYear();
      for (const dateStr of dateMatches) {
        if (dateStr.includes(currentYear.toString())) {
          const testDate = new Date(dateStr);
          if (testDate > emailDate) {
            selectedDate = testDate;
            selectedDateStr = dateStr;
            Logger.log(`Selected current year future date as due date: ${dateStr}`);
            break;
          }
        }
      }
    }
    
    // Third priority: Dates without year that aren't near the email date
    if (!selectedDate) {
      const emailMonth = emailDate.getMonth();
      const emailDay = emailDate.getDate();
      
      for (const dateStr of dateMatches) {
        if (!dateStr.match(/\d{4}/)) { // No year in date
          const testDate = parseWrikeDate(dateStr, emailDate);
          
          // Skip if this date is too close to email date (likely the email date itself)
          const dayDiff = Math.abs((testDate - emailDate) / (1000 * 60 * 60 * 24));
          if (dayDiff < 2) {
            Logger.log(`Skipping date too close to email date: ${dateStr}`);
            continue;
          }
          
          // Use this date if it's in the future
          if (testDate > emailDate) {
            selectedDate = testDate;
            selectedDateStr = dateStr;
            Logger.log(`Selected future date as due date: ${dateStr}`);
            break;
          }
        }
      }
    }
    
    info.dueDate = selectedDate;
    
    if (info.dueDate) {
      Logger.log(`Final due date selected: ${selectedDateStr} -> ${info.dueDate.toDateString()}`);
    }
  }
  
  // If task is explicitly overdue and no due date found, set to yesterday
  if (info.isExplicitlyOverdue && !info.dueDate) {
    info.dueDate = new Date();
    info.dueDate.setDate(info.dueDate.getDate() - 1);
    Logger.log('Task explicitly marked as overdue, setting due date to yesterday');
  }
  
  // Extract status
  const statusMatch = body.match(/(Not Started|In Progress|Complete|Done|On Hold|Blocked|At Risk)/i);
  if (statusMatch) {
    info.status = statusMatch[1];
  }
  
  // Extract assigned people
  const assignedMatch = body.match(/assigned task to ([^\n]+)/);
  
  if (assignedMatch) {
    const names = assignedMatch[1].split(/[,&]/).map(name => name.trim());
    names.forEach(name => {
      if (name && !name.includes('Reply') && !name.includes('»')) {
        info.assignedTo.push(name);
      }
    });
  }
  
  info.assignedTo = [...new Set(info.assignedTo)];
  
  Logger.log(`=== EXTRACTION COMPLETE ===`);
  Logger.log(`Client: ${info.clientCode} - ${info.clientName}`);
  Logger.log(`Task: ${info.taskType}`);
  Logger.log(`Due Date: ${info.dueDate ? info.dueDate.toDateString() : 'Not found'}`);
  Logger.log(`Email Received: ${info.emailDate.toDateString()}`);
  Logger.log(`Status: ${info.status}`);
  Logger.log(`Overdue: ${info.isExplicitlyOverdue}`);
  Logger.log(`===========================`);
  
  return info;
}

// ====================
// CREATE CALENDAR EVENT - WITH EMAIL DATE IN DESCRIPTION
// ====================
function createCalendarEvent(taskInfo, reminderDate, eventColor, isPastDue, isUrgent) {
  const calendar = CalendarApp.getDefaultCalendar();
  
  // Format due date for title
  const dueMonth = taskInfo.dueDate.toLocaleDateString('en-US', { month: 'short' });
  const dueDay = taskInfo.dueDate.getDate();
  
  // Build event title with due date
  let eventTitle = `${taskInfo.clientCode}`;
  if (taskInfo.clientName && !taskInfo.clientName.includes('…')) {
    eventTitle = `${taskInfo.clientCode} - ${taskInfo.clientName}`;
  }
  eventTitle += `: ${taskInfo.taskType} [Due ${dueMonth} ${dueDay}]`;
  
  if (isUrgent && !isPastDue) eventTitle = `🔥 ${eventTitle}`;
  
  // Set the event time
  const eventStart = new Date(reminderDate);
  if (eventStart.getHours() === 0) {
    eventStart.setHours(9, 0, 0, 0); // Default to 9 AM if no time set
  }
  
  const eventEnd = new Date(eventStart);
  eventEnd.setMinutes(eventEnd.getMinutes() + CONFIG.calendar.eventDuration);
  
  // Build description with clear due date AND email date
  const daysUntilDue = Math.ceil((taskInfo.dueDate - reminderDate) / (1000 * 60 * 60 * 24));
  
  const description = [
    `📅 DUE DATE: ${taskInfo.dueDate.toDateString()}`,
    `⏰ This reminder is ${daysUntilDue} days before the due date`,
    `📧 Task assigned via email on: ${taskInfo.emailDate.toDateString()}`,
    '',
    '═══════════════════════════════════════',
    '',
    `Client: ${taskInfo.clientName || taskInfo.clientCode}`,
    `Task: ${taskInfo.taskType}`,
    taskInfo.campaignName ? `Campaign: ${taskInfo.campaignName}` : '',
    `Status: ${taskInfo.status || 'Not Started'}`,
    `Assigned to: ${taskInfo.assignedTo.join(', ')}`,
    '',
    '─────────────────────────────',
    'This is a reminder to complete the task before the due date.',
    '',
    `Auto-created from Wrike email received on ${taskInfo.emailDate.toLocaleString()}`
  ].filter(line => line !== null && line !== '').join('\n');
  
  const event = calendar.createEvent(
    eventTitle,
    eventStart,
    eventEnd,
    {
      description: description,
      colorId: eventColor,
      reminders: {
        useDefault: false,
        overrides: [{
          method: 'email',
          minutes: CONFIG.calendar.reminderTime
        }]
      }
    }
  );
  
  Logger.log(`Created calendar event: "${eventTitle}" on ${eventStart.toDateString()}`);
  
  return event;
}

// ====================
// CREATE CONSOLIDATED PAST DUE EVENT - FOR ALL OVERDUE TASKS
// ====================
function createConsolidatedPastDueEvent(pastDueTasks) {
  if (pastDueTasks.length === 0) return;
  
  const calendar = CalendarApp.getDefaultCalendar();
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);
  
  // Check if consolidated event already exists today
  const existingEvents = calendar.getEvents(today, tomorrow);
  let consolidatedEvent = existingEvents.find(event => 
    event.getTitle().includes('[PAST DUE TASKS]')
  );
  
  // Build description with all tasks
  let description = `⚠️ ${pastDueTasks.length} PAST DUE TASKS REQUIRING IMMEDIATE ATTENTION ⚠️\n`;
  description += `═══════════════════════════════════════\n\n`;
  
  pastDueTasks.forEach((task, index) => {
    const daysOverdue = Math.floor((today - task.dueDate) / (1000 * 60 * 60 * 24));
    
    description += `${index + 1}. ${task.clientCode} - ${task.clientName || task.clientCode}\n`;
    description += `   Task: ${task.taskType}\n`;
    description += `   ${task.campaignName ? 'Campaign: ' + task.campaignName + '\n' : ''}`;
    description += `   Original Due Date: ${task.dueDate.toDateString()}\n`;
    description += `   Days Overdue: ${daysOverdue}\n`;
    description += `   Status: ${task.status || 'Not Started'}\n`;
    description += `   Email Received: ${task.emailDate ? task.emailDate.toDateString() : 'Unknown'}\n`;
    description += `   ─────────────────────────────\n\n`;
  });
  
  description += `\nAll tasks require immediate attention.\n`;
  description += `Please address these tasks or communicate blockers to the team.\n\n`;
  description += `Auto-updated: ${new Date().toLocaleString()}`;
  
  const eventTitle = `[PAST DUE TASKS] ${pastDueTasks.length} Tasks Need Immediate Attention`;
  
  if (consolidatedEvent) {
    // Update existing consolidated event
    consolidatedEvent.setTitle(eventTitle);
    consolidatedEvent.setDescription(description);
    consolidatedEvent.setColor(CONFIG.calendar.colors.late);
    Logger.log(`Updated consolidated past due event with ${pastDueTasks.length} tasks`);
  } else {
    // Create new consolidated event for today
    const eventStart = new Date();
    if (eventStart.getHours() < CONFIG.calendar.workHoursStart) {
      eventStart.setHours(CONFIG.calendar.workHoursStart, 0, 0, 0);
    } else if (eventStart.getHours() >= CONFIG.calendar.workHoursEnd) {
      eventStart.setDate(eventStart.getDate() + 1);
      eventStart.setHours(CONFIG.calendar.workHoursStart, 0, 0, 0);
    } else {
      eventStart.setHours(eventStart.getHours() + 1, 0, 0, 0);
    }
    
    const eventEnd = new Date(eventStart);
    eventEnd.setMinutes(eventEnd.getMinutes() + 30); // 30 minutes for past due review
    
    calendar.createEvent(
      eventTitle,
      eventStart,
      eventEnd,
      {
        description: description,
        colorId: CONFIG.calendar.colors.late,
        reminders: {
          useDefault: false,
          overrides: [{
            method: 'email',
            minutes: 0  // Immediate reminder
          }]
        }
      }
    );
    
    Logger.log(`Created consolidated past due event with ${pastDueTasks.length} tasks`);
  }
}

// ====================
// PARSE DATE WITH CONTEXT
// ====================
function parseWrikeDate(dateStr, emailDate) {
  const currentYear = new Date().getFullYear();
  const emailYear = emailDate ? emailDate.getFullYear() : currentYear;
  const emailMonth = emailDate ? emailDate.getMonth() : new Date().getMonth();
  
  // If year is already in the string, use it
  if (dateStr.match(/\d{4}/)) {
    return new Date(dateStr);
  }
  
  // Parse the month from the date string
  const monthMatch = dateStr.match(/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/i);
  if (!monthMatch) {
    return new Date(dateStr + ', ' + currentYear);
  }
  
  const monthMap = {
    'jan': 0, 'feb': 1, 'mar': 2, 'apr': 3, 'may': 4, 'jun': 5,
    'jul': 6, 'aug': 7, 'sep': 8, 'oct': 9, 'nov': 10, 'dec': 11
  };
  
  const dateMonth = monthMap[monthMatch[1].toLowerCase()];
  
  // Determine the year based on email date context
  let year = emailYear;
  
  // If the month is before the email month, it's likely next year
  if (dateMonth < emailMonth) {
    year = emailYear + 1;
  }
  // If it's December email and January date, it's next year
  else if (emailMonth === 11 && dateMonth === 0) {
    year = emailYear + 1;
  }
  
  const testDate = new Date(dateStr + ', ' + year);
  
  Logger.log(`Parsed ${dateStr} with email context ${emailDate?.toDateString()} -> ${testDate.toDateString()}`);
  
  return testDate;
}

// ====================
// FIND EXISTING EVENT
// ====================
function findExistingEvent(taskInfo) {
  const calendar = CalendarApp.getDefaultCalendar();
  const searchStart = new Date();
  searchStart.setDate(searchStart.getDate() - 30);
  const searchEnd = new Date();
  searchEnd.setDate(searchEnd.getDate() + 90);
  
  const events = calendar.getEvents(searchStart, searchEnd);
  
  return events.find(event => {
    const title = event.getTitle();
    
    // Skip consolidated past due event
    if (title.includes('[PAST DUE TASKS]')) return false;
    
    // Match on client code AND partial task type
    const hasClient = taskInfo.clientCode && title.includes(taskInfo.clientCode);
    const hasTask = taskInfo.taskType && 
                   title.toLowerCase().includes(taskInfo.taskType.substring(0, 20).toLowerCase());
    
    return hasClient && hasTask;
  });
}

// ====================
// UPDATE EXISTING EVENT
// ====================
function updateExistingEvent(event, taskInfo, newColor, isPastDue) {
  const currentColor = event.getColor();
  
  if (currentColor !== newColor) {
    event.setColor(newColor);
    
    const description = event.getDescription();
    const updatedDescription = description + 
      `\n\nStatus Updated: ${new Date().toLocaleString()}`;
    
    event.setDescription(updatedDescription);
    
    Logger.log(`Updated existing event: ${event.getTitle()}`);
  }
}

// ====================
// UPDATE PAST DUE EVENTS - CHECK EXISTING EVENTS
// ====================
function updatePastDueEvents() {
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - 30);
  
  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEvents(startDate, now);
  
  const newPastDueTasks = [];
  
  events.forEach(event => {
    const title = event.getTitle();
    
    // Skip the consolidated event itself
    if (title.includes('[PAST DUE TASKS]')) {
      return;
    }
    
    if (title.includes(':') && event.getColor() !== CONFIG.calendar.colors.completed) {
      const description = event.getDescription();
      
      const dueDateMatch = description.match(/DUE DATE: ([^\n]+)/);
      if (dueDateMatch) {
        const dueDate = new Date(dueDateMatch[1]);
        dueDate.setHours(0, 0, 0, 0);
        
        if (dueDate < now && !title.includes('[PAST DUE]')) {
          // Extract info from event for consolidation
          const clientMatch = title.match(/([A-Z]{2,6}\d{3,4})(?:\s*[-–]\s*([^:]+))?:/);
          if (clientMatch) {
            const emailDateMatch = description.match(/Email Received: ([^\n]+)/);
            const emailDate = emailDateMatch ? new Date(emailDateMatch[1]) : new Date();
            
            const taskInfo = {
              clientCode: clientMatch[1],
              clientName: clientMatch[2] ? clientMatch[2].trim() : clientMatch[1],
              taskType: title.split(':')[1]?.trim().replace(/\[Due.*\]/, '').trim(),
              dueDate: dueDate,
              emailDate: emailDate,
              status: 'Past Due'
            };
            newPastDueTasks.push(taskInfo);
          }
          
          // Delete the individual past due event
          event.deleteEvent();
          Logger.log(`Moved to past due consolidation: ${title}`);
        }
      }
    }
  });
  
  // If we found new past due tasks, add them to consolidated event
  if (newPastDueTasks.length > 0) {
    globalPastDueTasks.push(...newPastDueTasks);
  }
}

// ====================
// HELPER FUNCTIONS
// ====================
function checkIfUrgent(taskInfo, body) {
  const hasUrgentKeyword = CONFIG.urgentKeywords.some(keyword => 
    body.toUpperCase().includes(keyword.toUpperCase())
  );
  
  const hoursUntilDue = (taskInfo.dueDate - new Date()) / (1000 * 60 * 60);
  const isTimeSensitive = hoursUntilDue < 48;
  
  const isCriticalClient = CONFIG.alerts.criticalClientsNeedEarlier.some(client =>
    taskInfo.clientCode?.includes(client)
  );
  
  return hasUrgentKeyword || (isTimeSensitive && isCriticalClient);
}

function getUpcomingHolidays() {
  const currentYear = new Date().getFullYear();
  const nextYear = currentYear + 1;
  
  return [
    new Date(currentYear, 11, 25), // Christmas
    new Date(currentYear, 11, 31), // New Year's Eve
    new Date(nextYear, 0, 1), // New Year's Day
    new Date(nextYear, 6, 4), // July 4th
    new Date(nextYear, 10, 28), // Thanksgiving
  ];
}

function checkForOverload(dailyTaskCount) {
  const overloadedDays = [];
  
  for (const [date, count] of Object.entries(dailyTaskCount)) {
    if (count > CONFIG.alerts.maxDailyTasks) {
      overloadedDays.push(`${date}: ${count} tasks`);
    }
  }
  
  if (overloadedDays.length > 0) {
    GmailApp.sendEmail(
      Session.getActiveUser().getEmail(),
      '⚠️ Task Overload Alert',
      `The following days have more than ${CONFIG.alerts.maxDailyTasks} tasks scheduled:\n\n` +
      overloadedDays.join('\n') +
      '\n\nConsider redistributing or delegating some tasks.'
    );
  }
}

function sendErrorNotification(error) {
  GmailApp.sendEmail(
    Session.getActiveUser().getEmail(),
    'Wrike Automation Error',
    `Error processing Wrike emails:\n\n${error.toString()}\n\nStack trace:\n${error.stack}`
  );
}

function storeProcessedTasks(tasks) {
  const properties = PropertiesService.getUserProperties();
  const today = new Date().toDateString();
  
  const existingTasks = JSON.parse(properties.getProperty(today) || '[]');
  const allTasks = existingTasks.concat(tasks.map(t => ({
    client: t.taskInfo.clientCode,
    task: t.taskInfo.taskType,
    dueDate: t.taskInfo.dueDate,
    emailDate: t.taskInfo.emailDate,
    isPastDue: t.isPastDue,
    isUrgent: t.isUrgent
  })));
  
  properties.setProperty(today, JSON.stringify(allTasks));
}

// ====================
// DAILY SUMMARY EMAIL AT 8 PM
// ====================
function sendDailySummary() {
  try {
    verifyAndSetTimezone();
    
    const today = new Date();
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);
    const weekFromNow = new Date(today);
    weekFromNow.setDate(weekFromNow.getDate() + 7);
    
    const calendar = CalendarApp.getDefaultCalendar();
    
    const todayEvents = calendar.getEvents(today, tomorrow)
      .filter(event => event.getTitle().includes(':') || event.getTitle().includes('['));
    
    const tomorrowEvents = calendar.getEvents(tomorrow, new Date(tomorrow.getTime() + 86400000))
      .filter(event => event.getTitle().includes(':') || event.getTitle().includes('['));
    
    const weekEvents = calendar.getEvents(tomorrow, weekFromNow)
      .filter(event => event.getTitle().includes(':') || event.getTitle().includes('['));
    
    const pastDueEvents = todayEvents.concat(weekEvents)
      .filter(event => event.getTitle().includes('[PAST DUE'));
    
    const urgentEvents = todayEvents.concat(weekEvents)
      .filter(event => event.getTitle().includes('🔥'));
    
    const clientStats = {};
    todayEvents.concat(weekEvents).forEach(event => {
      const clientMatch = event.getTitle().match(/([A-Z]{2,6}\d{3,4})/);
      if (clientMatch) {
        const client = clientMatch[1];
        clientStats[client] = (clientStats[client] || 0) + 1;
      }
    });
    
    let emailBody = buildSummaryEmail(todayEvents, tomorrowEvents, weekEvents, pastDueEvents, urgentEvents, clientStats);
    
    GmailApp.sendEmail(
      Session.getActiveUser().getEmail(),
      `📊 Wrike Daily Summary - ${today.toLocaleDateString()}`,
      emailBody,
      {
        htmlBody: emailBody
      }
    );
    
    Logger.log('Daily summary email sent at ' + new Date().toLocaleString());
    
  } catch (error) {
    console.error('Error sending daily summary:', error);
    sendErrorNotification(error);
  }
}

// ====================
// BUILD SUMMARY EMAIL HTML
// ====================
function buildSummaryEmail(todayEvents, tomorrowEvents, weekEvents, pastDueEvents, urgentEvents, clientStats) {
  const today = new Date().toLocaleDateString();
  
  let html = `
<div style="font-family: Arial, sans-serif; max-width: 600px;">
  <h2 style="color: #333;">📊 Wrike Task Summary for ${today}</h2>
  
  ${pastDueEvents.length > 0 ? `
  <div style="background: #ffebee; padding: 10px; border-left: 4px solid #f44336; margin-bottom: 20px;">
    <h3 style="color: #c62828; margin-top: 0;">⚠️ PAST DUE (${pastDueEvents.length})</h3>
    <ul style="margin: 0;">
      ${pastDueEvents.map(event => 
        `<li>${event.getTitle().replace('[PAST DUE TASKS]', '').replace('[PAST DUE]', '').trim()}</li>`
      ).join('')}
    </ul>
  </div>
  ` : ''}
  
  ${urgentEvents.length > 0 ? `
  <div style="background: #fff3e0; padding: 10px; border-left: 4px solid #ff9800; margin-bottom: 20px;">
    <h3 style="color: #e65100; margin-top: 0;">🔥 URGENT (${urgentEvents.length})</h3>
    <ul style="margin: 0;">
      ${urgentEvents.map(event => 
        `<li>${event.getTitle().replace('🔥', '').trim()}</li>`
      ).join('')}
    </ul>
  </div>
  ` : ''}
  
  <div style="background: #e8f5e9; padding: 10px; border-left: 4px solid #4caf50; margin-bottom: 20px;">
    <h3 style="color: #2e7d32; margin-top: 0;">📅 TODAY (${todayEvents.length})</h3>
    ${todayEvents.length > 0 ? `
    <ul style="margin: 0;">
      ${todayEvents.map(event => {
        const time = event.getStartTime().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
        return `<li>${time} - ${event.getTitle()}</li>`;
      }).join('')}
    </ul>
    ` : '<p style="margin: 0;">No tasks scheduled for today</p>'}
  </div>
  
  <div style="background: #e3f2fd; padding: 10px; border-left: 4px solid #2196f3; margin-bottom: 20px;">
    <h3 style="color: #1565c0; margin-top: 0;">📅 TOMORROW (${tomorrowEvents.length})</h3>
    ${tomorrowEvents.length > 0 ? `
    <ul style="margin: 0;">
      ${tomorrowEvents.map(event => {
        const time = event.getStartTime().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
        return `<li>${time} - ${event.getTitle()}</li>`;
      }).join('')}
    </ul>
    ` : '<p style="margin: 0;">No tasks scheduled for tomorrow</p>'}
  </div>
  
  <div style="background: #f3e5f5; padding: 10px; border-left: 4px solid #9c27b0; margin-bottom: 20px;">
    <h3 style="color: #6a1b9a; margin-top: 0;">📆 THIS WEEK (${weekEvents.length})</h3>
    ${weekEvents.length > 0 ? `
    <ul style="margin: 0;">
      ${weekEvents.slice(0, 10).map(event => {
        const date = event.getStartTime().toLocaleDateString([], {weekday: 'short', month: 'short', day: 'numeric'});
        return `<li>${date} - ${event.getTitle()}</li>`;
      }).join('')}
      ${weekEvents.length > 10 ? `<li><em>... and ${weekEvents.length - 10} more</em></li>` : ''}
    </ul>
    ` : '<p style="margin: 0;">No upcoming tasks this week</p>'}
  </div>
  
  <div style="background: #e1f5fe; padding: 10px; border-left: 4px solid #0288d1; margin-bottom: 20px;">
    <h3 style="color: #01579b; margin-top: 0;">📈 CLIENT WORKLOAD</h3>
    ${Object.keys(clientStats).length > 0 ? `
    <table style="width: 100%; border-collapse: collapse;">
      ${Object.entries(clientStats)
        .sort((a, b) => b[1] - a[1])
        .map(([client, count]) => 
          `<tr>
            <td style="padding: 5px; border-bottom: 1px solid #ddd;">${client}</td>
            <td style="padding: 5px; border-bottom: 1px solid #ddd; text-align: right;">
              ${count} task${count > 1 ? 's' : ''}
            </td>
          </tr>`
        ).join('')}
    </table>
    ` : '<p style="margin: 0;">No active client tasks</p>'}
  </div>
  
  <div style="background: #f5f5f5; padding: 10px; border-radius: 4px; margin-top: 20px;">
    <h4 style="margin: 5px 0;">📊 Summary</h4>
    <p style="margin: 5px 0;">• Total Active: ${todayEvents.length + weekEvents.length}</p>
    <p style="margin: 5px 0;">• Past Due: ${pastDueEvents.length}</p>
    <p style="margin: 5px 0;">• Urgent: ${urgentEvents.length}</p>
    <p style="margin: 5px 0;">• Clients: ${Object.keys(clientStats).length}</p>
  </div>
  
  <hr style="margin: 20px 0; border: 0; border-top: 1px solid #ddd;">
  
  <p style="color: #666; font-size: 12px;">
    Generated: ${new Date().toLocaleString()}<br>
    <a href="https://calendar.google.com">View Calendar</a>
  </p>
</div>
  `;
  
  return html;
}

// ====================
// SETUP ALL TRIGGERS
// ====================
function setupAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  
  // Process emails every hour
  ScriptApp.newTrigger('processWrikeEmails')
    .timeBased()
    .everyHours(1)
    .create();
    
  // Send daily summary at 8 PM
  ScriptApp.newTrigger('sendDailySummary')
    .timeBased()
    .atHour(20)  // 8 PM
    .everyDays(1)
    .create();
  
  verifyAndSetTimezone();
    
  Logger.log(`Triggers set up: 
    - Process emails: Every hour
    - Daily summary: 8:00 PM
    - Timezone: ${CONFIG.timezone}`);
}

// ====================
// TEST FUNCTIONS
// ====================
function testDateExtraction() {
  const testBody = `Monthly Spend Maintenance | January 2025
Review PPC Spend File Not Started
CHF0225 - Ch…
Jan 5, 2026
Dec 19
New
Frankie 7:13
assigned task to "update with your name", "update with sender name"`;
  
  const testSubject = "[CHF0225 - Children's Hospital Foundation] Review PPC Spend File";
  const testEmailDate = new Date("2025-12-19");
  
  const info = extractTaskInfo(testBody, testSubject, testEmailDate);
  Logger.log('Test extraction results:');
  Logger.log('Due Date: ' + info.dueDate);
  Logger.log('Email Date: ' + info.emailDate);
  Logger.log('Client: ' + info.clientCode);
  Logger.log('Task: ' + info.taskType);
  Logger.log('Assigned: ' + info.assignedTo);
  
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  const dueDate = new Date(info.dueDate);
  dueDate.setHours(0, 0, 0, 0);
  const isPastDue = dueDate < now;
  
  Logger.log('Would be past due? ' + isPastDue);
}

function testOverdueScenario() {
  const testBody = `This task is overdue, please complete it
Weekly QA of Budget Pacing
WECU0125
Dec 18
Dec 19`;
  
  const testSubject = "[WECU0125] Weekly QA of Budget Pacing";
  const testEmailDate = new Date();
  
  const info = extractTaskInfo(testBody, testSubject, testEmailDate);
  Logger.log('Overdue test:');
  Logger.log('Is explicitly overdue: ' + info.isExplicitlyOverdue);
  Logger.log('Due Date: ' + info.dueDate);
}

function debugLastEmail() {
  const threads = GmailApp.search('from:wrike@app-us2.wrike.com', 0, 5);
  
  if (threads.length > 0) {
    const messages = threads[0].getMessages();
    const lastMessage = messages[messages.length - 1];
    
    Logger.log('Last email from Wrike:');
    Logger.log('Subject: ' + lastMessage.getSubject());
    Logger.log('From: ' + lastMessage.getFrom());
    Logger.log('Email Date: ' + lastMessage.getDate());
    Logger.log('Is Unread: ' + lastMessage.isUnread());
    Logger.log('Body preview: ' + lastMessage.getPlainBody().substring(0, 500));
    
    const taskInfo = extractTaskInfo(lastMessage.getPlainBody(), lastMessage.getSubject(), lastMessage.getDate());
    Logger.log('Extracted info: ' + JSON.stringify(taskInfo, null, 2));
  } else {
    Logger.log('No emails found from Wrike');
  }
}

function testDailySummary() {
  sendDailySummary();
}

function markTaskComplete(clientCode, taskType) {
  const calendar = CalendarApp.getDefaultCalendar();
  const searchStart = new Date();
  searchStart.setDate(searchStart.getDate() - 30);
  const searchEnd = new Date();
  searchEnd.setDate(searchEnd.getDate() + 30);
  
  const events = calendar.getEvents(searchStart, searchEnd);
  
  const event = events.find(e => 
    e.getTitle().includes(clientCode) && 
    e.getTitle().includes(taskType)
  );
  
  if (event) {
    event.setColor(CONFIG.calendar.colors.completed);
    event.setTitle(event.getTitle() + ' ✓');
    Logger.log(`Marked complete: ${event.getTitle()}`);
  }
}
