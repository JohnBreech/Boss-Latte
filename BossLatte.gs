/**
 * ============================================================
 * BOSS LATTE — Weekly Strategy Review Automation
 * ============================================================
 * Runs every Sunday at 6pm via a time-based trigger.
 * Pulls Google Calendar + Google Tasks for the past week,
 * sends everything to Gemini for a full executive review,
 * emails you the report, and logs scores to a Google Sheet
 * (with a trend chart) so you can see week-over-week progress.
 *
 * SETUP (one-time):
 * 1. Create a new Google Sheet. Copy its ID from the URL
 *    (the long string between /d/ and /edit) into SHEET_ID below.
 * 2. In Apps Script: Project Settings > Script Properties,
 *    add a property GEMINI_API_KEY with your Gemini API key.
 * 3. Update MY_EMAIL below.
 * 4. Run setupTrigger() once (manually, from the editor) to
 *    install the Sunday 6pm trigger. You only need to do this once.
 * 5. Run runBossLatte() once manually to test before trusting
 *    the automation.
 * ============================================================
 */

const SHEET_ID = 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE';
const MY_EMAIL = 'your-email@example.com';
const GEMINI_MODEL = 'gemini-2.0-flash'; // swap for whichever Gemini model you use in Axile

const LIFE_AREAS = [
  'Academics', 'Career/Freelancing', 'Software Projects', 'Learning',
  'Finances', 'Health', 'Relationships', 'Rest/Recovery'
];

const SCORE_FIELDS = ['productivity', 'focus', 'energy', 'goalAlignment', 'consistency'];

/**
 * Installs the weekly Sunday 6pm trigger. Run this ONCE manually.
 */
function setupTrigger() {
  // Clear any existing Boss Latte triggers first, to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runBossLatte') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('runBossLatte')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(18) // 6pm, in the script's project timezone
    .create();

  Logger.log('Boss Latte trigger installed: every Sunday at 6pm.');
}

/**
 * Main entry point — this is what the trigger calls.
 */
function runBossLatte() {
  const weekData = gatherWeekData_();
  const reviewText = callGemini_(buildPrompt_(weekData));
  const entry = parseEntry_(reviewText);

  emailReview_(reviewText, entry);
  logToSheet_(entry);
  updateChart_();

  Logger.log('Boss Latte review complete for ' + entry.date);
}

/**
 * Pulls Calendar events (past 7 days + next 7 days) and Google Tasks.
 */
function gatherWeekData_() {
  const now = new Date();
  const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  const weekAhead = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);

  // --- Calendar ---
  const cal = CalendarApp.getDefaultCalendar();
  const pastEvents = cal.getEvents(weekAgo, now).map(e => ({
    title: e.getTitle(),
    start: e.getStartTime(),
    end: e.getEndTime(),
  }));
  const upcomingEvents = cal.getEvents(now, weekAhead).map(e => ({
    title: e.getTitle(),
    start: e.getStartTime(),
    end: e.getEndTime(),
  }));

  // --- Tasks ---
  let completedTasks = [];
  let incompleteTasks = [];
  try {
    const taskLists = Tasks.Tasklists.list().items || [];
    taskLists.forEach(list => {
      const tasks = Tasks.Tasks.list(list.id, { showCompleted: true, maxResults: 100 }).items || [];
      tasks.forEach(t => {
        const record = { title: t.title, due: t.due || null, notes: t.notes || '' };
        if (t.status === 'completed') completedTasks.push(record);
        else incompleteTasks.push(record);
      });
    });
  } catch (e) {
    Logger.log('Tasks API not enabled or no access: ' + e);
  }

  return { pastEvents, upcomingEvents, completedTasks, incompleteTasks };
}

/**
 * Builds the full Boss Latte prompt, injecting this week's real data
 * and, if available, a summary of the last few weeks' scores for trend context.
 */
function buildPrompt_(weekData) {
  const history = getRecentHistory_(4); // last 4 weeks for trend context

  const historyBlock = history.length
    ? history.map(h => `- ${h.date}: Overall ${h.overallScore}/10`).join('\n')
    : 'No previous entries yet — this is the first review.';

  const eventsBlock = (list) => list.length
    ? list.map(e => `- ${Utilities.formatDate(e.start, Session.getScriptTimeZone(), 'EEE MMM d, h:mma')}: ${e.title}`).join('\n')
    : '(none)';

  const tasksBlock = (list) => list.length
    ? list.map(t => `- ${t.title}${t.due ? ' (due ' + t.due + ')' : ''}`).join('\n')
    : '(none)';

  return `You are Boss Latte, my executive assistant, strategist, accountability
partner, and weekly coach. You run automatically every Sunday at 6pm.

Today is my Weekly Strategy Review.

Your job is NOT to congratulate me for being busy.
Your job is to determine whether the way I spent my time actually moved
me closer to my long-term goals.

==========================
MY LONG TERM VISION
==========================
1. Finish my Computer Science degree with good academic performance.
2. Become financially independent by finding remote freelance work or internships.
3. Build high-quality software products that become portfolio pieces.
4. Improve my software engineering skills every week.
5. Maintain my health, relationships, and hobbies without burning out.

Whenever you make recommendations, these goals always take priority.

==========================
THIS WEEK'S REAL DATA
==========================
Calendar — past 7 days:
${eventsBlock(weekData.pastEvents)}

Calendar — next 7 days:
${eventsBlock(weekData.upcomingEvents)}

Completed tasks this week:
${tasksBlock(weekData.completedTasks)}

Incomplete / outstanding tasks:
${tasksBlock(weekData.incompleteTasks)}

Recent score history (for trend context):
${historyBlock}

==========================
TL;DR — READ THIS FIRST
==========================
Before anything else, output a short skimmable summary, max 6 lines
total, no sub-bullets:
- Overall Score: X/10 (one line, include the number and a 1-sentence verdict)
- Biggest win this week (1 sentence)
- Biggest miss this week (1 sentence)
- Top priority for next week (1 sentence)
This block must stand alone — I should be able to read only this and
know exactly where I stand, even if I read nothing else below.

==========================
STEP 1 — WEEKLY SCORECARD
==========================
Overall Week Score (/10), Productivity, Focus, Energy, Goal Alignment,
Consistency scores. Then max 3 sentences on where time really went and
whether this week moved me closer to my goals or was just "busy."

==========================
STEP 2 — GOAL ALIGNMENT
==========================
For each area, give a 1-line Evidence note, a 1-line Recommended
adjustment, and an Alignment score /10. Keep each area to 2 lines max,
no separate Progress/What's working/What's not subheadings:
${LIFE_AREAS.join(' · ')}

==========================
STEP 3 — TIME AUDIT
==========================
Estimate hours per category in a short list. One closing sentence on
whether the allocation was aligned with my goals.

==========================
STEP 4 — ACCOMPLISHMENTS & BOTTLENECKS
==========================
Combine these: up to 3 wins (1 line each) and up to 2 recurring
bottlenecks, each with a 1-line fix. Do not use separate
Major/Small/Momentum/Hidden Wins subcategories — just a flat list.

==========================
STEP 5 — PROJECT HEALTH
==========================
Status, risk, and next milestone for each active project, 1-2 lines
per project. Be honest if something no longer serves my goals.

==========================
STEP 6 — NEXT WEEK PRIORITIES
==========================
Only 3-5, ranked, each with a one-line reason. This is the most
important section — do not pad it.

==========================
STEP 7 — NEXT WEEK FOCUS
==========================
Weekly Mission (one sharp sentence). One habit to strengthen, one to
reduce, one to stop, one to double down on — 1 line each.

==========================
STEP 8 — RISKS
==========================
Up to 2 likely problems next week, each with a 1-line prevention step.

==========================
STEP 9 — PERSONAL COACH
==========================
Honest, encouraging closing message, max 3 sentences. Consistency
beats intensity.

==========================
IMPORTANT RULES
==========================
Never reward busyness. Recommend fewer, higher-impact tasks. Protect
recovery time. Connect advice to the score history above when relevant.

==========================
FINAL OUTPUT — REQUIRED
==========================
After the full narrative review, output a fenced code block titled
"BOSS LATTE ENTRY" containing ONLY valid JSON in this exact shape
(no comments, no trailing commas):

{
  "date": "YYYY-MM-DD",
  "overallScore": 0,
  "scores": {"productivity":0,"focus":0,"energy":0,"goalAlignment":0,"consistency":0},
  "areas": {
    "Academics": {"score": 0, "note": ""},
    "Career/Freelancing": {"score": 0, "note": ""},
    "Software Projects": {"score": 0, "note": ""},
    "Learning": {"score": 0, "note": ""},
    "Finances": {"score": 0, "note": ""},
    "Health": {"score": 0, "note": ""},
    "Relationships": {"score": 0, "note": ""},
    "Rest/Recovery": {"score": 0, "note": ""}
  },
  "mission": ""
}

Use "date" = today's date, ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}.`;
}

/**
 * Calls the Gemini API with the built prompt. Mirrors however Axile
 * calls Gemini — adjust the endpoint/model if Axile uses a different one.
 */
function callGemini_(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('Missing GEMINI_API_KEY in Script Properties.');

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${apiKey}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const json = JSON.parse(response.getContentText());
  if (!json.candidates || !json.candidates.length) {
    throw new Error('Gemini returned no candidates: ' + response.getContentText());
  }
  return json.candidates[0].content.parts[0].text;
}

/**
 * Extracts the "BOSS LATTE ENTRY" JSON block from the model's response.
 */
function parseEntry_(reviewText) {
  const match = reviewText.match(/```(?:json)?\s*([\s\S]*?)```/);
  if (!match) throw new Error('Could not find BOSS LATTE ENTRY JSON block in response.');

  let entry;
  try {
    entry = JSON.parse(match[1].trim());
  } catch (e) {
    throw new Error('BOSS LATTE ENTRY block was not valid JSON: ' + e);
  }
  if (!entry.date || typeof entry.overallScore !== 'number') {
    throw new Error('BOSS LATTE ENTRY missing required fields (date, overallScore).');
  }
  return entry;
}

/**
 * Emails the full review to yourself as a styled HTML email.
 */
function emailReview_(reviewText, entry) {
  const cleanMarkdown = reviewText.replace(/```[\s\S]*?```/g, '').trim();
  const htmlBody = markdownToEmailHtml_(cleanMarkdown, entry);
  const plainFallback = cleanMarkdown.replace(/[#*]/g, '');

  GmailApp.sendEmail(
    MY_EMAIL,
    `☕ Boss Latte — Weekly Review (${entry.date}) — Score: ${entry.overallScore}/10`,
    plainFallback,
    { htmlBody: htmlBody }
  );
}

/**
 * Converts the model's markdown-style review into styled HTML for email.
 * Handles: #### / ### headers, **bold**, *** horizontal rules,
 * "* " bullet lists, and paragraph breaks.
 */
function markdownToEmailHtml_(md, entry) {
  const escapeHtml = (s) => s
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  const inline = (line) => escapeHtml(line)
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/\*(.+?)\*/g, '<em>$1</em>');

  const lines = md.split('\n');
  let html = '';
  let inList = false;
  let inBox = false;

  const closeList = () => { if (inList) { html += '</ul>'; inList = false; } };
  const closeBox = () => { if (inBox) { html += '</div>'; inBox = false; } };

  lines.forEach(raw => {
    const line = raw.trim();

    if (!line) { closeList(); return; }

    if (/^={5,}$/.test(line)) { return; } // decorative ==== rule lines, skip
    if (/^\*{3,}$/.test(line)) { closeList(); closeBox(); html += '<hr style="border:none;border-top:1px solid #E4D9C8;margin:24px 0;">'; return; }

    if (/^TL;DR/i.test(line)) {
      closeList(); closeBox();
      html += `<div style="background:#FBF3E7;border-left:4px solid #E3A23C;border-radius:6px;padding:14px 18px 6px;margin:4px 0 20px;">`;
      html += `<div style="font-family:Georgia,serif;color:#3A2A1D;font-weight:bold;font-size:13px;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:6px;">TL;DR</div>`;
      inBox = true;
      return;
    }
    if (/^####\s+/.test(line)) { closeList(); closeBox(); html += `<h4 style="font-family:Georgia,serif;color:#3A2A1D;margin:20px 0 8px;">${inline(line.replace(/^####\s+/, ''))}</h4>`; return; }
    if (/^###\s+/.test(line)) { closeList(); closeBox(); html += `<h3 style="font-family:Georgia,serif;color:#3A2A1D;border-bottom:2px solid #E3A23C;padding-bottom:6px;margin:28px 0 12px;">${inline(line.replace(/^###\s+/, ''))}</h3>`; return; }

    if (/^\*\s+/.test(line) || /^-\s+/.test(line)) {
      if (!inList) { html += '<ul style="margin:8px 0;padding-left:22px;">'; inList = true; }
      html += `<li style="margin-bottom:6px;line-height:1.5;">${inline(line.replace(/^[\*\-]\s+/, ''))}</li>`;
      return;
    }

    closeList();
    html += `<p style="line-height:1.6;color:#3A2A1D;margin:10px 0;">${inline(line)}</p>`;
  });
  closeList();
  closeBox();

  return `
  <div style="font-family:Verdana,Arial,sans-serif;max-width:640px;margin:0 auto;background:#FBF3E7;padding:0;">
    <div style="background:#3A2A1D;padding:24px 28px;border-radius:10px 10px 0 0;">
      <div style="font-family:Georgia,serif;color:#F7ECDD;font-size:22px;font-weight:bold;">☕ Boss Latte</div>
      <div style="color:#C9B49C;font-size:12px;letter-spacing:1px;text-transform:uppercase;margin-top:2px;">Weekly Strategy Review — ${entry.date}</div>
      <div style="color:#E3A23C;font-family:Georgia,serif;font-size:28px;font-weight:bold;margin-top:10px;">${entry.overallScore} / 10</div>
    </div>
    <div style="background:#FFFFFF;padding:8px 28px 28px;border:1px solid #E4D9C8;border-top:none;border-radius:0 0 10px 10px;">
      ${html}
    </div>
  </div>`;
}

/**
 * Appends a row to the Google Sheet. Creates the header row and
 * sheet structure if this is the first run.
 */
function logToSheet_(entry) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Weekly Log');
  if (!sheet) {
    sheet = ss.insertSheet('Weekly Log');
    const headers = ['Date', 'Overall Score', ...SCORE_FIELDS, ...LIFE_AREAS, 'Mission'];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }

  const row = [
    entry.date,
    entry.overallScore,
    ...SCORE_FIELDS.map(f => (entry.scores && entry.scores[f] != null) ? entry.scores[f] : ''),
    ...LIFE_AREAS.map(a => (entry.areas && entry.areas[a]) ? entry.areas[a].score : ''),
    entry.mission || '',
  ];
  sheet.appendRow(row);
}

/**
 * Reads the last N weeks of Overall Score from the sheet, for trend
 * context to feed back into next week's prompt.
 */
function getRecentHistory_(n) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Weekly Log');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const rows = data.slice(1); // drop header
    const recent = rows.slice(-n);
    return recent.map(r => ({ date: r[0], overallScore: r[1] }));
  } catch (e) {
    return [];
  }
}

/**
 * Creates (or refreshes) a line chart of Overall Score over time
 * on the "Weekly Log" sheet, so you get visual trends natively in Sheets.
 */
function updateChart_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Weekly Log');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return; // need at least 2 data points for a meaningful line

  // Remove existing Boss Latte chart before adding a fresh one
  sheet.getCharts().forEach(chart => sheet.removeChart(chart));

  const dateRange = sheet.getRange(1, 1, lastRow, 1);   // Date column
  const scoreRange = sheet.getRange(1, 2, lastRow, 1);  // Overall Score column

  const chart = sheet.newChart()
    .asLineChart()
    .addRange(dateRange)
    .addRange(scoreRange)
    .setPosition(2, LIFE_AREAS.length + SCORE_FIELDS.length + 5, 0, 0)
    .setOption('title', 'Overall Score Trend')
    .setOption('legend', { position: 'none' })
    .setOption('vAxis', { minValue: 0, maxValue: 10, title: 'Score /10' })
    .setOption('hAxis', { title: 'Week' })
    .setOption('colors', ['#E3A23C'])
    .build();

  sheet.insertChart(chart);
}