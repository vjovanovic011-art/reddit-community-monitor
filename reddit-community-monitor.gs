// ============================================================
// REDDIT COMMUNITY MONITOR
// Google Apps Script — monitors recovery subreddits for 
// people seeking help, logs to Google Sheet for human review
// ============================================================
// No automated posting — all responses are manually reviewed
// and posted by a human community outreach coordinator.
// ============================================================

// ======================== CONFIG ========================
// All credentials stored in Script Properties (File > Project Properties > Script Properties)
// Required properties: REDDIT_CLIENT_ID, REDDIT_CLIENT_SECRET, REDDIT_USERNAME, 
// REDDIT_PASSWORD, REDDIT_USER_AGENT, CLAUDE_API_KEY, NOTIFICATION_EMAIL

const CONFIG = {
  REDDIT_CLIENT_ID: PropertiesService.getScriptProperties().getProperty('REDDIT_CLIENT_ID'),
  REDDIT_CLIENT_SECRET: PropertiesService.getScriptProperties().getProperty('REDDIT_CLIENT_SECRET'),
  REDDIT_USERNAME: PropertiesService.getScriptProperties().getProperty('REDDIT_USERNAME'),
  REDDIT_PASSWORD: PropertiesService.getScriptProperties().getProperty('REDDIT_PASSWORD'),
  REDDIT_USER_AGENT: PropertiesService.getScriptProperties().getProperty('REDDIT_USER_AGENT'),
  CLAUDE_API_KEY: PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY'),
  NOTIFICATION_EMAIL: PropertiesService.getScriptProperties().getProperty('NOTIFICATION_EMAIL'),
  
  SHEET_NAME: 'Reddit Monitor',
  LOG_SHEET_NAME: 'Log',
  CHECK_INTERVAL_MINUTES: 30,
  MAX_POSTS_PER_RUN: 25,
  RELEVANCE_THRESHOLD: 6,
  DAILY_DIGEST_HOUR: 9,
};

// ======================== TARGET SUBREDDITS ========================
// Recovery-focused and local Florida subreddits
const TARGET_SUBREDDITS = [
  'stopdrinking',
  'redditorsinrecovery',
  'addiction',
  'alcoholism',
  'OpiatesRecovery',
  'Sober',
  'alcoholicsanonymous',
  'NarcoticsAnonymous',
  'benzorecovery',
  'leaves',
  'florida',
  'southflorida',
  'boyntonbeach',
  'delraybeach',
  'bocaraton',
  'WestPalmBeach',
  'fortlauderdale',
  'Miami',
];

// ======================== KEYWORDS ========================
// Terms indicating someone may be seeking treatment or help
const KEYWORDS = [
  'looking for rehab', 'need rehab', 'rehab florida', 'rehab south florida',
  'treatment center florida', 'detox florida', 'rehab recommendation',
  'rehab near me', 'best rehab', 'affordable rehab', 'rehab that takes insurance',
  'inpatient florida', 'outpatient florida', 'PHP program', 'IOP program',
  'dual diagnosis', 'need help with addiction', 'how to get sober',
  'ready to get clean', 'tired of drinking', 'want to quit',
  'family member needs help', 'intervention', 'rock bottom',
  'withdrawal symptoms', 'medical detox', 'suboxone', 'MAT program',
  'sober living florida', 'boynton beach', 'delray beach', 'boca raton',
  'palm beach county', 'south florida treatment', 'florida recovery',
];

// ======================== AI SYSTEM PROMPT ========================
// Instructions for the AI to generate helpful, non-promotional responses
const SYSTEM_PROMPT = `
You are a community outreach coordinator who works in addiction treatment in South Florida.
Your goal is to provide genuinely helpful responses to people seeking recovery resources.

RESPONSE GUIDELINES:
1. Always lead with empathy and helpful advice
2. Never sound like an advertisement
3. Share general recovery wisdom and practical tips
4. If someone is in crisis, prioritize safety resources (988 Lifeline, SAMHSA 1-800-662-4357)
5. Use natural Reddit conversational tone
6. Keep responses 2-4 paragraphs max
7. Only mention specific treatment options when directly asked
8. Reference professional knowledge naturally (e.g., "from working in treatment...")

RESPONSE TYPES:
- TYPE A: Someone asks for FL rehab recommendations → helpful advice + resources
- TYPE B: Someone seeking help generally → empathy + guidance + general resources  
- TYPE C: Information question about treatment → educational response
- TYPE D: Someone sharing their story → pure encouragement and support
`;

// ======================== SHEET SETUP ========================
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  
  const headers = [
    'Timestamp', 'Subreddit', 'Post Title', 'Post URL', 'Author',
    'Post Text (excerpt)', 'Matched Keywords', 'Relevance Score',
    'Response Type', 'Draft Response', 'Status', 'Notes', 'Reddit Post ID',
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff');
  
  // Data validation for Status column
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['NEW', 'APPROVED', 'REJECTED', 'POSTED', 'SKIPPED'])
    .build();
  sheet.getRange('K2:K1000').setDataValidation(statusRule);
  sheet.setFrozenRows(1);
  
  // Conditional formatting
  const rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('APPROVED').setBackground('#d4edda')
    .setRanges([sheet.getRange('K2:K1000')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('REJECTED').setBackground('#f8d7da')
    .setRanges([sheet.getRange('K2:K1000')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('NEW').setBackground('#fff3cd')
    .setRanges([sheet.getRange('K2:K1000')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('POSTED').setBackground('#cce5ff')
    .setRanges([sheet.getRange('K2:K1000')]).build());
  sheet.setConditionalFormatRules(rules);
  
  // Log sheet
  let logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!logSheet) logSheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
  logSheet.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Event', 'Details']]);
  logSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  
  Logger.log('Sheet setup complete');
}

// ======================== TRIGGERS ========================
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  
  ScriptApp.newTrigger('runMonitor')
    .timeBased()
    .everyMinutes(CONFIG.CHECK_INTERVAL_MINUTES)
    .create();
  
  ScriptApp.newTrigger('sendDailyDigest')
    .timeBased()
    .atHour(CONFIG.DAILY_DIGEST_HOUR)
    .everyDays(1)
    .create();
  
  Logger.log('Triggers configured');
}

// ======================== REDDIT AUTH ========================
function getRedditToken_() {
  const cache = CacheService.getScriptCache();
  let token = cache.get('reddit_token');
  if (token) return token;
  
  const credentials = Utilities.base64Encode(
    CONFIG.REDDIT_CLIENT_ID + ':' + CONFIG.REDDIT_CLIENT_SECRET
  );
  
  const response = UrlFetchApp.fetch('https://www.reddit.com/api/v1/access_token', {
    method: 'post',
    headers: {
      'Authorization': 'Basic ' + credentials,
      'User-Agent': CONFIG.REDDIT_USER_AGENT,
    },
    payload: {
      grant_type: 'password',
      username: CONFIG.REDDIT_USERNAME,
      password: CONFIG.REDDIT_PASSWORD,
    },
    muteHttpExceptions: true,
  });
  
  const data = JSON.parse(response.getContentText());
  if (data.access_token) {
    cache.put('reddit_token', data.access_token, 3500);
    return data.access_token;
  }
  
  throw new Error('Reddit auth failed: ' + JSON.stringify(data));
}

// ======================== FETCH POSTS ========================
function fetchSubredditPosts_(subreddit, token) {
  const url = `https://oauth.reddit.com/r/${subreddit}/new?limit=${CONFIG.MAX_POSTS_PER_RUN}&raw_json=1`;
  
  try {
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token,
        'User-Agent': CONFIG.REDDIT_USER_AGENT,
      },
      muteHttpExceptions: true,
    });
    
    if (response.getResponseCode() !== 200) {
      logEvent_('ERROR', `Failed to fetch r/${subreddit}: ${response.getResponseCode()}`);
      return [];
    }
    
    return JSON.parse(response.getContentText()).data.children.map(c => c.data);
  } catch (e) {
    logEvent_('ERROR', `Exception fetching r/${subreddit}: ${e.message}`);
    return [];
  }
}

// ======================== KEYWORD MATCHING ========================
function matchKeywords_(post) {
  const text = (post.title + ' ' + (post.selftext || '')).toLowerCase();
  return KEYWORDS.filter(kw => text.includes(kw.toLowerCase()));
}

// ======================== AI ANALYSIS ========================
function analyzeAndDraft_(post, matchedKeywords) {
  const postContent = `
SUBREDDIT: r/${post.subreddit}
TITLE: ${post.title}
TEXT: ${(post.selftext || '').substring(0, 1500)}
MATCHED KEYWORDS: ${matchedKeywords.join(', ')}
  `.trim();
  
  const prompt = `Analyze this Reddit post and generate a helpful response if appropriate.

${postContent}

Return your analysis as JSON (no markdown, no backticks, raw JSON only):
{
  "relevance_score": <1-10>,
  "response_type": "<A|B|C|D|SKIP>",
  "reasoning": "<1 sentence>",
  "draft_response": "<draft comment or empty string if SKIP>"
}

SCORING: 9-10 = explicitly seeking FL treatment, 7-8 = asking about rehab generally, 
5-6 = sharing struggles, 3-4 = recovery discussion, 1-2 = not relevant`;

  try {
    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': CONFIG.CLAUDE_API_KEY,
        'anthropic-version': '2023-06-01',
      },
      payload: JSON.stringify({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 1000,
        system: SYSTEM_PROMPT,
        messages: [{ role: 'user', content: prompt }],
      }),
      muteHttpExceptions: true,
    });
    
    const data = JSON.parse(response.getContentText());
    const clean = data.content[0].text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    return JSON.parse(clean);
  } catch (e) {
    logEvent_('ERROR', `AI API error: ${e.message}`);
    return { relevance_score: 0, response_type: 'SKIP', reasoning: 'API error', draft_response: '' };
  }
}

// ======================== MAIN MONITOR ========================
function runMonitor() {
  const token = getRedditToken_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  
  // Load existing post IDs to avoid duplicates
  const existingIds = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 13, lastRow - 1, 1).getValues().forEach(row => {
      if (row[0]) existingIds.add(row[0]);
    });
  }
  
  let newPosts = 0;
  let totalChecked = 0;
  
  for (const subreddit of TARGET_SUBREDDITS) {
    const posts = fetchSubredditPosts_(subreddit, token);
    totalChecked += posts.length;
    
    for (const post of posts) {
      if (existingIds.has(post.id)) continue;
      if (post.stickied || post.distinguished === 'moderator') continue;
      
      const matchedKeywords = matchKeywords_(post);
      
      // For local FL subreddits, check broader recovery terms
      const isLocalSub = ['florida','southflorida','boyntonbeach','delraybeach',
                          'bocaraton','WestPalmBeach','fortlauderdale','Miami'].includes(subreddit);
      
      if (matchedKeywords.length === 0 && !isLocalSub) continue;
      
      if (matchedKeywords.length === 0 && isLocalSub) {
        const text = (post.title + ' ' + (post.selftext || '')).toLowerCase();
        const localTerms = ['rehab','recovery','addiction','sober','detox','treatment','drug','alcohol'];
        if (!localTerms.some(t => text.includes(t))) continue;
        matchedKeywords.push('(local subreddit match)');
      }
      
      const analysis = analyzeAndDraft_(post, matchedKeywords);
      if (analysis.relevance_score < CONFIG.RELEVANCE_THRESHOLD) continue;
      
      sheet.appendRow([
        new Date(),
        'r/' + post.subreddit,
        post.title,
        `https://www.reddit.com${post.permalink}`,
        post.author,
        (post.selftext || '').substring(0, 500).replace(/\n/g, ' '),
        matchedKeywords.join(', '),
        analysis.relevance_score,
        analysis.response_type,
        analysis.draft_response,
        'NEW',
        analysis.reasoning,
        post.id,
      ]);
      
      newPosts++;
      existingIds.add(post.id);
    }
    
    Utilities.sleep(1500); // Rate limiting
  }
  
  logEvent_('MONITOR', `Checked ${totalChecked} posts, added ${newPosts} relevant`);
  
  if (newPosts > 0) sendNewPostAlert_(newPosts);
}

// ======================== EMAIL ALERTS ========================
function sendNewPostAlert_(count) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const lastRow = sheet.getLastRow();
  
  const highPriority = [];
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
    const oneHourAgo = new Date(Date.now() - 3600000);
    data.forEach(row => {
      if (row[0] > oneHourAgo && row[7] >= 8 && row[10] === 'NEW') {
        highPriority.push({ subreddit: row[1], title: row[2], url: row[3], score: row[7] });
      }
    });
  }
  
  if (highPriority.length === 0 && count < 3) return;
  
  let body = `Reddit Monitor: ${count} new relevant post(s)\n\n`;
  if (highPriority.length > 0) {
    body += `HIGH PRIORITY (score 8+):\n`;
    highPriority.forEach(p => { body += `\n[${p.score}/10] ${p.subreddit} — ${p.title}\n${p.url}\n`; });
  }
  body += `\nReview: ${sheetUrl}`;
  
  MailApp.sendEmail({
    to: CONFIG.NOTIFICATION_EMAIL,
    subject: `[Reddit Monitor] ${count} new posts`,
    body: body,
  });
}

// ======================== DAILY DIGEST ========================
function sendDailyDigest() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  
  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const yesterday = new Date(Date.now() - 86400000);
  
  let stats = { total: 0, new_: 0, approved: 0, posted: 0 };
  
  data.forEach(row => {
    if (row[0] > yesterday) {
      stats.total++;
      if (row[10] === 'NEW') stats.new_++;
      if (row[10] === 'APPROVED') stats.approved++;
      if (row[10] === 'POSTED') stats.posted++;
    }
  });
  
  if (stats.total === 0) return;
  
  MailApp.sendEmail({
    to: CONFIG.NOTIFICATION_EMAIL,
    subject: `[Reddit Digest] ${stats.total} posts | ${stats.new_} awaiting review`,
    body: `Daily Digest:\nTotal: ${stats.total}\nAwaiting review: ${stats.new_}\nApproved: ${stats.approved}\nPosted: ${stats.posted}\n\nReview: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`,
  });
}

// ======================== MANUAL POST (human-triggered) ========================
// Posts APPROVED responses — must be triggered manually from the Sheet menu
function postApprovedResponses() {
  const token = getRedditToken_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  
  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  
  data.forEach((row, i) => {
    if (row[10] === 'APPROVED' && row[9]) {
      try {
        UrlFetchApp.fetch('https://oauth.reddit.com/api/comment', {
          method: 'post',
          headers: {
            'Authorization': 'Bearer ' + token,
            'User-Agent': CONFIG.REDDIT_USER_AGENT,
          },
          payload: { thing_id: 't3_' + row[12], text: row[9] },
          muteHttpExceptions: true,
        });
        sheet.getRange(i + 2, 11).setValue('POSTED');
        Utilities.sleep(3000);
      } catch (e) {
        logEvent_('ERROR', `Post error: ${e.message}`);
      }
    }
  });
}

// ======================== CUSTOM MENU ========================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Reddit Monitor')
    .addItem('Run Monitor Now', 'runMonitor')
    .addItem('Post Approved Responses', 'postApprovedResponses')
    .addItem('Send Daily Digest', 'sendDailyDigest')
    .addSeparator()
    .addItem('Setup Sheet', 'setupSheet')
    .addItem('Setup Triggers', 'setupTriggers')
    .addToUi();
}

// ======================== LOGGING ========================
function logEvent_(event, details) {
  try {
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.LOG_SHEET_NAME);
    logSheet.appendRow([new Date(), event, details]);
    const lastRow = logSheet.getLastRow();
    if (lastRow > 501) logSheet.deleteRows(2, lastRow - 501);
  } catch (e) {
    Logger.log('Log error: ' + e.message);
  }
}
