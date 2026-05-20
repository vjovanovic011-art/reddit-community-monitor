# Reddit Community Monitor

A Google Apps Script tool that monitors recovery and addiction-related subreddits to help community outreach teams identify and respond to people seeking help.

## What It Does

- **Monitors** 18+ subreddits (recovery-focused and local Florida communities) for relevant posts
- **Matches** posts against 30+ keywords related to addiction, recovery, and treatment-seeking behavior
- **Analyzes** each matched post using AI to score relevance (1-10) and determine response type
- **Drafts** helpful, empathetic responses for human review
- **Logs** everything to a Google Sheet with approval workflow (NEW → APPROVED/REJECTED → POSTED)
- **Alerts** via email for high-priority posts and sends daily digest summaries

## What It Does NOT Do

- ❌ No automated posting — every response is manually reviewed by a human
- ❌ No vote manipulation — does not upvote, downvote, or interact with posts automatically
- ❌ No private data access — only reads public posts
- ❌ No user tracking or profiling

## Architecture

```
Reddit API (read-only) → Google Apps Script → Google Sheet (human review) → Manual posting
                              ↓
                         AI API (draft generation)
                              ↓
                         Gmail (alerts & digest)
```

## API Usage

- **Reddit API**: ~20-40 read requests per run, runs every 30 minutes, well within rate limits
- **AI API**: 1 request per matched post for relevance scoring and draft generation
- **Gmail**: Alert emails for high-priority posts + daily digest

## Setup

1. Create a Google Sheet
2. Go to Extensions > Apps Script
3. Paste the script
4. Add credentials to Script Properties (File > Project Properties > Script Properties):
   - `REDDIT_CLIENT_ID`
   - `REDDIT_CLIENT_SECRET`
   - `REDDIT_USERNAME`
   - `REDDIT_PASSWORD`
   - `REDDIT_USER_AGENT`
   - `CLAUDE_API_KEY`
   - `NOTIFICATION_EMAIL`
5. Run `setupSheet()` once
6. Run `setupTriggers()` once

## License

MIT
