# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

Google Apps Script (GAS) project that syncs events from 伝助 (Densuke, a Japanese group scheduling service) to Google Calendar. The entire script lives in a single file: [gas/Code.gs](gas/Code.gs).

## Development Workflow

This is a GAS project — there is no local build, lint, or test command. Development cycle:

1. Edit `gas/Code.gs` locally
2. Copy the entire contents into the GAS editor at https://script.google.com
3. Run `testScheduledSync()` in the GAS editor to test
4. Check execution logs in the GAS editor console

## Configuration

Settings are stored as GAS Script Properties (not in code). Run `setupProperties()` once to create the template, then edit values in **Project Settings → Script Properties**:

| Property | Description |
|----------|-------------|
| `ENABLED` | `true` to enable auto-sync |
| `DENSUKE_URL` | Full 伝助 URL (`https://densuke.biz/list?cd=XXXXXXXX`) |
| `CALENDAR_ID` | `primary` or a shared calendar's ID |
| `TARGET_MEMBER` | Member name as it appears in 伝助 (e.g., `XX.メンバー名`). Empty = **team mode**: no response symbol in titles, no "あなたの回答/メンバー" lines, reminders on every event (`config.teamMode`) |

Files under `_private/` (git-ignored) are personal local notes and must never be committed.

## Architecture

The sync pipeline in `doScheduledSync()` follows these steps:

1. **Fetch** — `fetchDensukeHtml()` fetches the 伝助 attendance table via `UrlFetchApp`
2. **Parse** — `parseDensukeHtml()` extracts events (rows) and the member's attendance response; `parseComments()` extracts the comment section
3. **Convert** — `convertToCalendarEvents()` builds calendar event objects with attendance counts, member lists (within 1 week only), and comments; IDs are generated via `generateEventId()` using a hash of date + title
4. **Sync** — `syncEvents()` compares existing calendar events (identified by `[DENSUKE:id]` tag in description) against current 伝助 data, then adds/updates/deletes accordingly
5. **Notify** — `sendNotifications()` sends email for add/update/delete; same-day delete+add collapses into a single "update" email
6. **Schedule next** — `scheduleNextSync()` creates a dynamic 5-minute trigger if an event starts within ±1 hour; otherwise sets a trigger 1 hour before the next event

## Trigger Management

- **Static trigger**: `setupHourlyTrigger()` runs `doScheduledSync` near the top of every hour
- **Dynamic trigger**: Near events, a short-interval trigger is created and stored in Script Properties under key `dynamicTriggerId`; it is deleted and recreated on each sync via `deleteDynamicTriggers()` / `createDynamicTrigger()`
- `removeAllTriggers()` cleans up all `doScheduledSync` triggers

## Event ID Scheme

Event IDs have the format `densuke_YYYYMMDD_HHHHHH` where the suffix is a 6-char hex hash of the event title. This allows stable identification of same-day events across syncs. IDs are embedded in Google Calendar event descriptions as `[DENSUKE:id]`.

## Key Parsing Notes

- `parseEventText()` handles: single-day with/without time, all-day, date ranges (`3/20～22`), OR patterns (`3/20～21または21～22`), and `上記、` (inherits previous row's time/location)
- Emojis in event titles (e.g., `🟥🟨`) are preserved throughout
- Column positions for ◎○△× are auto-detected from the header row, so the parser works even if columns are rearranged or absent
- Year is inferred: if the parsed month is more than 6 months in the past relative to now, the next year is assumed
