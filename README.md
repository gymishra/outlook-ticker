# Outlook Desktop Ticker

An always-on-top, scrolling desktop widget for Windows that surfaces your
**Outlook emails, calendar, tasks, and Slack DMs** in a compact ticker bar — plus
an optional **LED meeting badge** integration. It talks **directly to the Outlook
desktop app over COM** (no server APIs, no auth tokens), so it works with whatever
account your Outlook is already signed into.

> Companion project: this repo also documents the **Amazon Quick "Inbox To Tasks"
> scheduled agent**, which auto-creates Outlook To-Do tasks from your inbox, Slack,
> and activity feed. The ticker then displays those tasks. See
> [Inbox To Tasks agent](#inbox-to-tasks-agent-amazon-quick-desktop).

---

## Features

- **Four scrolling ticker bars**, always on top:
  - **Emails** — last 10 inbox messages (sender, subject, time). New (<15 min) items sparkle; meeting invites and conflicts are color-coded.
  - **Calendar** — next 5 days of meetings with conflict detection and inline Accept/Decline.
  - **Tasks** — open Outlook/To-Do tasks across all task lists; the top task sparkles.
  - **Slack DMs** — recent direct messages (via a cached snapshot, see below).
- **Floating vertical panel** — Tasks + Slack in a scrollable, auto-scrolling list.
- **Top "Current / Upcoming" task sticky** — cycles between your current and next task.
- **Interactive**: click to open/delete emails, accept/decline meetings; drag to reposition; adjust scroll speed; hover to pause.
- **LED badge integration** (optional) — pushes your next meetings to a Bluetooth/USB LED name badge.
- **Leak-free Outlook COM access** — every MAPI session is explicitly released each cycle (see [COM hygiene](#com-hygiene)).

---

## How it gets its data

The ticker reads from **four sources**. Email, calendar, and tasks come straight
from the locally running Outlook desktop application via Windows COM/MAPI:

| Bar | Source | Mechanism |
|-----|--------|-----------|
| Emails | Outlook **Inbox** | COM `GetDefaultFolder(6)` |
| Calendar | Outlook **Calendar** | COM `GetDefaultFolder(9)` |
| Tasks | Outlook **Tasks** (+ subfolders) | COM `GetDefaultFolder(13)` |
| Slack DMs | `slack_cache.json` | refreshed via `kiro-cli` subprocess |

> Email/calendar/task data requires the **Outlook desktop app to be running**.
> There is no Microsoft Graph call and no stored credential — it automates the
> Outlook process directly. The Slack bar is a cached snapshot, not live.

### COM hygiene

Each fetch (`fetch_emails`, `fetch_calendar`, `fetch_tasks`) uses
`GetFirst()`/`GetNext()` enumeration and **explicitly releases every COM
reference** (`app`, namespace, folders, items) with `gc.collect()` before
`CoUninitialize()`. This prevents the MAPI-session exhaustion that can otherwise
make Outlook unable to connect after the app runs for a while.

---

## Requirements

- Windows 10/11
- Microsoft **Outlook desktop** (classic), signed in
- Python 3.10+
- Python packages:
  ```
  pip install pywin32
  ```
- (Optional) `kiro-cli` on PATH for the Slack bar
- (Optional) `bleak` + a compatible LED badge for the LED integration

---

## Running

```bash
python outlook_ticker.py
```

Build a standalone Windows executable:

```bash
pyinstaller --onefile --noconsole --name OutlookTicker outlook_ticker.py
```

### Run at login

Drop a shortcut to the exe (or a `pythonw outlook_ticker.py` shortcut) into:

```
%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
```

---

## Controls

| Action | Result |
|--------|--------|
| Hover a bar | Pause scrolling |
| Click an email | Open it (or search it in OWA) |
| Click ✕ on an email | Delete it |
| Click Y / N on a meeting | Accept / Decline |
| Drag the left icon | Reposition the bar |
| ◀ / ▶ (− / +) | Adjust scroll speed |
| Drag the right edge | Resize the bar width |

---

## Configuration

Key constants at the top of `outlook_ticker.py`:

| Constant | Default | Meaning |
|----------|---------|---------|
| `REFRESH_SEC` | `60` | How often Outlook is polled (seconds) |
| `SCROLL_SPEED` | `2` | Pixels per animation frame |
| `BAR_W` / `BAR_H` | `900` / `38` | Bar dimensions |

LED badge settings live in `led_badge_config.json` (see `led_meeting_badge.py`).

---

## Inbox To Tasks agent (Amazon Quick Desktop)

The ticker only **reads** tasks. Tasks are **created** by an **Amazon Quick Desktop
scheduled agent** named **"Inbox To Tasks"**. It continuously monitors your inbox,
Slack, and activity feed, then writes actionable tasks into your Outlook **To-Do
"Work Tasks"** list. The Outlook Ticker's Tasks bar (and the floating panel) then
surface those tasks — so the agent *produces* and the ticker *displays*.

```
                 Amazon Quick Desktop
        ┌──────────────────────────────────────┐
        │  "Inbox To Tasks"  (scheduled agent)  │
        │   every ~5 min via quickwork-agent    │
        └───────────────┬──────────────────────┘
                         │ reads
     ┌───────────────────┼───────────────────────┐
     ▼                   ▼                        ▼
  Outlook            Slack (AWS              Activity Feed
  inbox/cal          Slack V2)               (get_feed_item)
     │                   │                        │
     └───────────────────┴───────────────────────┘
                         │ creates actionable tasks
                         ▼
            Outlook To-Do  ▸  "Work Tasks" list
                         │ read by COM GetDefaultFolder(13)
                         ▼
                   Outlook Ticker  ▸  Tasks bar + panel
```

The agent is configured through three tabs in the Quick Desktop UI:
**Schedule**, **Capabilities**, and **Task objectives & model**.

### Capabilities tab

**Preferred Slack workspace:** `AWS Slack V2` — the agent's Slack tool calls target
this workspace.

| Capability | Description | Tools enabled |
|------------|-------------|---------------|
| **Outlook** | Read and manage emails, calendar events, and tasks | 16 |
| **Slack** | Read and send Slack messages, manage channels and reactions | 11 |
| **Knowledge Graph Tools** | Search, add, edit, and manage knowledge graph entities and relationships | 3 |
| **Feed Tools** | Activity feed, notifications, briefings, and day plans | 2 |

These map to the underlying MCP tools the agent calls, e.g.:
`email_inbox`, `email_search`, `todo_tasks` (Outlook); `search_messages` (Slack);
`get_feed_item`, `update_feed`, `skip_cycle` (Feed Tools); plus knowledge-graph
lookups for context.

### Task objectives & model tab — the agent prompt

```text
You are a task management agent for Gyan Mishra (gyanmis@amazon.com). Your job is
to monitor incoming emails, Slack DMs/mentions, and activity feed items, then
create actionable tasks in Outlook To-Do when something requires Gyan's response
or attention.

## Your Workflow
1. Check for new emails  — email_inbox(unreadOnly=true) or email_search to find
   recent unread emails directed at Gyan that require a response or action.
2. Check Slack           — search_messages to find recent DMs and mentions
   directed at Gyan (query: "to:me") that need a response.
3. Check the activity feed — get_feed_item to see if there are important feed
   items that imply action is needed.
4. Evaluate importance   — for each item, decide if it genuinely requires Gyan's
   action or response. Skip:
     - Automated notifications, newsletters, mass distributions
     - FYI-only messages that don't need a reply
     - Bot messages and system alerts
     - Items that are purely informational
5. Check existing tasks  — todo_tasks(operation="list", listId="<WORK_TASKS_LIST_ID>")
   to see current tasks and avoid creating duplicates.
6. Create tasks          — for actionable items, create a task in "Work Tasks":
     - listId:      <WORK_TASKS_LIST_ID>
     - title:       Clear, actionable (e.g., "Reply to [sender] re: [subject]")
     - body:        Brief context — who sent it, what they need, any deadline
     - importance:  "high" if urgent/time-sensitive, "normal" otherwise
     - dueDateTime: end of today if urgent, end of tomorrow if normal
7. Post to feed          — after creating tasks, call update_feed(importance="fyi")
   to log what you did (e.g., "Created 2 tasks from new emails"). If nothing
   actionable was found, call skip_cycle().

## What Qualifies as Actionable
- Someone asking Gyan a direct question
- A request for review, approval, or feedback
- Action items assigned to Gyan
- Meeting-related requests (scheduling, prep needed)
- Emails where Gyan is in the To: line (not just CC) and a response is expected
- Important escalations or time-sensitive matters

## What to Skip
- Newsletters, automated reports, system notifications
- CC-only emails that are FYI
- Messages where someone else already responded
- Slack channel noise that doesn't mention Gyan directly
- Tasks that already exist in the Work Tasks list (avoid duplicates)

## Writing Style for Task Titles
- Start with an action verb: "Reply to...", "Review...", "Approve...", "Follow up on..."
- Include the sender/requestor name
- Keep it concise but informative
```

> **Configuration note:** `<WORK_TASKS_LIST_ID>` is the id of your Outlook To-Do
> "Work Tasks" list. It is environment-specific — keep your real list id out of
> source control and substitute it in the Quick Desktop agent config.

### Schedule tab

Set the run cadence here (this agent runs roughly **every 5 minutes** via the
background `quickwork-agent` process). Lower intervals mean fresher tasks but more
frequent inbox/Slack reads. The agent persists run state under Quick's profile
(`scheduler_state/inbox-to-tasks.json`).

### Relationship to the ticker

The agent and the ticker are **decoupled** — they communicate only through the
Outlook To-Do store:

- **Agent (writer):** evaluates inbox/Slack/feed → creates tasks in "Work Tasks".
- **Ticker (reader):** `fetch_tasks()` reads `GetDefaultFolder(13)` and all task
  subfolders, sorts "Work Tasks" + high-importance items to the front, and scrolls
  them. No coupling, no shared process — either can run without the other.

---

## Files

| File | Purpose |
|------|---------|
| `outlook_ticker.py` | The ticker app (emails / calendar / tasks / Slack + panel) |
| `led_meeting_badge.py` | LED badge driver (BLE/USB) |
| `led_badge_config.json` | LED badge settings |
| `OutlookTicker.spec` | PyInstaller spec |

---

## Notes & limitations

- Requires the **classic Outlook desktop** app running (COM automation).
- Slack data is a cached snapshot via `kiro-cli`, not a live connection.
- Designed for a single Windows user profile.

## License

MIT — see `LICENSE` (add one if publishing publicly).
