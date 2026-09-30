# seaberry-mail-mcp

A local stdio MCP server that gives Claude **read + write access to your own Outlook mailbox, calendar, and tasks** — and nothing else.

> **Status:** reference implementation of an opinionated-safety MCP server. Not a maintained product. No SLA. See [SECURITY.md](SECURITY.md) for the threat model and intended deployment.

This is a deliberate alternative to broad-consent integrations: only three delegated Microsoft Graph scopes (`Mail.ReadWrite`, `Calendars.ReadWrite`, `Tasks.ReadWrite`), all calls hit `/me/`, no application permissions, no Teams / SharePoint / Planner / tenant-wide reads, no other users' data.

## Why this exists

Most public MCP servers for Microsoft 365 follow the pattern "give me the broadest scope, trust me." This one does the opposite. The interesting code isn't the Graph calls — it's the layered safety design:

- **Tool-name separation** for destructive operations. `block_time` is structurally incapable of sending invites; `create_meeting` is the only tool that can. The dangerous capability announces itself in the tool name.
- **Explicit confirmation flags** on anything that leaves the mailbox. `create_meeting` requires `sendInvites: true` to actually fire invites; default is a no-op preview.
- **Per-domain allowlist** for invites, configured in `~/.config/mail-mcp/config.json`. Defaults to **empty** — every attendee is rejected on a fresh install until the user opts in explicitly.
- **HTTP-pipeline DELETE guard.** `ReadOnlyGuardHandler` default-denies every outbound `DELETE`. The only allowlisted path is `/me/todo/lists/.../tasks/...`. Adding a new permitted DELETE path requires a deliberate code edit in one place.
- **Out-of-band provisioning.** The MCP runtime never invokes `az`. App-registration creation is a one-shot script the user runs from their own shell before launching the MCP. Keeps the runtime tool surface free of identity-system mutation authority.

If you're building your own MCP server, the patterns above are the takeaways more than the code itself.

## What it does

| Tool               | What it does                                                                   |
| ------------------ | ------------------------------------------------------------------------------ |
| `ensure_setup`     | Bootstraps the MSAL token cache (one-time interactive sign-in). Idempotent.    |
| `list_messages`    | Recent messages in a folder (default Inbox), lean metadata, for triage.        |
| `search_messages`  | Mail search by query (KQL-style operators supported). Paged automatically.     |
| `get_message`      | Full body of one message; text-extracted or HTML.                              |
| `create_draft`     | Create a draft email. Never sends — review and send from Outlook yourself.     |
| `archive_messages` | Soft-archive Inbox messages to "Archived by AI" subfolder. Not a delete.       |
| `unarchive_messages` | Restore previously-archived messages back to Inbox.                          |
| `list_events`      | Calendar over a date range. Uses `calendarView` so recurring events expand.    |
| `search_events`    | Find events by subject substring; optionally constrained to a date range.      |
| `block_time`       | Personal calendar block — no attendees, structurally cannot send invites.      |
| `create_meeting`   | Calendar event with attendees. Two safety gates before invites are sent.       |
| `list_todo_lists`  | List the user's To Do lists (each has its own id).                             |
| `list_todos`       | List tasks in a list (defaults to the "Tasks" list, open tasks only).          |
| `create_todo`      | Create a new task — title, optional body, due date, importance.                |
| `update_todo`      | Patch fields on a task (only fields you pass are changed).                     |
| `complete_todo`    | Convenience: mark a task done.                                                 |
| `delete_todo`      | Permanently delete a task (explicit id required; no bulk-delete-by-query).     |

**To Do vs Planner:** this MCP integrates with **Microsoft To Do** (personal, lives in your Exchange mailbox, `/me/todo/...`). Microsoft Planner is a separate, team-based product tied to M365 Groups — intentionally not supported here because it would require widening into group/tenant data.

## Safety gates on `create_meeting`

Calendar invites are the one thing this tool can do that leaves your mailbox, so they are gated twice:

1. **Explicit `sendInvites: true`** — default is `false`; the tool returns a preview instead of writing to Graph until you re-invoke with `sendInvites: true`.
2. **Domain allowlist** — every attendee's email domain must appear in `AllowedInviteDomains` in `~/.config/mail-mcp/config.json`. **Default is `[]`** — every attendee is rejected on a fresh install until you opt in explicitly. To allow `example.com`, edit the file to `"AllowedInviteDomains": ["example.com"]` and restart the MCP.

`create_draft` has no domain check because drafts never leave the mailbox — you send them from Outlook manually.

### Defense-in-depth: path-aware DELETE guard

Microsoft's `Mail.ReadWrite` and `Tasks.ReadWrite` scopes are misleadingly named — at the OAuth layer both also permit **delete**. The MCP defends against this with `ReadOnlyGuardHandler`, a `DelegatingHandler` inserted into the Graph HTTP pipeline that **default-denies every outbound `DELETE`** with an explicit allowlist for paths where deletion is a supported feature.

Current allowlist:

- `DELETE /me/todo/lists/.../tasks/...` — task-level deletion (the `delete_todo` tool).

Notably **not** allowed (and never should be without serious thought):

- `DELETE /me/messages/{id}` — deleting an email.
- `DELETE /me/mailFolders/{id}` — deleting a whole mail folder.
- `DELETE /me/events/{id}` — deleting a calendar event (and notifying attendees).
- `DELETE /me/todo/lists/{id}` — deleting a whole task list (and every task in it).

Adding a new permitted deletion path requires a deliberate code edit in `ReadOnlyGuardHandler.cs`. To delete the other types of items, use Outlook directly.

### Soft-archive instead of delete

Because real delete on messages is permanently off the table, the cleanup workflow is a **move** to a dedicated subfolder named "Archived by AI" (configurable via `ArchiveFolderName` in `config.json`). `archive_messages` moves Inbox messages into that folder; `unarchive_messages` puts them back. Properties:

- **Nothing is destroyed.** Messages remain in the mailbox under a clearly-labeled folder, visible to anyone auditing the mailbox.
- **Atomic per message.** Each move either succeeds or fails individually; no batch transaction. Failures surface per-id without aborting the rest.
- **Source-folder gated.** `archive_messages` accepts only messages currently in Inbox; messages already elsewhere (Sent, Drafts, the archive itself) are skipped per-id with an explicit reason. `unarchive_messages` mirrors this — accepts only messages currently in the archive folder.
- **Fixed destination.** Both tools move to a single, configured folder. No "move to arbitrary folder" surface is exposed.
- **Batch cap of 25.** Prevents a stray "archive everything" prompt from sweeping the whole inbox in one shot.
- **Lazy folder create.** The archive folder is created on the first archive call if it doesn't already exist.

## Prerequisites

- macOS, Linux, or WSL.
- `.NET 10 SDK` installed (`dotnet --version` should print `10.x`).
- `az` CLI installed, signed in as a tenant admin of the target Microsoft 365 tenant. (`az login`)
- `jq` (used by `setup.sh`).

> **A note on privilege.** The `az` + tenant-admin requirement above is a one-time *provisioning* footprint — creating the app registration and granting consent. Once that's done, the running MCP only ever holds the three narrow delegated `/me/` scopes; it never invokes `az`, never touches identity-system state, and never carries any privilege beyond what your own user account has against your own mailbox. The "least-privilege" claim is about runtime, where the ongoing risk lives — provisioning is heavier by necessity but strictly bounded to a one-shot that you run deliberately, from your own shell, with full visibility into what's happening.

## One-time setup

### 1. Provision the Entra app registration

From your shell (not from Claude):

```bash
./setup.sh
```

This is idempotent. It will:

- Create the public-client app `local-mail-mcp` if it doesn't exist (reuses it if it does).
- Stage all three delegated Microsoft Graph permissions: `Mail.ReadWrite`, `Calendars.ReadWrite`, and `Tasks.ReadWrite`. Scope IDs are resolved dynamically from the live Graph service principal in your tenant, not hardcoded.
- Atomically replace the app's required-resource-accesses (re-running the script corrects any prior wrong state rather than accumulating orphan permissions).
- Grant admin consent (with a retry loop for the service-principal replication race).
- Write `~/.config/mail-mcp/config.json` with `ClientId`, `TenantId`, `AllowedInviteDomains` (defaults to `[]` on fresh installs — see "Widening the allowlist"; preserved across re-runs), and `ArchiveFolderName` (defaults to `"Archived by AI"`; preserved across re-runs).

The MCP server itself does **not** invoke `az`. Provisioning is a one-shot script you run deliberately from your shell.

### 2. Build and register with Claude Desktop

Build the server:

```bash
dotnet build src/SeaberryMailMcp.csproj
```

Then edit your Claude Desktop config file:

- **macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows:** `%APPDATA%\Claude\claude_desktop_config.json`

Add an entry to `mcpServers` (merge with any existing entries — keep their trailing commas right):

```json
{
  "mcpServers": {
    "seaberry-mail": {
      "command": "/absolute/path/to/dotnet",
      "args": [
        "/absolute/path/to/seaberry-mail-mcp/src/bin/Debug/net10.0/SeaberryMailMcp.dll"
      ]
    }
  }
}
```

**Absolute paths are required.** Claude Desktop launches subprocesses with a stripped PATH, so `dotnet` on its own won't resolve. Find your absolute `dotnet` path with `which dotnet` (macOS/Linux) or `where dotnet` (Windows).

After editing the config, fully quit Claude Desktop (⌘Q on macOS — not just closing the window) and relaunch. The MCP appears in the tools panel.

### 3. First-run sign-in

From Claude, invoke `ensure_setup`. The first call will trigger an interactive browser sign-in (or a device-code prompt as a fallback for WSL / headless environments). Subsequent calls return instantly using the cached token.

The token cache is stored OS-encrypted (Keychain on macOS, DPAPI on Windows, libsecret on Linux) under `~/.config/mail-mcp/msal.cache.bin`.

On headless Linux, where no keyring is available, set `MAIL_MCP_ALLOW_FILE_CACHE=1` to store the cache as a plain file readable only by the current user (mode 0600). Sign-in then uses the device code flow.

## Widening the allowlist

A fresh install starts with `AllowedInviteDomains: []` — no attendee can be invited until you opt in. To allow your own domain (e.g. `example.com`):

```bash
jq '.AllowedInviteDomains += ["example.com"]' ~/.config/mail-mcp/config.json > /tmp/c && mv /tmp/c ~/.config/mail-mcp/config.json
```

Then restart the MCP. `setup.sh` is safe to re-run — it preserves any edits you've made to `AllowedInviteDomains`.

## Example invocations

- *"What came into my inbox today?"* → `list_messages` (defaults to Inbox, 25 most recent).
- *"Find emails from Alice about Q3 budget."* → `search_messages` with query `from:alice subject:"Q3 budget"`.
- *"Show me the full body of that third message."* → `get_message` with the id, `bodyFormat: "text"`.
- *"Draft a reply to Bob saying I'll review by Friday."* → `create_draft`.
- *"Clean up the newsletters and notifications in my Inbox."* → `search_messages` to enumerate ids, then `archive_messages` to move them to "Archived by AI". Restore with `unarchive_messages` if needed.
- *"What's on my calendar this week?"* → `list_events` with the week's range.
- *"Block 2-4pm tomorrow for deep work."* → `block_time`.
- *"Set up a 30-min sync with carol@example.com tomorrow at 10am."* → `create_meeting` (first call returns a preview; re-invoke with `sendInvites: true` to send — assuming `example.com` is in your allowlist).

## Not in scope

- **No `send_mail` tool.** Outgoing mail is your manual step — review the draft in Outlook and hit Send. If added later, it must apply the same domain allowlist as `create_meeting`.
- **No reads or writes against other users' mailboxes/calendars.** Only `/me/` endpoints.
- **No Teams chat, SharePoint, OneDrive, or transcripts.** Those are exactly the surfaces this tool intentionally avoids.
- **No remote/HTTP deployment.** The safety properties documented above only hold for the stdio-from-Claude-Desktop deployment described in [SECURITY.md](SECURITY.md). Exposing this as an HTTPS MCP endpoint for a hosted LLM would require a fundamentally different auth model — re-evaluate from scratch if you go there.

## Architecture notes

Single C# project, .NET 10 LTS, packages:

- `ModelContextProtocol` — official C# MCP SDK, stdio transport.
- `Microsoft.Identity.Client` + `Microsoft.Identity.Client.Extensions.Msal` — MSAL.NET public client with OS-encrypted token cache.
- `Microsoft.Graph` — typed Graph SDK with built-in retry handler (respects `Retry-After` on 429).

Logging goes to **stderr only** — stdout is the MCP protocol channel; a stray log line corrupts the transport.

`search_messages` correctly applies `ConsistencyLevel: eventual` on every paged request (the Graph SDK does **not** propagate headers across `@odata.nextLink` fetches by default — a common source of "page 2 returns 400" bugs).

Calendar tools default to **Europe/Stockholm** for both input interpretation and returned event times (via the `Prefer: outlook.timezone="..."` header).
