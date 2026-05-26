# Security

## Status

This is a **reference implementation** of an opinionated-safety MCP server, not a maintained product. There is no SLA, no security release schedule, and no commitment to respond to issues on any particular timeline. The author runs it for personal use and publishes the code so others can read, fork, and learn from the design.

## Threat model

The intended deployment is **strictly local**:

- The MCP server runs as a child process of a single Claude client on the user's own machine.
- It speaks stdio JSON-RPC. It does **not** expose an HTTP endpoint.
- It signs in as the *running user only*, against the user's own M365 mailbox/calendar/tasks via delegated `/me/` Graph endpoints.
- The token cache is OS-encrypted (Keychain on macOS, DPAPI on Windows, libsecret on Linux; ACL-restricted plaintext fallback otherwise).

If your deployment differs from the above (e.g. running the server on a shared machine, exposing it via a tunnel, modifying the auth to confidential-client / multi-tenant), the safety properties described in the README **do not transfer automatically**. Re-evaluate.

## Designed-in safety properties (what the code actually enforces)

1. **Least-privilege scopes.** Three delegated Graph scopes (`Mail.ReadWrite`, `Calendars.ReadWrite`, `Tasks.ReadWrite`). No application permissions. No tenant-wide access.
2. **No deletion of mail or calendar entries**, ever, regardless of what tools call what Graph paths. `ReadOnlyGuardHandler` default-denies `DELETE` at the HTTP-pipeline layer; the only allowlisted DELETE path is `/me/todo/lists/.../tasks/...`. Widening this requires a deliberate edit in one place.
3. **Email never leaves the mailbox without manual confirmation.** `create_draft` lands in Drafts; the user reviews and sends from Outlook.
4. **Calendar invites are triple-gated.** Tool-name separation (`block_time` cannot invite anyone; `create_meeting` is the only tool that can) **and** explicit `sendInvites: true` confirmation flag **and** per-domain attendee allowlist (defaults to `[]` — every attendee is rejected until the user opts in explicitly via `~/.config/mail-mcp/config.json`).
5. **No provisioning at runtime.** The MCP server never invokes `az` or any privileged CLI. Provisioning is out-of-band via `setup.sh`, run manually by the user.

## Reporting a vulnerability

If you find a vulnerability in the codebase that could affect deployments matching the intended threat model above, please open a GitHub issue with the `security` label. Do not include exploit details for unpatched issues in public posts.

For deployments outside the intended threat model — you're on your own.
