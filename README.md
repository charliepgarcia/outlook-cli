# outlook-cli

A lightweight CLI wrapper for the **Microsoft Graph v1.0 Outlook Mail API** —
read, search, send, reply, forward, draft, organize, and handle attachments
from the shell or from an agent.

Built so an AI agent can do everything email-related through a Microsoft
account without a browser.

---

## Quick start

```bash
pnpm install && pnpm build

# one-time: register an Azure app and export its client id (see SETUP.md)
export GRAPH_CLIENT_ID=<application-client-id>

outlook auth login      # device-code sign-in, ~30 seconds
outlook me              # confirm which mailbox you are on
outlook list            # newest inbox messages
```

**First time?** Read [SETUP.md](./SETUP.md). The Azure app registration is the
only step that cannot be automated, and it takes about three minutes.

---

## Safety model

Email is irreversible and reaches other people. Six deliberate guardrails:

**1. `--dry-run` on every mutating command.** Prints the exact HTTP request that
would be issued, then exits without touching the network. The short-circuit
lives in one shared chokepoint (`guardMutation`), not in each command, so no
command can forget it.

```bash
outlook send --to a@example.com --subject "Test" --body "Hi" --dry-run
```
```
DRY RUN — nothing was sent. The request below is what would go out:

POST https://graph.microsoft.com/v1.0/me/sendMail
Content-Type: application/json

{
  "message": { "subject": "Test", ... },
  "saveToSentItems": true
}
```

**2. `delete` is recoverable; `purge` is not and is gated.** `outlook delete`
*moves* the message to Deleted Items, where you can get it back. Microsoft's
documentation never states whether a raw `DELETE` is recoverable, so that path
is exposed separately as `outlook purge` and refuses to run without `--yes`.

**3. Sends are never retried.** `sendMail`, `reply`, `replyAll`, `forward` and
draft-send carry no idempotency guarantee — an automatic retry after a network
timeout can deliver the message twice. The HTTP layer retries idempotent verbs
(GET/DELETE/PATCH) on throttling, and deliberately does not retry POST.

**4. Unknown flags are rejected, not ignored.** A misspelled flag is a hard
error before anything runs. This matters more than it sounds: with a
permissive parser, `--dryrun` is silently ignored, `dryRun` stays false, and
"show me what would be sent" becomes "send it now." Every command declares its
accepted flags in `src/flagSpec.ts`, and a test asserts no command can be
registered without an entry.

```bash
outlook send --to a@example.com --subject "Test" --body "Hi" --dryrun
```
```
Unknown flag for 'send': --dryrun
Did you mean --dry-run?
```

**5. Confirmation flags cannot be faked.** `--yes` accepts only explicit
boolean spellings. `--yes=false`, `--yes=0` and `--no-yes` all mean *no* — a
user typing `--yes=false` to decline a permanent delete gets a refusal, not a
deletion.

**6. Least privilege is available.** `outlook auth login --readonly` requests
`Mail.Read` instead of `Mail.ReadWrite` and drops `Mail.Send` entirely, so the
stored refresh token cannot send, delete, or modify mail even if it leaks. The
mode is recorded with the token and survives refreshes, and any mutating
command fails fast against a read-only session. Recommended whenever an agent
only needs to read.

```bash
outlook auth login --readonly
outlook send --to a@example.com --subject S --body B
# This session was signed in with --readonly, so it cannot POST .../sendMail.
```

Credentialed requests are additionally pinned to HTTPS on the Graph host, so a
malformed or hostile `@odata.nextLink` cannot redirect your bearer token to
another server. See [docs/security-review-2026-08-07.md](./docs/security-review-2026-08-07.md)
for the full audit and the two items deliberately left open.

---

## Commands

### Auth
| Command | Purpose |
|---|---|
| `outlook auth login` | Device-code sign-in |
| `outlook auth status` | Signed-in identity and token expiry (token never printed) |
| `outlook auth refresh` | Force a token refresh |
| `outlook auth logout` | Delete the cached token |

### Reading
| Command | Purpose |
|---|---|
| `outlook me` | Your profile and primary address |
| `outlook list [--folder <name\|id>] [--unread-only] [--top N] [--filter <odata>] [--orderby <expr>] [--select <fields>] [--all] [--limit N]` | List messages |
| `outlook get <messageId> [--html]` | Full message, plain text by default |
| `outlook search <kql> [--folder <name\|id>] [--limit N]` | KQL search across the mailbox |
| `outlook folders list` | Mail folders with unread/total counts |
| `outlook attachments list <messageId>` | Attachments on a message |
| `outlook attachments download <messageId> <attachmentId> --out <path> [--force]` | Save an attachment |

Folder arguments accept the well-known names — `inbox`, `drafts`, `sentitems`,
`deleteditems`, `archive`, `junkemail`, and others — or a raw folder id.

### Sending
| Command | Purpose |
|---|---|
| `outlook send --to <addr> --subject <s> [--body <text> \| --body-file <path>] [--cc] [--bcc] [--html] [--attach <path>...] [--no-save-sent]` | Send a message |
| `outlook reply <messageId> --body <text>` | Reply to the sender |
| `outlook reply-all <messageId> --body <text>` | Reply to everyone |
| `outlook forward <messageId> --to <addr> [--body <text>]` | Forward |

`--to`, `--cc`, `--bcc` and `--attach` are repeatable, and address flags also
accept comma-separated lists.

### Drafts
| Command | Purpose |
|---|---|
| `outlook draft list` | List drafts |
| `outlook draft create --subject <s> [--to ...] [--body ...] [--attach ...]` | Create a draft |
| `outlook draft update <draftId> [--subject] [--body] [--to]` | Edit a draft |
| `outlook draft send <draftId>` | Send an existing draft |

Drafts are the safest way to send: compose, inspect, then send as a separate act.

### Organizing
| Command | Purpose |
|---|---|
| `outlook delete <messageId>` | Move to Deleted Items (**recoverable**) |
| `outlook purge <messageId> --yes` | Hard delete (**may be unrecoverable**) |
| `outlook move <messageId> --folder <name\|id>` | Move to a folder |
| `outlook mark-read <messageId>` / `mark-unread` | Set read state |
| `outlook flag <messageId>` / `unflag` | Set follow-up flag |
| `outlook folders create <displayName> [--parent <id>]` | Create a folder |

### Global flags
| Flag | Effect |
|---|---|
| `--json` | Machine-readable output on every command |
| `--dry-run` | Render the request, send nothing (mutating commands) |
| `--all` / `--limit N` | Follow pagination to exhaustion / stop after N |

---

## Notes and limits

- **Attachments are inline only, under 3 MB.** Larger files need a Graph upload
  session, which is not implemented; the CLI fails loudly with the actual file
  size rather than truncating.
- **Message ids change when a message moves between folders.** All reads send
  `Prefer: IdType="ImmutableId"` so stored ids stay stable, but an id captured
  before a `move` should not be reused after it.
- **`$search` and `$orderby` are not combined.** Microsoft's docs do not confirm
  that pairing for messages, so `search` omits ordering rather than emitting a
  request that may be rejected.
- **`$orderby` must agree with `$filter`.** Graph returns `InefficientFilter`
  when ordering references properties the filter does not; the CLI surfaces that
  with an explanation.

## Development

```bash
pnpm test          # file-size gate + vitest
pnpm typecheck     # tsc --noEmit
pnpm build         # tsup -> dist/cli.js
pnpm test:coverage
```

Tests never touch the network and never send mail: the HTTP layer is injected,
and command tests run against a real client whose `fetch` throws, so any
unguarded network call fails the suite instead of quietly passing.

**No source file may exceed 350 lines** — enforced by
`scripts/check-file-size.sh`, wired into `pnpm test`. At the ceiling, extract a
concern into a new file rather than compressing an existing one.

API reference notes gathered during the build live in `docs/`.
