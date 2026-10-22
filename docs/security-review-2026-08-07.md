# Security Review — outlook-cli @ fcb309b

**Source:** external review by Charlie Garcia's Hermes agent, received 2026-08-07 08:00
(Gmail msg `19fdc185b971b2ad`, subject "see the corrections it wants").

**Verification:** all 7 findings independently re-derived against source by four parallel
read-only agents, 2026-08-07. **7 / 7 CONFIRMED.** Every cited line range contained exactly
what the reviewer described. Reviewer's test-count claim (304 total) matches this repo's suite
exactly (304 pass on macOS; their 2 failures are the Windows-only 0600 assertions).

Reviewer's own framing is worth preserving: *"It is not malware, and its basic Graph
implementation appears functional... But the flag-parser defect alone is enough to make it
unsuitable for unattended Hermes use: one missing hyphen could change 'show me what would be
sent' into 'send this now.'"*

---

## Findings

| # | Finding | Severity | Evidence |
|---|---------|----------|----------|
| 2 | `--yes=false` still purges (permanent delete) | **CRITICAL** | `args.ts:166-170`, `move.ts:84` |
| 1 | Typo'd `--dryrun` silently sends for real | **HIGH** | `args.ts:71-95`, `cli.ts:166` |
| 5 | Bearer token sent to ANY absolute URL, incl. `http://` | **HIGH** | `client.ts:34,145-147,108-118`, `paginate.ts:33,39` |
| 6b | `attachments download --dry-run` writes files | **HIGH** | `attachments.ts:164,107-119` |
| 6d | `attachments list --json` leaks full base64 bodies | **HIGH** | `attachments.ts:67-71`, `command.ts:70-76` |
| 3 | No read-only scope mode; always Mail.ReadWrite+Mail.Send | **MEDIUM-HIGH** | `types.ts:208-214`, `auth.ts:82,129` |
| 4 | Tokens plaintext JSON; chmod is a no-op on Windows; non-atomic write | **MEDIUM-HIGH** | `tokenStore.ts:38-44` |
| 6c | JSON dry-run applies zero redaction (human path applies partial) | **MEDIUM** | `command.ts:52,64-67` vs `dryRun.ts:22` |
| 7 | Windows filename hazards unhandled (CON/NUL, `:` ADS, trailing `.`/space) | **MEDIUM** | `attachments.ts:79-89` |
| 6a | `auth logout --dry-run` deletes local token | **LOW** (documented intent) | `auth.ts:122,87-90` |

### 2 — `--yes=false` still deletes (CRITICAL)

`args.ts:166-170` coerces with `list[list.length - 1] !== false`. A `--yes=false` records the
**string** `'false'` (via the `eq` branch, `args.ts:75-76`), and `'false' !== false` is `true`.
Same for `--yes=0`, `--yes=""`. Only `--no-yes` or omission yields real `false`.

A cautious user typing `--yes=false` to *decline* instead **confirms an unrecoverable delete.**
`move.test.ts` covers bare `--yes` and absent `--yes`, but has no test for `--yes=false`.

### 1 — Typo turns preview into send (HIGH)

`args.ts:71-95` records any `--xxx` under whatever key the user typed; no allowlist exists.
`cli.ts:166` reads only the exact key `'dry-run'`. So `--dryrun` → `dryRun: false` →
`guardMutation` skipped → `send.ts:52` fires a real `POST /me/sendMail`.

Systemic: every mutating command shares the single `ctx.dryRun` from `cli.ts:166`.
*Nuance:* reviewer's "commands never validate which flags are allowed" is slightly overstated —
`folders.ts:57-63` has a one-off `--parent` rejection. It is a named special case, not a
mechanism, and does nothing for `--dry-run` spelling.

### 5 — Token exfiltration via nextLink (HIGH)

`ABSOLUTE_URL_RE = /^https?:\/\//i` (`client.ts:34`) accepts any host and **both schemes**.
`buildUrl` (`client.ts:145-147`) passes absolutes straight through; `send` (`client.ts:108-118`)
attaches `Authorization: Bearer` unconditionally. `paginate.ts:33,39` feeds `@odata.nextLink`
from the response body directly in, guarded only by a loop-check.

`GRAPH_BASE` is never used as an allowlist — only as a prefix for relative paths.
A spoofed nextLink to `http://attacker.example` receives the live token **in cleartext**;
it replays against real Graph (aud matches regardless of presenter). Also: no `redirect:` set
(default `follow`), and no timeout/AbortController anywhere in `client.ts`.

### 3 / 4 — Least privilege and credential storage

`DEFAULT_SCOPES` (`types.ts:208-214`) is hardcoded at both call sites (`auth.ts:82`, `auth.ts:129`).
`deviceCode.ts` accepts a `scopes` param but is never called with anything else — a seam, not a
feature. `Mail.Read` appears only in docs and one test fixture.

`tokenStore.ts:38-44` writes plaintext JSON directly to the destination path, then chmods 0600.
Zero platform branching exists in `src/` (grep for `win32` → nothing), so the 0600 guarantee is
POSIX-only; the code comment claiming "0600 either way" overclaims. No tmpfile+rename, no lock.
`docs/00-auth-and-fundamentals.md:496` already recommends an OS keychain — never implemented.

### 7 — Windows filename hazards

`safeDestination` (`attachments.ts:79-89`) defends `../` traversal well (two independent layers:
`basename()` plus a `resolve`/`relative`/`isAbsolute` recheck). It does **not** handle reserved
device names, NTFS ADS colons, trailing dots/spaces, or case-collisions. 1 of 8 hazards is tested.

---

## Recommended fix order

**Tier 1 — unambiguous safety bugs, no design tradeoff (do first):**
1. `#2` strict boolean coercion for confirmation flags — reject non-empty values, or accept only
   bare `--yes`. Add tests for `--yes=false|0|""`.
2. `#1` unknown-flag rejection: per-command allowlist, fail closed on unrecognized flags.
   This is the single highest-leverage fix — it is what makes the CLI safe for unattended use.
3. `#5` restrict authenticated absolute URLs to `https://` on the Graph host; set
   `redirect: 'error'` (or manual validation) and add a request timeout.
4. `#6b` route `attachments download` through `guardMutation` (it writes to disk).
5. `#6c`/`#6d` apply `redactObject` on the JSON path; strip `contentBytes` from
   `attachments list --json`.
6. `#7` extend `safeDestination` with Windows reserved-name/ADS/trailing-char sanitization.

Tier 1 is additive and backward-compatible except `#1` and `#2`, which intentionally turn
previously-silent acceptance into a loud error — that is the point.
Headroom against the 350-LOC gate: `args.ts` 188 (162 free), `attachments.ts` 167 (183 free).

**Tier 2 — real design decisions, need a call:**
7. `#3` read-only auth mode. Needs a decision: separate token bucket per scope set, or a
   `--readonly` login flag that stores a distinct file? Affects existing tokens.
8. `#4` Windows-protected credential storage. Conflicts with the zero-runtime-dependency
   property (package.json has `"dependencies": {}`, and that is a deliberate selling point).
   Options: DPAPI via a native dep, delegate to an OS keychain, or accept + document the
   limitation and at minimum make the write atomic (tmpfile+rename) and warn on Windows.

**Tier 3 — accept as designed:**
9. `#6a` `auth logout --dry-run` deleting the local token is documented intent
   (`auth.ts:12-18`). Worth a doc note, not a code change.

---

## Resolution status

All of Tier 1 plus the read-only auth mode are **fixed**. Test count went 304 → 390.

| # | Status | Where |
|---|--------|-------|
| 1 | FIXED — per-command allowlist, fail closed, with a "did you mean" hint | `flagSpec.ts` (new), `cli.ts` |
| 2 | FIXED — `bool()` accepts only explicit boolean spellings | `args.ts` |
| 3 | FIXED — `auth login --readonly`, mode persisted and preserved across refresh | `types.ts`, `auth.ts`, `command.ts` |
| 4 | PARTIAL — atomic write + Windows warning; encrypted storage deferred (see below) | `tokenStore.ts` |
| 5 | FIXED — HTTPS + Graph-host pin at the point the bearer is attached; `redirect: 'error'`; 60s timeout | `client.ts` |
| 6a | ACCEPTED as designed — documented local-only behavior | `commands/auth.ts` |
| 6b | FIXED — dry-run returns before `mkdir`/`writeFile` | `commands/attachments.ts` |
| 6c | FIXED — JSON dry-run now redacts like the human path | `command.ts` |
| 6d | FIXED — `contentBytes` stripped from `attachments list --json` | `commands/attachments.ts` |
| 7 | FIXED — reserved names, ADS colons, illegal chars, trailing dot/space | `safeFilename.ts` (new) |

### Deliberately not done

**#4 (remainder) — encrypted / OS-keychain credential storage.** Would require a
native dependency (DPAPI, keytar). This package has `"dependencies": {}`, and
that is a deliberate property, not an accident: zero runtime deps is what makes
it auditable and safe to run unattended. Trading that for Windows-only
at-rest encryption is a bad exchange, and the mitigation available without it
is now in place — the write is atomic, the temp file is created 0600 so the
secret is never briefly world-readable, and `auth login` warns on Windows that
file mode is not enforced there. The stronger answer for an exposed host is
`--readonly` (#3), which shrinks what a stolen token is *worth* rather than
trying to hide it. Revisit if a Windows user actually needs it.

**#6a — `auth logout --dry-run` still clears the local token.** Documented
intent (`commands/auth.ts:12-18`): `--dry-run` protects the *mailbox* from a
surprising Graph call. Logout touches only a local file and is trivially
recoverable by signing in again. Changing it would make `--dry-run` mean two
different things depending on the command.

### Regression found during the fix

Implementing #3 introduced a privilege-escalation bug that the review did not
contain: `toTokenSet` did not carry `mode` forward, so a `--readonly` grant
lost its mode on the first refresh and silently re-escalated to full
read/write on the second. Caught by an independent verification pass, fixed in
`auth.ts`, and covered by a compounding double-refresh test in
`auth.grantMode.test.ts`.

## Notes

- The installed copy is inert: no `GRAPH_CLIENT_ID`, no token — it cannot reach a mailbox today.
