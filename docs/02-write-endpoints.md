# Microsoft Graph v1.0 — Mail Write/Send Endpoints Reference

Implementation-ready reference for a Node.js CLI. Every claim below is sourced from a live fetch
of the cited Microsoft Learn page on 2026-08-05. Where a doc is silent or internally inconsistent,
that gap is called out explicitly under **⚠ DOC GAP** rather than filled by inference — except in
§4 and §8, where structural (HTTP-semantics) reasoning is added and clearly labeled as such, because
the CLI's safety defaults depend on it.

Base URL for all paths below: `https://graph.microsoft.com/v1.0`

---

## Table of contents

1. [sendMail](#1-sendmail)
2. [reply / replyAll / forward (one-shot) vs createReply / createReplyAll / createForward (draft)](#2-reply--replyall--forward-vs-createreply--createreplyall--createforward)
3. [Drafts: create, update, send](#3-drafts-create-update-send)
4. [delete vs move — recoverability verdict](#4-delete-vs-move--recoverability-verdict)
5. [Flags & read state](#5-flags--read-state)
6. [Attachments: inline + upload session](#6-attachments-inline--upload-session)
7. [Folders](#7-folders)
8. [Status codes & retry-safety matrix](#8-status-codes--retry-safety-matrix)
9. [Source citations](#9-source-citations)

---

## 1. sendMail

**Source:** https://learn.microsoft.com/en-us/graph/api/user-sendmail

```
POST /me/sendMail
POST /users/{id | userPrincipalName}/sendMail
```

### Exact envelope (verbatim from doc)

```json
{
  "message": {
    "subject": "Meet for lunch?",
    "body": {
      "contentType": "Text",
      "content": "The new cafeteria is open."
    },
    "toRecipients": [
      { "emailAddress": { "address": "frannis@contoso.com" } }
    ],
    "ccRecipients": [
      { "emailAddress": { "address": "danas@contoso.com" } }
    ]
  },
  "saveToSentItems": "false"
}
```

- Top-level envelope is exactly `{ message: {...}, saveToSentItems: <bool> }`. No other top-level keys.
- `saveToSentItems` — doc's parameter table types it **Boolean**, default `true`, "Specify it only if the parameter is false." **⚠ DOC GAP / quirk:** the doc's own worked example sends it as the **quoted string** `"false"`, not a bare JSON boolean. Graph evidently tolerates the string form (the example's response is `202 Accepted`), but a spec-correct client should send a real boolean (`false`, unquoted) and not rely on string coercion.
- Recipients: `toRecipients` / `ccRecipients` / `bccRecipients` are each an array of `{ "emailAddress": { "address": "...", "name": "..." } }`. `name` is optional (the doc's own sendMail examples omit it and use `address` only); other pages (e.g. message-reply) show `name` populated alongside `address`.
- `bccRecipients` is not shown in the sendMail examples but exists identically on the `message` resource — same array-of-`emailAddress` shape.

### `body.contentType` — exact casing

This needs two data points read together:
- The **bodyType** resource schema (https://learn.microsoft.com/en-us/graph/api/resources/itembody) states: *"Possible values are `text` and `html`"* — **lowercase**.
- Every sendMail / create-draft **request** example on Microsoft's own pages uses **`"Text"`** and **`"HTML"`** (Title case / all-caps) — and both succeed (`202 Accepted` / `201 Created` in the docs' own responses).
- Every **response** body Microsoft shows (the actual stored/echoed message) normalizes to **lowercase** — `"contentType":"html"`, `"contentType":"text"` (seen in message-move, message-createreply, user-post-messages response examples).

**Practical rule for the CLI:** the field is case-insensitive on write (Graph accepts `Text`/`HTML`/`text`/`html`), but always reads back lowercase. To match the documented enum exactly, **send lowercase `"text"` / `"html"`** in code you write yourself; if you copy Microsoft's own sample payloads verbatim you'll see `"Text"`/`"HTML"` and that is also documented-working. Do not emit any other casing variant (`"HTML "`, `"Html"`, etc.) — only the two forms above appear anywhere in the fetched docs.

### Success response

> **"If successful, this method returns `202 Accepted` response code. It doesn't return anything in the response body."** (verbatim)

Confirmed exactly as stated in the task: **202, empty body.** The doc adds an important operational caveat, verbatim:

> "A 202 Accepted response code indicates that the request has been accepted; however, it doesn't indicate that the request processing has completed. Delivery of the message is subject to Exchange Online limitations and throttling."

### Errors

- Malformed MIME body → `400 Bad Request`, body: `{"error":{"code":"ErrorMimeContentInvalidBase64String","message":"Invalid base64 string for MIME content."}}`

### Permissions

`Mail.Send` — identical for Delegated (work/school), Delegated (personal Microsoft account), and Application. No higher-privileged alternative listed ("Not available" in all three rows).

### curl

```bash
curl -X POST 'https://graph.microsoft.com/v1.0/me/sendMail' \
  -H "Authorization: Bearer $TOKEN" \
  -H 'Content-Type: application/json' \
  -d '{
    "message": {
      "subject": "Meet for lunch?",
      "body": { "contentType": "text", "content": "The new cafeteria is open." },
      "toRecipients": [{ "emailAddress": { "address": "frannis@contoso.com" } }]
    },
    "saveToSentItems": false
  }'
# → HTTP/1.1 202 Accepted   (empty body)
```

Attachments can ride along in the same call — `message.attachments[]`, same shape as §6:

```json
{
  "message": {
    "subject": "Meet for lunch?",
    "body": { "contentType": "text", "content": "The new cafeteria is open." },
    "toRecipients": [{ "emailAddress": { "address": "meganb@contoso.com" } }],
    "attachments": [
      {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": "attachment.txt",
        "contentType": "text/plain",
        "contentBytes": "SGVsbG8gV29ybGQh"
      }
    ]
  }
}
```

---

## 2. reply / replyAll / forward vs createReply / createReplyAll / createForward

Two independent families. **Send-now** family sends immediately and returns 202/no-body. **Draft**
family creates an editable draft (201 + full Message) that you later hit with `POST .../send`.

Doc's own framing of the distinction (verbatim, from message-reply):

> "This method saves the message in the **Sent Items** folder. Alternatively, [create a draft to reply to an existing message](...) and [send](...) it later."

And from message-createreply (the mirror statement):

> "[Send](...) the draft message in a subsequent operation. Alternatively, [reply to a message](...) in a single operation."

### 2a. Send-now: reply / replyAll / forward

| Op | Path | Body params | Recipients required? |
|---|---|---|---|
| reply | `POST /me/messages/{id}/reply` | `comment` (string), `message` (Message object) | no — replies to original sender/replyTo automatically |
| replyAll | `POST /me/messages/{id}/replyAll` | `comment` (string) only — see ⚠ below | no — loads sender + all original recipients automatically |
| forward | `POST /me/messages/{id}/forward` | `comment` (string), `toRecipients` (Recipient[]) | **yes** — see rule below |

All three: **`202 Accepted`, no response body.** Malformed MIME → `400 Bad Request`,
`"Invalid base64 string for MIME content."` Permission: `Mail.Send` (delegated work/school,
delegated personal, application — identical to sendMail).

**`comment` vs `message` — exact doc rule** (reply page, verbatim):

> "Specify either a comment or the **body** property of the `message` parameter. Specifying both will return an HTTP `400 Bad Request` error."

This means `comment` + a `message` object **can coexist** — the conflict is specifically between
`comment` and `message.body`. You may legally send `comment` + `message.toRecipients` together, e.g.
(verbatim reply example):

```json
{
  "message": {
    "toRecipients": [
      { "emailAddress": { "address": "samanthab@contoso.com", "name": "Samantha Booth" } },
      { "emailAddress": { "address": "randiw@contoso.com", "name": "Randi Welch" } }
    ]
  },
  "comment": "Samantha, Randi, would you name the group please?"
}
```

**⚠ DOC GAP:** the `message-replyall` page's own parameter table lists **only** `comment` — no
`message` row — even though its prose (quoted above) references "the body property of the message
parameter" as if one exists, mirroring reply/forward. This is an inconsistency in Microsoft's own
docs, not a confirmed absence. Treat replyAll's support for an accompanying `message` object as
**unconfirmed** by this page; if your CLI needs to add recipients on a reply-all, prefer
`createReplyAll` (draft) → PATCH → send, which has an unambiguous, fully-documented Message body.

**forward's recipient rule is stricter and mandatory** (verbatim):

> "Specify either the `toRecipients` parameter or the **toRecipients** property of the `message`
> parameter. Specifying both or specifying neither will return an HTTP 400 Bad Request error."

Forward example (verbatim):

```json
{
  "comment": "comment-value",
  "toRecipients": [
    { "emailAddress": { "name": "name-value", "address": "address-value" } }
  ]
}
```

curl:

```bash
curl -X POST "https://graph.microsoft.com/v1.0/me/messages/$MSG_ID/reply" \
  -H "Authorization: Bearer $TOKEN" -H 'Content-Type: application/json' \
  -d '{"comment": "Sounds good, see you there."}'
# → 202 Accepted, empty body

curl -X POST "https://graph.microsoft.com/v1.0/me/messages/$MSG_ID/forward" \
  -H "Authorization: Bearer $TOKEN" -H 'Content-Type: application/json' \
  -d '{"comment": "FYI", "toRecipients": [{"emailAddress": {"address": "danas@contoso.com"}}]}'
# → 202 Accepted, empty body
```

### 2b. Draft-based: createReply / createReplyAll / createForward

| Op | Path | Request body |
|---|---|---|
| createReply | `POST /me/messages/{id}/createReply` | none required (JSON mode) |
| createReplyAll | `POST /me/messages/{id}/createReplyAll` | none required (JSON mode) |
| createForward | `POST /me/messages/{id}/createForward` | none required (JSON mode); optionally `{message:{...}, comment:"..."}` to prefill, e.g. `message.toRecipients` and `message.isDeliveryReceiptRequested` |

All three: verbatim — **"This method doesn't require a request body"** for the plain JSON call
(you can call it with an empty POST). All three return **`201 Created`** with a full **Message**
object in the body — the new draft, with `"isDraft": true` and a fresh `id`.

**⚠ DOC GAP (templating artifact):** each of these three pages *also* shows a second, generic
example response block that says `HTTP/1.1 200 OK` — this is inconsistent with the concrete
worked MIME example on the same page, which correctly shows `HTTP/1.1 201 Created` plus the full
Message payload. Treat `201 Created` as authoritative — it's confirmed by the real example, prose
("this method returns `201 Created` response code and Message object"), and downstream consistency
with `user-post-messages`/`message-send`. The `200 OK` block is very likely a copy-paste template
bug in Microsoft's docs.

createForward accepts an optional body to prefill the draft (verbatim example):

```json
{
  "message": {
    "isDeliveryReceiptRequested": true,
    "toRecipients": [
      { "emailAddress": { "address": "danas@contoso.com", "name": "Dana Swope" } }
    ]
  },
  "comment": "Dana, just want to make sure you get this; you'll need this if the project gets approved."
}
```

Once you have the draft `id` from the 201 response, edit it with `PATCH /me/messages/{id}` (§3) and
send it with `POST /me/messages/{id}/send` (§3) — never re-call `createReply` to "finish" it.

Permissions for the whole create* family: `Mail.ReadWrite` (not `Mail.Send` — creating a draft
doesn't send anything).

curl:

```bash
curl -X POST "https://graph.microsoft.com/v1.0/me/messages/$MSG_ID/createReply" \
  -H "Authorization: Bearer $TOKEN"
# → 201 Created + full Message JSON (isDraft: true)
```

---

## 3. Drafts: create, update, send

### Create — `POST /me/messages`

**Source:** https://learn.microsoft.com/en-us/graph/api/user-post-messages

```
POST /me/messages
POST /users/{id|userPrincipalName}/messages
```

Body = a JSON `message` object (verbatim example):

```json
{
  "subject": "Did you see last night's game?",
  "importance": "Low",
  "body": { "contentType": "HTML", "content": "They were <b>awesome</b>!" },
  "toRecipients": [{ "emailAddress": { "address": "AdeleV@contoso.com" } }]
}
```

- **"By default, this operation saves the draft in the Drafts folder."** (verbatim)
- Success: **`201 Created`** + full Message object in body (consistent across all 3 worked examples on the page — no `200 OK` inconsistency here, unlike §2b).
- S/MIME note (verbatim): **"S/MIME message payloads are currently limited to 4 MB. Submission attempts that exceed this limit will result in an `HTTP 413 Request Entity Too Large` error response."**
- Permission: `Mail.ReadWrite`.

### Update — `PATCH /me/messages/{id}`

**Source:** https://learn.microsoft.com/en-us/graph/api/message-update

```
PATCH /me/messages/{id}
```

> "In the request body, supply the values for relevant fields that should be updated. Existing
> properties that are not included in the request body will maintain their previous values...
> For best performance you shouldn't include existing values that haven't changed."

Success: **`200 OK`** + updated Message object in body.

**Writable properties (doc's own positive list — this is the complete table on the page, verbatim conditions):**

| Property | Type | Condition |
|---|---|---|
| `bccRecipients` | Recipient collection | Updatable only if `isDraft = true` |
| `body` | ItemBody | Updatable only if `isDraft = true` |
| `categories` | String collection | none |
| `ccRecipients` | Recipient collection | Updatable only if `isDraft = true` |
| `flag` | followupFlag | none |
| `from` | Recipient | "Must correspond to the actual mailbox used" |
| `importance` | String (`Low`/`Normal`/`High`) | none |
| `inferenceClassification` | String (`focused`/`other`) | none |
| `internetMessageId` | String | Updatable only if `isDraft = true` |
| `isDeliveryReceiptRequested` | Boolean | none |
| `isRead` | Boolean | none |
| `isReadReceiptRequested` | Boolean | none |
| `multiValueExtendedProperties` | collection | Updatable only if `isDraft = true` |
| `replyTo` | Recipient collection | Updatable only if `isDraft = true` |
| `sender` | Recipient | shared-mailbox/delegate scenario; must match actual mailbox |
| `singleValueExtendedProperties` | collection | Updatable only if `isDraft = true` |
| `subject` | String | Updatable only if `isDraft = true` |
| `toRecipients` | Recipient collection | Updatable only if `isDraft = true` |

**⚠ DOC GAP:** the doc never publishes an explicit "read-only properties" list — it only gives the
positive "can be updated" table above. Anything not in that table (`id`, `createdDateTime`,
`changeKey`, `hasAttachments`, `receivedDateTime`, `sentDateTime`, `webLink`, `conversationId`,
`parentFolderId`, etc.) is not documented as updatable via PATCH; treating it as read-only is an
inference from omission, not an explicit doc statement.

**The single most important practical rule buried in this table:** almost every content-shaping
field — `subject`, `body`, `toRecipients`, `ccRecipients`, `bccRecipients`, `replyTo`,
`internetMessageId` — is **only updatable while `isDraft = true`**. Once a message is sent
(`isDraft = false`), your CLI can still PATCH `categories`, `flag`, `isRead`,
`isDeliveryReceiptRequested`, `isReadReceiptRequested`, `inferenceClassification` — but not the
content or recipients.

Generic PATCH example (verbatim):

```bash
curl -X PATCH "https://graph.microsoft.com/v1.0/me/messages/$MSG_ID" \
  -H "Authorization: Bearer $TOKEN" -H 'Content-Type: application/json' \
  -d '{
    "subject": "subject-value",
    "body": { "contentType": "text", "content": "content-value" },
    "inferenceClassification": "other"
  }'
# → 200 OK + updated Message JSON
```

(See §5 for the exact `isRead` and `flag` PATCH bodies — this page's own examples don't show those
two specific snippets, so they're composed from the property table + the cross-referenced
`followupFlag` resource page, not copy-pasted from a single example on this page.)

### Send an existing draft — `POST /me/messages/{id}/send`

**Source:** https://learn.microsoft.com/en-us/graph/api/message-send

```
POST /me/messages/{id}/send
```

> "Since this method sends an already existing draft message, specifying a request body is not
> necessary." Header table requires `Content-Length: 0`.

Success: **`202 Accepted`, no response body** — confirmed exactly as stated in the task. Applies
to a new-message draft, a reply draft, a reply-all draft, or a forward draft (all four are
interchangeable inputs to this one endpoint). Saves the message in Sent Items. Permission:
`Mail.Send`.

```bash
curl -X POST "https://graph.microsoft.com/v1.0/me/messages/$DRAFT_ID/send" \
  -H "Authorization: Bearer $TOKEN" -H 'Content-Length: 0'
# → 202 Accepted, empty body
```

---

## 4. delete vs move — recoverability verdict

**This is the safety-critical section. Read it in full before wiring a "delete" command.**

### `DELETE /me/messages/{id}`

**Source:** https://learn.microsoft.com/en-us/graph/api/message-delete

```
DELETE /me/messages/{id}
```

- No request body ("Don't supply a request body for this method.")
- Success: **`204 No Content`**, no response body.
- Permission: `Mail.ReadWrite`.
- **Recoverability — exact doc text, in full, is only this:** *"Delete a message in the specified
  user's mailbox, or delete a relationship of the message."* Plus one Note, which is about a
  **different** scenario:

  > "You may not be able to delete items in the recoverable items deletions folder (represented by
  > the well-known folder name `recoverableitemsdeletions`)."

  That note governs deleting an item that is **already inside** the hidden
  Recoverable-Items-Deletions folder — it says nothing about what happens on an ordinary first
  `DELETE` of a message sitting in Inbox or any normal folder.

**⚠ DOC GAP — this is the crux of the safety question.** Microsoft's `message-delete` page **never
states**, anywhere, whether a first-pass `DELETE /me/messages/{id}` permanently purges the message
or soft-deletes it into a recoverable state (e.g. moves it to Deleted Items, or straight to the
hidden Recoverable Items\Deletions store, bypassing the visible Deleted Items folder a user could
casually undo from). Do not treat this as settled by general Exchange/Outlook folk-knowledge about
multi-stage retention — that knowledge, even if true, is **not present in this official page**, so
it must not be silently assumed by the CLI's safety logic.

### `POST /me/messages/{id}/move`

**Source:** https://learn.microsoft.com/en-us/graph/api/message-move

```
POST /me/messages/{id}/move
```

Doc's own description (verbatim, first line of the page):

> **"Move a message to another folder within the specified user's mailbox. This creates a new copy
> of the message in the destination folder and removes the original message."**

Body: single property `destinationId` — **"The destination folder ID, or a well-known folder
name"** (verbatim). Confirmed exact example, moving to Deleted Items:

```json
{ "destinationId": "deleteditems" }
```

with the doc's own sentence: *"The following request moves the specified message to the Deleted
Items folder, identified by its well-known folder name `deleteditems`."* — **no folder-id
resolution step is required**; the literal string `"deleteditems"` is valid.

- Success: **`201 Created`** + a full `message` resource in the body.
- The response's `id` is **provably different** from the request's original id (doc's own example:
  request id `AAMkADhAAATs28OAAA=` → response id `AAMkADhAAAW-VPeAAA=`) — direct evidence that move
  mints a new object and the old id no longer resolves. This is a structural fact from the example,
  not an inference.
- Permission: `Mail.ReadWrite`.

### Verdict

| | `DELETE /me/messages/{id}` | `POST /me/messages/{id}/move` → `deleteditems` |
|---|---|---|
| Destination | Undocumented | Explicitly the named, visible **"Deleted Items"** folder |
| Recoverability | **Not stated by the doc** | Lands in the same folder a user browses/empties in the Outlook UI — structurally the safer, inspectable landing spot |
| Response evidence | 204, no body — nothing to inspect | 201 + full Message resource confirming exactly where it landed and under what new id |
| Reversible by the CLI/user | Unknown / unconfirmed | Yes — move it back with a second `move` call to the original `parentFolderId` |

**CLI default: use `move` → `{"destinationId": "deleteditems"}` as the "delete" verb exposed to
users.** Reserve raw `DELETE /me/messages/{id}` for an explicit, separately-named "purge" /
"hard-delete" command that requires extra confirmation — because its recoverability is
**not documented** by Microsoft, and an undocumented irreversible-or-not operation must be treated
as irreversible by default. This verdict follows directly from what the two pages do and do not
say, not from outside assumptions about Exchange retention internals.

curl (safe default):

```bash
curl -X POST "https://graph.microsoft.com/v1.0/me/messages/$MSG_ID/move" \
  -H "Authorization: Bearer $TOKEN" -H 'Content-Type: application/json' \
  -d '{"destinationId": "deleteditems"}'
# → 201 Created + moved Message resource (note the new "id")
```

curl (hard delete — gate behind explicit confirmation):

```bash
curl -X DELETE "https://graph.microsoft.com/v1.0/me/messages/$MSG_ID" \
  -H "Authorization: Bearer $TOKEN"
# → 204 No Content, empty body — recoverability NOT confirmed by Microsoft's own docs
```

---

## 5. Flags & read state

Both are `PATCH /me/messages/{id}` (§3) bodies — the message-update page lists `isRead` and `flag`
as writable properties but does not show either JSON snippet as a worked example on that page.
The exact allowed values below come from `message-update`'s property table combined with the
`followupFlag` resource page (https://learn.microsoft.com/en-us/graph/api/resources/followupflag).

### Read state

```json
{ "isRead": true }
```
```json
{ "isRead": false }
```

`isRead` is typed `Boolean` with no update condition (works whether or not the message is a draft).

### Flag

`followupFlag` resource, verbatim: **"flagStatus ... The status for follow-up for an item.
Possible values are `notFlagged`, `complete`, and `flagged`."** — exact spelling and casing,
all lowercase/camelCase, no capitalized variants documented anywhere.

```json
{ "flag": { "flagStatus": "flagged" } }
```
```json
{ "flag": { "flagStatus": "notFlagged" } }
```
```json
{ "flag": { "flagStatus": "complete" } }
```

Bonus, from the `followupFlag` resource page (verbatim): **"To set the due date, you must also
specify the `startDateTime`; otherwise, you get a `400 Bad Request` response."** — i.e. `dueDateTime`
has a hard dependency on `startDateTime`. Full shape when scheduling a follow-up (seen in a
sendMail example that sets a flag on send):

```json
{
  "flag": {
    "flagStatus": "flagged",
    "startDateTime": { "dateTime": "2023-08-30T12:13:00", "timeZone": "Eastern Standard Time" },
    "dueDateTime":   { "dateTime": "2023-09-01T17:00:00", "timeZone": "Eastern Standard Time" }
  }
}
```

curl:

```bash
curl -X PATCH "https://graph.microsoft.com/v1.0/me/messages/$MSG_ID" \
  -H "Authorization: Bearer $TOKEN" -H 'Content-Type: application/json' \
  -d '{"isRead": true}'
# → 200 OK + updated Message

curl -X PATCH "https://graph.microsoft.com/v1.0/me/messages/$MSG_ID" \
  -H "Authorization: Bearer $TOKEN" -H 'Content-Type: application/json' \
  -d '{"flag": {"flagStatus": "flagged"}}'
# → 200 OK + updated Message
```

---

## 6. Attachments: inline + upload session

### Inline (small) — `POST /me/messages/{id}/attachments`

**Source:** https://learn.microsoft.com/en-us/graph/api/message-post-attachments

```
POST /me/messages/{id}/attachments
```

**Exact size limit (verbatim): "This operation limits the size of the attachment you can add to
under 3 MB."** — confirmed as "under 3 MB," not "up to" or "≤".

Doc's own literal request example for a `#microsoft.graph.fileAttachment`:

```json
{
  "@odata.type": "#microsoft.graph.fileAttachment",
  "name": "smile",
  "contentBytes": "R0lGODdhEAYEAA7"
}
```

**⚠ Note on `contentType`:** the doc's own worked example does **not** include `contentType` in
the request — only in the 201 response (`"contentType": "image/gif"`, apparently server-inferred
from content or filename). However, `contentType` **is** a documented, writable property of the
`fileAttachment` resource (https://learn.microsoft.com/en-us/graph/api/resources/fileattachment:
*"contentType | String | The content type of the attachment."*), and other official examples (the
sendMail inline-attachment example, §1) **do** send it explicitly:

```json
{
  "@odata.type": "#microsoft.graph.fileAttachment",
  "name": "attachment.txt",
  "contentType": "text/plain",
  "contentBytes": "SGVsbG8gV29ybGQh"
}
```

**Recommendation: always send `contentType` explicitly** (don't rely on server inference) — it's
documented as legal and is what a real MIME-aware client should know. `contentBytes` is the
base64-encoded raw file content (`fileAttachment` resource, verbatim: *"Make sure to encode the
file content in base64 before assigning it to contentBytes."*).

Success: **`201 Created`** + full `Attachment` object in body (includes server-assigned `id`,
`size`, `isInline`). Permission: `Mail.ReadWrite`.

```bash
curl -X POST "https://graph.microsoft.com/v1.0/me/messages/$MSG_ID/attachments" \
  -H "Authorization: Bearer $TOKEN" -H 'Content-Type: application/json' \
  -d '{
    "@odata.type": "#microsoft.graph.fileAttachment",
    "name": "report.pdf",
    "contentType": "application/pdf",
    "contentBytes": "'"$(base64 -i report.pdf)"'"
  }'
# → 201 Created + Attachment JSON
```

### Large files — upload session (3 MB – 150 MB)

**Source:** https://learn.microsoft.com/en-us/graph/outlook-large-attachments

Overall ceiling (verbatim): **"Using the Microsoft Graph API, you can attach files up to 150 MB to
an Outlook message or event item."** Rule: **under 3 MB → single POST (previous section); between
3 MB and 150 MB → upload session** (this section). Attempting `createUploadSession` for a file
under 3 MB is explicitly rejected — the doc's own troubleshooting note says that scenario returns
an error telling you to use the single-POST method instead.

**1. Create the session:**

```
POST /me/messages/{id}/attachments/createUploadSession
```

```json
{
  "AttachmentItem": {
    "attachmentType": "file",
    "name": "flower",
    "size": 3483322
  }
}
```

**⚠ Note:** this doc's own `AttachmentItem` example has **no `contentType` field** — only
`attachmentType`, `name`, `size`.

Success: **`201 Created`** + `uploadSession` object:

```json
{
  "uploadUrl": "https://outlook.office.com/api/v2.0/Users('...')/Messages('...')/AttachmentSessions('...')?authtoken=...",
  "expirationDateTime": "2019-09-25T01:09:30.7671707Z",
  "nextExpectedRanges": ["0-"]
}
```

Permission: `Mail.ReadWrite` to create the session for a message (`Calendars.ReadWrite` for an event).

**2. Upload chunks — `PUT` to the opaque `uploadUrl`:**

**⚠ CORRECTION vs. common assumption:** this doc does **not** state a "multiple of 320 KiB
(327,680 bytes)" chunk rule anywhere. That rule belongs to the separate OneDrive/SharePoint
`driveItem` upload-session doc — **do not port it here.** The actual rule stated on *this* page is
only a performance recommendation, verbatim:

> "For better performance, keep each byte range less than 4 MB" / "keep the upper limit of the
> number of bytes for each PUT operation to 4 MB."

So: **no hard divisibility requirement** for mail/event attachment upload sessions — just stay
under ~4 MB per chunk. You may even upload the whole file in one PUT if it's under that guideline.

Required headers per chunk:

| Header | Value |
|---|---|
| `Content-Length` | number of bytes in *this* PUT |
| `Content-Range` | **exact format:** `bytes {start}-{end}/{total}` (0-based, inclusive end) |
| `Content-Type` | `application/octet-stream` |

Do **not** send an `Authorization` header on these PUTs — the `uploadUrl` is pre-authenticated.

```
PUT {uploadUrl}
Content-Type: application/octet-stream
Content-Length: 2097152
Content-Range: bytes 0-2097151/3483322

<raw bytes 0..2097151>
```

**Status codes across the chunk sequence — CORRECTED, not the commonly assumed 200/202 split:**

- **Intermediate** (non-final) successful chunk PUT → **`200 OK`** + an `uploadSession` object
  whose `nextExpectedRanges` tells you where to resume, e.g. `"nextExpectedRanges":["2097152"]`.
  Bytes must be uploaded **in order**.
- **Final** chunk PUT (last byte of the file) → **`201 Created`**, with a `Location` response
  header pointing at the new attachment resource — extract the attachment id from that URL.
- **No `202` appears anywhere in this flow.** (This corrects a common but wrong assumption that
  intermediate chunks return 202.)
- `nextExpectedRanges` may list multiple ranges if parts are still missing — this is the documented
  resume mechanism: *"This is useful if you need to resume a transfer that was interrupted and your
  client is unsure of the state on the service."* Re-query/resume from there rather than restarting.
- Cancel an in-progress session: `DELETE {uploadUrl}` (no Authorization header) → `204 No Content`.

```bash
curl -X PUT "$UPLOAD_URL" \
  -H 'Content-Type: application/octet-stream' \
  -H 'Content-Range: bytes 0-2097151/3483322' \
  -H 'Content-Length: 2097152' \
  --data-binary @chunk0.bin
# intermediate → 200 OK + {"nextExpectedRanges": ["2097152"]}

curl -X PUT "$UPLOAD_URL" \
  -H 'Content-Type: application/octet-stream' \
  -H 'Content-Range: bytes 2097152-3483321/3483322' \
  -H 'Content-Length: 1386170' \
  --data-binary @chunk1.bin
# final → 201 Created, Location: https://outlook.office.com/.../Attachments('...')
```

Known issue (flagged by the doc itself, linked not detailed): attaching large files in a
**shared or delegated mailbox** has a documented open issue — check
https://developer.microsoft.com/en-us/graph/known-issues/?search=13644 before relying on this path
against a shared mailbox.

---

## 7. Folders

**Source:** https://learn.microsoft.com/en-us/graph/api/user-post-mailfolders

```
POST /me/mailFolders
```

> "Use this API to create a new mail folder in the root folder of the user's mailbox."

Body (verbatim, and verbatim: **"displayName and isHidden are the only writable property for a
mailFolder object"**):

```json
{
  "displayName": "Clutter",
  "isHidden": true
}
```

- `displayName` (String) — required, no documented uniqueness constraint (see §8 retry note).
- `isHidden` (Boolean) — optional, default `false`. **"Once set, you cannot update this property"**
  (verbatim) — this is a one-way write; there is no un-hide via PATCH once set at creation.
- `parentFolderId` is **not** an accepted request parameter here — folders are always created at
  root; `parentFolderId` only appears read-only in the response.
- Success: **`201 Created`** + full `mailFolder` object.
- Permission: `Mail.ReadWrite`.

```bash
curl -X POST 'https://graph.microsoft.com/v1.0/me/mailFolders' \
  -H "Authorization: Bearer $TOKEN" -H 'Content-Type: application/json' \
  -d '{"displayName": "Clutter", "isHidden": true}'
# → 201 Created + mailFolder JSON
```

---

## 8. Status codes & retry-safety matrix

**None of the 13 pages fetched for this reference use the words "idempotent," "retry," or
"duplicate."** Every retry-safety judgment below is this document's own structural reasoning from
HTTP method semantics and the documented side effects — labeled `[inference]` — not sourced
Microsoft guidance. Do not present these as Microsoft's own claims.

| Call | Method | Success | Body on success | Retry-safe after a timeout? |
|---|---|---|---|---|
| Send mail | `POST /me/sendMail` | 202 | none | **UNSAFE.** No idempotency key, no dedup. A retried call after an ambiguous timeout can send the same mail twice. `[inference]` |
| Reply | `POST /me/messages/{id}/reply` | 202 | none | **UNSAFE** — same send-now risk. `[inference]` |
| Reply all | `POST /me/messages/{id}/replyAll` | 202 | none | **UNSAFE** — same. `[inference]` |
| Forward | `POST /me/messages/{id}/forward` | 202 | none | **UNSAFE** — same. `[inference]` |
| Create draft | `POST /me/messages` | 201 | Message | **UNSAFE** — no uniqueness constraint documented; retry creates a second draft. `[inference]` |
| Create reply/replyAll/forward draft | `POST /me/messages/{id}/create{Reply,ReplyAll,Forward}` | 201 | Message | **UNSAFE** — same duplicate-draft risk. `[inference]` |
| Send existing draft | `POST /me/messages/{id}/send` | 202 | none | **UNSAFE to blindly retry**, but partially self-mitigating: `[inference]` after a successful send the draft is no longer at that id/location, so a naive retry is more likely to 404 than to double-send — **still**, don't rely on that; re-`GET` the message by id first to check `isDraft`/location before retrying. |
| Update (PATCH) | `PATCH /me/messages/{id}` | 200 | Message | **SAFE for the documented field set** (`isRead`, `flag`, `categories`, `importance`, etc.) — every one of these sets an absolute value, not a relative/incremental change, so re-applying the same PATCH converges to the same state. `[inference]` |
| Delete | `DELETE /me/messages/{id}` | 204 | none | **SAFE-ish by HTTP semantics** — DELETE is idempotent by definition (repeat calls converge to "resource absent"); a retry either 204s again or 404s if it already succeeded. Doc doesn't confirm this explicitly, but it follows from HTTP method semantics rather than app-specific side effects. `[inference]` |
| Move | `POST /me/messages/{id}/move` | 201 | Message (new id) | **Loudly self-detecting, not silently unsafe:** a successful move removes the source id and mints a new one (confirmed by doc's own example — different `id` in response), so retrying with the same old id after success will **404**, not double-move. Treat a 404-on-retry as "probably already succeeded — verify by listing the destination folder," not as a hard failure. `[inference]` |
| Create attachment (inline) | `POST /me/messages/{id}/attachments` | 201 | Attachment | **UNSAFE** — retry can attach the same file twice with two different attachment ids. `[inference]` |
| Create upload session | `POST /me/messages/{id}/attachments/createUploadSession` | 201 | uploadSession | Creating a second session on timeout is low-cost (unused sessions simply expire per `expirationDateTime`) but still creates orphaned sessions — prefer checking for an existing session before retrying. `[inference]` |
| Upload chunk | `PUT {uploadUrl}` | 200 (intermediate) / 201 (final) | uploadSession / none+Location | **Designed to be retried** — this is the one call in this whole surface with an explicit, documented resume mechanism (`nextExpectedRanges`). On ambiguous failure, re-`PUT` the same `Content-Range`, or inspect `nextExpectedRanges` to see what the server actually has. Closest thing to a genuinely safe retry in this API surface. |
| Create folder | `POST /me/mailFolders` | 201 | mailFolder | **UNSAFE** — no documented uniqueness constraint on `displayName`; retry creates a duplicate folder. `[inference]` |

**Bottom line for CLI design:** every "send" verb (`sendMail`, `reply`, `replyAll`, `forward`,
`.../send`) and every plain "create" verb (`messages`, `createReply*`, `attachments`,
`mailFolders`) must be treated as **not safely retriable** on network timeout/ambiguity — surface
the ambiguity to the caller (e.g. "send status unknown, check Sent Items before resending") rather
than auto-retrying. The only calls with genuine, documented resume/idempotency characteristics are
the **upload-session PUT chunks** and, by HTTP semantics, **DELETE**.

---

## 9. Source citations

All fetched live on 2026-08-05 from learn.microsoft.com (graph-rest-1.0 / v1.0):

1. https://learn.microsoft.com/en-us/graph/api/user-sendmail
2. https://learn.microsoft.com/en-us/graph/api/message-reply
3. https://learn.microsoft.com/en-us/graph/api/message-replyall
4. https://learn.microsoft.com/en-us/graph/api/message-forward
5. https://learn.microsoft.com/en-us/graph/api/message-createreply
6. https://learn.microsoft.com/en-us/graph/api/message-createreplyall
7. https://learn.microsoft.com/en-us/graph/api/message-createforward
8. https://learn.microsoft.com/en-us/graph/api/user-post-messages
9. https://learn.microsoft.com/en-us/graph/api/message-send
10. https://learn.microsoft.com/en-us/graph/api/message-update
11. https://learn.microsoft.com/en-us/graph/api/message-delete
12. https://learn.microsoft.com/en-us/graph/api/message-move
13. https://learn.microsoft.com/en-us/graph/api/message-post-attachments
14. https://learn.microsoft.com/en-us/graph/outlook-large-attachments
15. https://learn.microsoft.com/en-us/graph/api/user-post-mailfolders
16. https://learn.microsoft.com/en-us/graph/api/resources/followupflag (cross-referenced for exact `flagStatus` enum)
17. https://learn.microsoft.com/en-us/graph/api/resources/itembody (cross-referenced for exact `contentType` enum)
18. https://learn.microsoft.com/en-us/graph/api/resources/fileattachment (cross-referenced for `fileAttachment` property list)

### Everything NOT verifiable from these official docs (flagged, not guessed)

- **`DELETE /me/messages/{id}` recoverability** — the doc never states whether this is a permanent
  purge or a soft-delete. See §4. This drove the CLI default to `move`.
- **`replyAll`'s support for a `message` object** alongside `comment` — the page's prose implies it
  (mirroring reply/forward) but its parameter table only lists `comment`. Unconfirmed either way.
- **Retry-safety / idempotency** — no page uses these words. §8 is entirely this document's own
  structural reasoning from HTTP semantics and documented side effects, explicitly labeled
  `[inference]` throughout, not sourced Microsoft guidance.
- **Whether `contentType` is accepted/required on inline `fileAttachment` POST** — the
  `message-post-attachments` worked example omits it from the request (only shows it in the
  response); the `fileAttachment` resource schema documents it as a writable property, and the
  sendMail page's own example does send it. Net: legal to send, not proven required.
