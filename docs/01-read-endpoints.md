# Microsoft Graph v1.0 — Mail Read Surface Reference

Implementation-ready reference for a Node.js CLI reading Outlook/Exchange mail via Microsoft Graph **v1.0**. Every endpoint, literal string, and constraint below is sourced from the official Microsoft Learn docs cited inline. Anything not confirmed from those docs is flagged explicitly in **§9 Unverified / Not Found**.

Sources fetched (all `learn.microsoft.com`, 2026-08-05):
- `/en-us/graph/api/user-list-messages` (S1)
- `/en-us/graph/api/message-get` (S2)
- `/en-us/graph/api/user-list-mailfolders` (S3)
- `/en-us/graph/api/mailfolder-list-messages` (S4)
- `/en-us/graph/search-concept-messages` (S5)
- `/en-us/graph/query-parameters` (S6)
- `/en-us/graph/paging` (S7)
- `/en-us/graph/api/resources/message` (S8)
- `/en-us/graph/api/message-list-attachments` (S9)
- `/en-us/graph/outlook-immutable-id` (S10)
- Supporting (fetched to satisfy exact-value requirements not fully covered above):
  - `/en-us/graph/api/resources/mailfolder` (S11) — well-known folder name table
  - `/en-us/graph/search-query-parameter` (S12) — KQL searchable-property table, `$search` mechanics
  - `/en-us/graph/filter-query-parameter` (S13) — `$filter` operators/syntax, `isRead eq false` example
  - `/en-us/graph/api/resources/attachment` (S14), `/fileattachment` (S15), `/itemattachment` (S16), `/referenceattachment` (S17) — attachment discriminators
  - `/en-us/graph/known-issues` (S18) — query-parameter limitations, `/attachments` `$filter` note

All requests below assume `Authorization: Bearer {token}` and base URL `https://graph.microsoft.com/v1.0`.

---

## 1. Endpoints (exact paths)

Base resource is `/me` (signed-in user, delegated) or `/users/{id | userPrincipalName}` (app permissions or delegated-with-shared-mailbox access). [S1][S2][S3][S4][S9]

| Purpose | Exact path |
|---|---|
| Current signed-in user | `GET /me` |
| List messages (whole mailbox incl. Deleted Items and Clutter) | `GET /me/messages` <br> `GET /users/{id \| userPrincipalName}/messages` |
| List mail folders (root-level, non-hidden by default) | `GET /me/mailFolders` <br> `GET /users/{id \| userPrincipalName}/mailFolders` |
| List mail folders incl. hidden | `GET /me/mailFolders?includeHiddenFolders=true` |
| List messages in a specific folder | `GET /me/mailFolders/{id}/messages` <br> `GET /users/{id \| userPrincipalName}/mailFolders/{id}/messages` |
| Get one message | `GET /me/messages/{id}` <br> `GET /users/{id \| userPrincipalName}/messages/{id}` <br> `GET /me/mailFolders/{id}/messages/{id}` <br> `GET /users/{id \| userPrincipalName}/mailFolders/{id}/messages/{id}` |
| Get MIME content of a message | `GET /me/messages/{id}/$value` (also under `/users/...` and `/mailFolders/.../messages/{id}/$value`) |
| List attachments on a message | `GET /me/messages/{id}/attachments` <br> `GET /users/{id \| userPrincipalName}/messages/{id}/attachments` <br> `GET /me/mailFolders/{id}/messages/{id}/attachments` (also nested under `childFolders`) |

Note: `/me` itself is the user resource root — not documented in S1–S10 as a mail-specific endpoint, but it is the implicit base for every `/me/...` mail path above and is required for e.g. resolving the signed-in user's `mail`/`userPrincipalName` before other calls. This CLI should treat `GET /me` as its identity-check endpoint.

**Permissions** (least-privileged first) [S1][S2][S9]:
- List/get messages, list folders: delegated `Mail.ReadBasic` (higher: `Mail.ReadWrite`, `Mail.Read`); application `Mail.ReadBasic.All`.
- List attachments: `Mail.Read` (delegated and application) — S9 does not list a `ReadBasic` tier for attachments.

---

## 2. Well-known folder names (exact literal strings)

Usable in place of a folder `{id}`, e.g. `GET /me/mailFolders/inbox/messages`. Work regardless of the mailbox's display-language locale. [S11]

```
archive
clutter
conflicts
conversationhistory
deleteditems
drafts
inbox
junkemail
localfailures
msgfolderroot
outbox
recoverableitemsdeletions
scheduled
searchfolders
sentitems
serverfailures
syncissues
```

Descriptions (verbatim, condensed) [S11]:
- `archive` — where One-Click Archive sends messages (Outlook client feature; **not** the Exchange Online Archive Mailbox feature).
- `clutter` — low-priority messages moved here by the Clutter feature.
- `conflicts` — conflicting items.
- `conversationhistory` — Skype IM conversations, if configured.
- `deleteditems` — items moved here on delete.
- `drafts` — unsent messages.
- `inbox` — the inbox.
- `junkemail` — junk email.
- `localfailures` — items that exist locally but failed to upload to the server.
- `msgfolderroot` — "Top of Information Store," parent of normal mail-client folders (e.g. inbox).
- `outbox` — the outbox.
- `recoverableitemsdeletions` — soft-deleted items (deleted from Deleted Items, or Shift+Delete); not visible in Outlook UI directly, reachable via "Recover Deleted Items from Server."
- `scheduled` — messages scheduled to reappear in the inbox (Outlook for iOS Schedule feature).
- `searchfolders` — parent folder for all search folders.
- `sentitems` — sent items.
- `serverfailures` — items on the server that failed to sync to the local client.
- `syncissues` — Outlook synchronization logs.

**Example:**
```bash
curl -s \
  -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/me/mailFolders/drafts"
```

`GET /me/mailFolders` also returns *non-well-known, actual* folders — including a `mailSearchFolder` subtype (`@odata.type: "#microsoft.graph.mailSearchFolder"`), e.g. a "Weekly digests" search folder child of Inbox with `filterQuery: "contains(subject, 'weekly digest')"`. [S3]

`GET /me/mailFolders` only returns **root-level child folders** — it does not recurse. To enumerate all folders, each child folder's `childFolders` must be traversed separately via `mailfolder-list-childfolders`. [S3]

---

## 3. OData query parameters

Supported by `/me/messages`, `/me/mailFolders/{id}/messages`, `/me/messages/{id}` per S1/S2/S4 ("This method supports the OData Query Parameters"). General mechanics from S6; message-specific constraints from S1/S13.

### 3.1 `$select` — reduce properties, improve latency
```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/me/messages?\$select=sender,subject"
```
Response (from S1, live example):
```json
{
  "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('bb8775a4-4d8c-42cf-a1d4-4d58c2bb668f')/messages(sender,subject)",
  "value": [
    {
      "@odata.etag": "W/\"CQAAABYAAADHcgC8Hl9tRZ/hc1wEUs1TAAAwR4Hg\"",
      "id": "AAMkAGUAAAwTW09AAA=",
      "subject": "You have late tasks!",
      "sender": { "emailAddress": { "name": "Microsoft Planner", "address": "noreply@Planner.Office365.com" } }
    }
  ]
}
```
Requesting without `$select` returns a `@microsoft.graph.tips` hint recommending its use. [S6] `internetMessageHeaders` specifically **requires** `$select` to be returned at all (it's not in the default set). [S8]

### 3.2 `$filter` — subset a collection
Syntax basics: GUID/DateTimeOffset values are **not** quoted; string values use single quotes; a literal single quote inside a string is escaped by doubling it. [S6][S13]
```bash
# Escaped single quote in a literal
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/me/messages?\$filter=subject eq 'let''s meet for lunch?'"
```
Confirmed working examples against messages [S13]:
```
# All unread mail in Inbox
GET /me/mailFolders/inbox/messages?$filter=isRead eq false

# Messages from one sender
GET /me/messages?$filter=from/emailAddress/address eq 'someuser@example.com'

# Date range (April 2017)
GET /me/mailFolders/inbox/messages?$filter=ReceivedDateTime ge 2017-04-01 and receivedDateTime lt 2017-05-01

# Combined subject + importance
GET /me/messages?$filter=Subject eq 'welcome' and importance eq 'normal'&$orderby=subject,importance,receivedDateTime desc
```
Operators Graph supports generally: equality (`eq`, `ne`, `not`, `in`, `has`), relational (`lt`, `gt`, `le`, `ge`), lambda (`any`, `all` — **`all` is currently not supported in Microsoft Graph**), conditional (`and`, `or`), functions (`startswith`, `endswith`, `contains`). [S13] Support varies per resource/property — some combinations (e.g. `ne`, `not`, `endswith` on many Entra ID directory-object properties) require **advanced query** (`ConsistencyLevel: eventual` header) and are marked `**` in S13; that advanced-query gate is a **directory-object** (Entra ID) requirement, not documented as applying to `/me/messages`.

Known limitation explicitly documented [S18]: **the `/attachments` endpoint doesn't support `$filter` — if present, the `$filter` parameter is ignored** (silent no-op, not an error).

### 3.3 `$orderby` — sort order
Default order is ascending; append `asc`/`desc` after a space. [S6]
```
GET /me/messages?$orderby=displayName          # ascending (implicit)
GET /me/messages?$orderby=from/emailAddress/name desc,subject   # multi-property sort; complex-type field supported
```
`receivedDateTime` ordering example (from `$filter` + `$orderby` combined query, S6):
```
GET /me/messages?$filter=Subject eq 'welcome' and importance eq 'normal'&$orderby=subject,importance,receivedDateTime desc
```

### 3.4 `$filter` + `$orderby` combined — documented CONSTRAINT (exact wording)

From S1 ("Using filter and orderby in the same query"), which S6 cross-references as authoritative for messages:

> "When using `$filter` and `$orderby` in the same query to get messages, make sure to specify properties in the following ways:
> 1. Properties that appear in `$orderby` must also appear in `$filter`.
> 2. Properties that appear in `$orderby` are in the same order as in `$filter`.
> 3. Properties that are present in `$orderby` appear in `$filter` before any properties that aren't.
>
> Failing to do this results in the following error:
> - Error code: `InefficientFilter`
> - Error message: `The restriction or sort order is too complex for this operation.`"

This is the exact, quotable error code/message pair the goal asked for. It is a **`$filter`+`$orderby` interaction**, not a `$search` interaction (see §3.7 for why `$search` combined with `$orderby` is a separate, disjoint constraint).

### 3.5 `$top` — page size
- Minimum `1`. **Maximum for `/me/messages` and `/me/mailFolders/{id}/messages` is 1000** (stated explicitly in S1: "Use `$top` to customize the page size, within the range of 1 and 1000"). Default page size (server-side, when `$top` is unspecified) is **10 messages**. [S1]
- General Graph doc (S6) states only "minimum value of `$top` is 1 and the maximum depends on the corresponding API" — the messages-specific ceiling of 1000 is the authoritative number, taken from S1.
- Large `$top` with full-payload messages risks HTTP 504 gateway timeout — S1 explicitly recommends pairing large `$top` with `$select`.
```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/me/messages?\$top=5"
```

### 3.6 `$skip` and `$count`
- `$skip` is used by the Mail API (message/event/calendar) as part of its **paging** mechanism (server round-trips via `@odata.nextLink`, see §5). S1 explicitly warns: **do not** extract the `$skip` value from `@odata.nextLink` and reuse/manipulate it manually — the API uses `$skip` internally as a running cursor count through the mailbox, and it can legitimately be larger than the page size even on page 1.
- `$count=true` returns an `@odata.count` alongside a page of items:
```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/me/messages?\$top=2&\$count=true"
```
`$count` used as a bare URL segment (`/me/messages/$count`) returns only the integer total — this pattern is documented generically in S6 for collections (e.g. `/users/$count`), not verified specifically against `/me/messages/$count` in the fetched docs (flagged §9).

### 3.7 `$search` + `$orderby` / `$filter` — documented CONSTRAINTS

Two separate, real constraints, from two different docs:

1. **`$search` on messages carries an implicit sort that cannot be overridden by `$orderby` for events** — S6 states: *"When you specify `$filter`, the service infers a sort order for the results... you must specify properties in certain ways"* — this is the `$filter`+`$orderby` rule already quoted in §3.4. For `$search` specifically, S12 states search results for messages "are sorted by the date and time the message was sent" — i.e. `$search` has its **own built-in sort**, separate from `$orderby`.
2. **Sorting is not supported for `$search` on events, and returns Bad Request** — this exact constraint is documented in S5 (Search API `/search/query`, not OData `$search`): *"Sorting results is not supported for events. A sort clause in the request will return a Bad Request error code in the response."* This is scoped to **events**, not messages, in S5's wording — flagged explicitly because the goal asked about message `$search`+`$orderby` and the only "Bad Request on sort" wording found in the fetched docs applies to events in the POST `/search/query` API, not to `GET /me/messages?$search=`. No sentence in S1, S2, S4, S6, or S12 states that combining OData `$search` and `$orderby` on `GET /me/messages` is rejected outright. See §9 — treat "does `$search`+`$orderby` on `/me/messages` (GET form) 400" as **unverified**; the CLI should defensively avoid combining them and fall back to client-side sort if needed.

`$search` + `$filter` combined (documented generically for other resources, S6/S12, not message-specific in the fetched pages): *"Results are filtered based on a logical conjunction (AND) of the `$filter` and the entire query in the `$search`."* No message-specific example of `$search`+`$filter` together was found in S1/S12 (flagged §9).

---

## 4. `$search` and KQL (message search)

Two distinct search surfaces exist — do not conflate them:

**(a) OData `$search` on `/me/messages` (GET, simple):**
```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  'https://graph.microsoft.com/v1.0/me/messages?$search="pizza"'
```
- Quoting: the search string itself must be wrapped in double quotes inside the query value (`$search="pizza"`), per S12's own examples.
- Default properties searched (no KQL property prefix given): **from, subject, body** — exact wording, S12: *"When you search messages without specifying message properties, the search targets these default properties: from, subject, and body."*
- Results: **up to 1,000 results**, sorted by **the date and time the message was sent** (S12, message-specific — note this differs from `sentDateTime` vs `receivedDateTime` naming; S12's own words say "sorted by the date and time the message was sent").
- **`ConsistencyLevel` header**: **not required** for message `$search` per the fetched docs. S12/S6 only document `ConsistencyLevel: eventual` as required for `$search` against **directory objects** (Entra ID resources deriving from `directoryObject`, e.g. `/groups`, `/users`) — every `ConsistencyLevel` example in S6/S12 is on a `/groups?...` or `/users?...` query, never on `/me/messages?...`. None of the message-search examples (S1, S5, S12) include a `ConsistencyLevel` header. Treat this absence as the answer: **no `ConsistencyLevel` header is documented as required for `/me/messages?$search=`.**

**(b) KQL property-scoped search (`property:value` syntax inside `$search`):**
```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  'https://graph.microsoft.com/v1.0/me/messages?$search="from:randiw"&$select=subject,from'
```
Full searchable-property table (verbatim from S12):

| KQL property | Meaning | Example |
|---|---|---|
| `attachment` | Names of attached files | `$search="attachment:api-catalog.md"` |
| `bcc` | Bcc field (SMTP/display name/alias) | `$search="bcc:samanthab@contoso.com"` |
| `body` | Message body | `$search="body:excitement"` |
| `cc` | Cc field | `$search="cc:danas"` |
| `from` | Sender | `$search="from:randiw"`; OR example: `$search="from:adelev OR from:alexw OR from: allanD"` |
| `hasAttachment` | `true`/`false` — has a non-inline attachment | `$search="hasAttachments:true"` *(sic — table's own example uses `hasAttachments`, plural, despite the property name column reading `hasAttachment`, singular — quoted verbatim from S12; treat both spellings as attested by the docs and test both if implementing)* |
| `importance` | `low`, `medium`, or `high` | `$search="importance:high"` |
| `kind` | `contacts`, `docs`, `email`, `faxes`, `im`, `journals`, `meetings`, `notes`, `posts`, `rssfeeds`, `tasks`, `voicemail` | `$search="kind:voicemail"` |
| `participants` | from/to/cc/bcc combined | `$search="participants:danas"` |
| `received` | Date recipient received the message | `$search="received:07/23/2018"` |
| `recipients` | to/cc/bcc combined | `$search="recipients:randiw"` |
| `sent` | Date sender sent the message | `$search="sent:07/23/2018"` |
| `size` | Item size in bytes, range syntax `N..M` | `$search="size:1..500000"` |
| `subject` | Subject line text | `$search="subject:has"` |
| `to` | To field | `$search="to:randiw"` |

Search-string general syntax rules (S12, applies broadly to `$search`, verified against message examples):
- General form: `$search="clause1" [AND | OR] "clauseX"`; any number of clauses; parentheses supported for precedence.
- Each clause: `"<property>:<text to search>"`, whole clause in double quotes; embedded double quotes or backslashes must be backslash-escaped; other special characters must be URL-encoded.
- `AND`/`OR` must be **outside** the double quotes and **uppercase**.
- Tokenization: splits on spaces, casing changes (lower→upper only), symbols, and digit boundaries — e.g. `hello.world` → tokens `hello`, `.`, `world`, `helloworld`; matches are case-insensitive and order-independent after tokenization.
- **No substring/"contains" search** — S12: *"The search implementation doesn't support 'contains' logic. Instead, it uses a tokenization approach..."*

**(c) `POST /search/query` (Microsoft Search API)** — a separate, richer API, distinct from `GET ?$search=`. [S5]
```bash
curl -s -X POST \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/json" \
  -d '{
    "requests": [
      { "entityTypes": ["message"], "query": { "queryString": "contoso" }, "from": 0, "size": 25 }
    ]
  }' \
  "https://graph.microsoft.com/v1.0/search/query"
```
Response (S5, condensed):
```json
{
  "value": [
    {
      "searchTerms": ["contoso"],
      "hitsContainers": [
        {
          "total": 1,
          "moreResultsAvailable": false,
          "hits": [
            {
              "hitId": "ptWLQ4o6HYpQg8xmAAATzOzRAAA=",
              "rank": 1,
              "summary": "Here is a summary of your messages from last week",
              "resource": {
                "@odata.type": "#microsoft.graph.message",
                "subject": "Weekly digest: Microsoft 365 changes",
                "sender": { "emailAddress": { "name": "Office365 Message Center", "address": "gtorres@contoso.com" } }
              }
            }
          ]
        }
      ]
    }
  ]
}
```
Key facts for this API, verbatim/paraphrased from S5:
- Message search applies to **the signed-in user's own mailbox only** — *"Users can search their own mailbox, but can't search delegated mailboxes."*
- Search covers **body and attachments**; supported attachment file types match SharePoint Online search's crawled file types.
- Results default-sorted by `receivedDateTime` descending; set `"enableTopResults": true` in the request to sort by relevance instead.
- `total` in the response's `hitsContainer` is **the count on the current page, not the total match count** — explicit known-limitation wording: *"For messages, the total property of the searchHitsContainer type contains the number of results on the page, not the total number of matching results."*
- Explicit sort-rejection wording (scoped to **events**, not messages — see §3.7 discussion): *"Sorting results is not supported for events. A sort clause in the request will return a Bad Request error code in the response."*

---

## 5. Pagination

Mechanics from S1, S6, S7 (S7 is the canonical paging doc).

- Default/server-side page size for `/me/messages` is **10** unless `$top` overrides it (max 1000, see §3.5). [S1]
- When more results remain, the response body includes `@odata.nextLink`: a **complete URL** (includes all original query parameters, re-encoded, plus a `$skiptoken` or `$skip`) to fetch the next page.
- **Client loop contract** (S7, verbatim guidance): *"Use the entire URL in the `@odata.nextLink` property in a GET request to retrieve the next page of results... Don't try to extract the `$skiptoken` or `$skip` value and use it in a different request."* Continue following `@odata.nextLink` until a response no longer contains it — that is the only correct termination condition.
- Mail API (`message`, `event`, `calendar`) specifically uses **`$skip`** in `@odata.nextLink` (not `$skipToken`, which other Graph APIs use). [S6: "Some Microsoft Graph APIs, like Outlook Mail and Calendars (message, event, and calendar), use `$skip` to implement paging."]
- Example `@odata.nextLink` shape (generic Graph pattern, S7):
```
"@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$top=5&$skiptoken=RFNwdAIAAQAAAD8...AAAAAAAA"
```
For messages specifically the equivalent parameter is `$skip` per S6's explicit callout above.

**Node.js client loop (pseudocode, grounded in the "follow nextLink verbatim, don't reconstruct" contract):**
```js
async function listAllMessages(client, startUrl) {
  let url = startUrl; // e.g. https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$select=subject,receivedDateTime&$top=50
  const all = [];
  while (url) {
    const res = await client.get(url); // GET with Authorization header
    all.push(...res.data.value);
    url = res.data['@odata.nextLink'] || null; // stop condition: key absent
  }
  return all;
}
```
- Reminder from S1: don't try to derive or manipulate the `$skip` value yourself — it's an opaque running cursor into the mailbox, not a plain row offset, and can already exceed the page size on page 1.
- Retry safety (documented for directory paging in S7 but stated as general paging risk guidance): if a page request is retried, don't use a token obtained from the retry response for the *next* page — always carry forward the token from the last **successful, non-retried** response, to avoid `DirectoryPageTokenNotFoundException`-class errors.
- Any custom headers (e.g. `Prefer: outlook.body-content-type`) present on the first request are **not automatically preserved** across `@odata.nextLink` calls for directory-object paging per S6/S7; the fetched docs don't state this explicitly for **mail** paging (flagged §9) — the CLI should defensively resend any `Prefer` headers on every page request rather than assume they persist.

---

## 6. The `message` resource

Source: S8 (resource reference), example payload merged from S2 (get-message example).

### 6.1 Property table (declared properties, verbatim from S8)

| Property | Type | Description |
|---|---|---|
| `bccRecipients` | `recipient` collection | The Bcc: recipients. |
| `body` | `itemBody` | The message body, HTML or text. |
| `bodyPreview` | String | First 255 characters of the body, always text format. |
| `ccRecipients` | `recipient` collection | The Cc: recipients. |
| `changeKey` | String | Version of the message. |
| `conversationId` | String | ID of the conversation the email belongs to. |
| `conversationIndex` | Edm.Binary | Position of the message within the conversation. |
| `createdDateTime` | DateTimeOffset | Creation time, ISO 8601 UTC. |
| `flag` | `followupFlag` | Status/start/due/completion date for follow-up. |
| `from` | `recipient` | Mailbox owner the message is sent from (usually == `sender`, except sharing/delegation). |
| `hasAttachments` | Boolean | True if there are non-inline attachments. Doesn't count inline attachments — a message with only inline attachments reports `false`. To detect inline attachments, parse `body` for a `src="cid:..."` attribute. |
| `id` | String | Unique identifier. **Changes by default when the item moves between containers (folders)** — see §8. |
| `importance` | `importance` enum | `low`, `normal`, `high`. |
| `inferenceClassification` | enum | `focused` or `other`. |
| `internetMessageHeaders` | `internetMessageHeader` collection | RFC5322 headers; **requires `$select` to retrieve**; read-only. |
| `internetMessageId` | String | RFC2822-format message ID. |
| `isDeliveryReceiptRequested` | Boolean | Delivery receipt requested. |
| `isDraft` | Boolean | True if not yet sent. |
| `isRead` | Boolean | Read/unread. |
| `isReadReceiptRequested` | Boolean | Read receipt requested. |
| `lastModifiedDateTime` | DateTimeOffset | Last-change time, ISO 8601 UTC. |
| `parentFolderId` | String | Unique ID of the containing `mailFolder`. |
| `receivedDateTime` | DateTimeOffset | Receipt time, ISO 8601 UTC. |
| `replyTo` | `recipient` collection | Addresses to use when replying. |
| `sender` | `recipient` | Account used to generate the message (usually == `from`; can differ for shared mailbox/delegate send). |
| `sentDateTime` | DateTimeOffset | Send time, ISO 8601 UTC. |
| `subject` | String | Message subject. |
| `toRecipients` | `recipient` collection | The To: recipients. |
| `uniqueBody` | `itemBody` | The part of the body unique to this message (i.e. minus quoted thread history). **Not returned by default** — requires `?$select=uniqueBody`. HTML or text. |
| `webLink` | String | URL to open the message in Outlook on the web. Supports `?ispopout=0` (reading-pane view) / `1` or absent (popout). Cannot be used inside an iframe. |

Relationships (navigation properties): `attachments` (`fileAttachment`/`itemAttachment`/`referenceAttachment` collection), `extensions`, `multiValueExtendedProperties`, `singleValueExtendedProperties`. [S8]

Recipient-count limit (documented alongside the resource, not a property but an operational ceiling): **max 500 total recipients** across `toRecipients` + `ccRecipients` + `bccRecipients` for a message sent from Exchange Online. [S8]

### 6.2 `body` vs `bodyPreview` vs `uniqueBody`

- **`body`** (`itemBody`: `{ contentType: "html"|"text", content: string }`) — the full message body. **Returned in HTML by default** if no `Prefer` header is sent. S1 additionally notes that the *list-messages* operation "currently... returns message bodies in only HTML format" regardless of Prefer header on that specific endpoint — treat single-message GET (`message-get`) as the operation where the `Prefer` header reliably switches format (see below); list behavior should be verified in-app before relying on plain text from `/me/messages`.
- **`bodyPreview`** — first 255 characters, always plain text, always returned by default (part of the default property set), useful for lightweight list rendering without asking for the full `body`.
- **`uniqueBody`** — the portion of the body unique to *this* message (excludes quoted earlier messages in the thread/conversation). Not in the default `$select` set; must explicitly request `?$select=uniqueBody`. Same `contentType` HTML/text duality as `body`.

### 6.3 `Prefer: outlook.body-content-type="text"` — forcing plain text

Documented on `message-get` (S2): forces both `body` and `uniqueBody` to be returned as `contentType: "text"` instead of the HTML default. Server confirms application of the preference via a `Preference-Applied` response header.

```bash
curl -s \
  -H "Authorization: Bearer $TOKEN" \
  -H 'Prefer: outlook.body-content-type="text"' \
  "https://graph.microsoft.com/v1.0/me/messages/AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAASoXUT3AAA=?\$select=subject,body,bodyPreview,uniqueBody"
```
Response (S2, verbatim):
```
HTTP/1.1 200 OK
Content-type: application/json
Preference-Applied: outlook.body-content-type="text"

{
    "@odata.context":"https://graph.microsoft.com/v1.0/$metadata#users('cd209b0b-3f83-4c35-82d2-d88a61820480')/messages(subject,body,bodyPreview,uniqueBody)/$entity",
    "id":"AAMkAGI1AAAoZCfHAAA=",
    "subject":"Welcome to our group!",
    "bodyPreview":"Welcome to our group, Dana! Hope you will enjoy working with us !\r\n\r\n...",
    "body":{
        "contentType":"text",
        "content":"Welcome to our group, Dana! Hope you will enjoy working with us [😊] ...\r\n\r\nThanks!\r\n\r\n"
    },
    "uniqueBody":{
        "contentType":"text",
        "content":"Welcome to our group, Dana! Hope you will enjoy working with us [😊] ...\r\nThanks!\r\n"
    }
}
```
Valid `Prefer: outlook.body-content-type` values: `"text"` or `"html"`. If omitted, `body`/`uniqueBody` are returned in HTML. [S2]

### 6.4 Full single-message example (`GET /me/messages/{id}`)

```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/me/messages/AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAASoXUT3AAA="
```
Response (S2, verbatim, condensed):
```json
{
    "id":"AAMkADhMGAAA=",
    "createdDateTime":"2018-09-09T03:15:05Z",
    "lastModifiedDateTime":"2018-09-09T03:15:08Z",
    "changeKey":"CQAAABYAAAC4ofQHEIqCSbQPot83AFcbAAAnjjuZ",
    "categories":[],
    "receivedDateTime":"2018-09-09T03:15:08Z",
    "sentDateTime":"2018-09-09T03:15:06Z",
    "hasAttachments":false,
    "internetMessageId":"<MWHPR6E1BE060@MWHPR1120.namprd22.prod.outlook.com>",
    "subject":"9/9/2018: concert",
    "bodyPreview":"The group represents Nevada.",
    "importance":"normal",
    "parentFolderId":"AAMkADcbAAAAAAEJAAA=",
    "conversationId":"AAQkADOUpag6yWs=",
    "isDeliveryReceiptRequested":false,
    "isReadReceiptRequested":false,
    "isRead":true,
    "isDraft":false,
    "webLink":"https://outlook.office365.com/owa/?ItemID=AAMkADMGAAA%3D&exvsurl=1&viewmodel=ReadMessageItem",
    "inferenceClassification":"focused",
    "body":{ "contentType":"html", "content":"<html>...The group represents Nevada....</html>" },
    "sender":{ "emailAddress":{ "name":"Adele Vance", "address":"adelev@contoso.com" } },
    "from":{ "emailAddress":{ "name":"Adele Vance", "address":"adelev@contoso.com" } },
    "toRecipients":[ { "emailAddress":{ "name":"Alex Wilber", "address":"AlexW@contoso.com" } } ],
    "ccRecipients":[],
    "bccRecipients":[],
    "replyTo":[],
    "flag":{ "flagStatus":"notFlagged" }
}
```

### 6.5 MIME content (`$value`)
```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/me/messages/4aade2547798441eab5188a7a2436bc1/\$value"
```
Returns raw RFC-822 MIME (`Content-type: text/plain`), not a JSON `message` resource — headers, `multipart/mixed` boundaries, and base64 attachment payloads exactly as transmitted. [S2]

---

## 7. Attachments

Source: S9 (list-attachments operation), S14 (base `attachment` resource), S15/S16/S17 (derived types).

### 7.1 Endpoints
```
GET /me/messages/{id}/attachments
GET /users/{id | userPrincipalName}/messages/{id}/attachments
GET /me/mailFolders/{id}/messages/{id}/attachments
GET /me/mailFolders/{id}/childFolders/{id}/.../messages/{id}/attachments/{id}   # arbitrary nesting depth
```
Permission required: `Mail.Read` (delegated and application — no `ReadBasic` tier documented for attachments). [S9]

`$filter` on `/attachments` is **silently ignored** if present (documented known limitation, S18) — don't rely on server-side attachment filtering; filter client-side after listing.

### 7.2 `#odata.type` discriminator

`attachment` is the **base type**; every item in the `attachments` collection is one of three derived types, distinguished by its `@odata.type` value: [S14]

| `@odata.type` | Resource | What it represents |
|---|---|---|
| `#microsoft.graph.fileAttachment` | `fileAttachment` | A file (text, Word doc, etc.) — carries the actual bytes in `contentBytes`. |
| `#microsoft.graph.itemAttachment` | `itemAttachment` | A contact, event, or **message** attached as an Outlook item — carries an `item` navigation property (`OutlookItem`), not `contentBytes`. |
| `#microsoft.graph.referenceAttachment` | `referenceAttachment` | A **link** to a file on OneDrive/SharePoint/other supported storage — no raw bytes, has location/URL metadata instead. |

Base properties common to all three (`attachment`, S14): `contentType` (MIME type string), `id`, `isInline` (bool), `lastModifiedDateTime`, `name`, `size` (Int32, bytes).

`fileAttachment`-specific properties (S15): `contentBytes` (Edm.Binary — **base64-encoded** file content), `contentId` (Exchange-store attachment ID), `contentLocation` (documented as **"Don't use this property as it isn't supported"**). When *creating* a file attachment, the request body must include `"@odata.type": "#microsoft.graph.fileAttachment"` plus required `name` and `contentBytes`.

`itemAttachment`-specific (S16): adds `item` relationship (`OutlookItem` — i.e. the attached contact/event/message itself, as a full resource), no `contentBytes`.

`referenceAttachment`-specific (S17): same base properties; `size` here is documented as **the size of the metadata stored on the message for the attachment, not the size of the actual linked file**.

Attachment size guidance (S14, cross-cutting): attachments under 3 MB can be added directly via `POST` to the attachments collection; 3 MB–150 MB requires an upload session (`attachment-createuploadsession`).

### 7.3 List-attachments example
```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/me/messages/AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAASoXUT3AAA=/attachments"
```
Response (S9, verbatim):
```json
{
  "value": [
    {
      "@odata.type": "microsoft.graph.fileAttachment",
      "contentType": "contentType-value",
      "contentLocation": "contentLocation-value",
      "contentBytes": "contentBytes-value",
      "contentId": "null",
      "lastModifiedDateTime": "datetime-value",
      "id": "id-value",
      "isInline": false,
      "name": "name-value",
      "size": 99
    }
  ]
}
```
(Note: S9's own example response omits the leading `#` on `@odata.type` — written as `"microsoft.graph.fileAttachment"` rather than `"#microsoft.graph.fileAttachment"`. The `#`-prefixed form is what S14/S15 specify for *creating* an attachment. Both forms are quoted verbatim from their respective source docs — implement tolerant parsing that strips an optional leading `#` when comparing discriminator values.)

### 7.4 Downloading `contentBytes`
`contentBytes` on a `fileAttachment` is **base64-encoded** binary (Edm.Binary). [S15] To reconstruct the original file in Node.js:
```js
const fs = require("fs");
const buf = Buffer.from(attachment.contentBytes, "base64");
fs.writeFileSync(attachment.name, buf);
```
`itemAttachment` and `referenceAttachment` have **no `contentBytes`** — don't attempt base64-decoding on those types; branch on `@odata.type` first.

---

## 8. Immutable IDs — operational landmine for a CLI that stores IDs

Source: S10.

### 8.1 The problem
By default, a Graph `message.id` (like all Outlook item IDs — messages, events, contacts, tasks) **changes when the item moves to a different folder** within the same mailbox (e.g. Inbox → a custom folder, or into Deleted Items). A CLI that persists `message.id` values (e.g. as a local cache key, a "seen" marker, or a dedup key) will silently lose the ability to look those messages up again after any move — including a user-initiated drag-and-drop in Outlook, a rule-based auto-file, or the CLI's own move/archive actions. [S10][S8 `id` property description confirms: "By default, this value changes when the item is moved from one container (such as a folder or calendar) to another."]

### 8.2 The fix — `Prefer: IdType="ImmutableId"`
```bash
curl -s \
  -H "Authorization: Bearer $TOKEN" \
  -H 'Prefer: IdType="ImmutableId"' \
  "https://graph.microsoft.com/v1.0/me/messages?\$select=id,subject"
```
- This header is **per-request** — "If you want to always use immutable IDs, you must include this header with every API request." [S10] There is no mailbox-level or app-level toggle; the CLI must attach this header to every single mail-read call if it wants ID stability guaranteed.
- IDs are **case-sensitive** — a landmine in its own right if the CLI does case-insensitive string comparison/dedup on stored IDs. [S10]

### 8.3 What immutable IDs do *not* protect against
An immutable ID still changes if: [S10]
- the item is moved to an **archive mailbox** (distinct from moving between folders in the same mailbox — this case is NOT covered), or
- the item is **exported** (PST, .msg, etc.) and **re-imported**.

So "immutable" means "stable across in-mailbox folder moves," not "permanently stable for the life of the item under all operations."

### 8.4 Which resources support it
`message`, `attachment`, `event`, `eventMessage`, `contact`, `outlookTask`. **Container types** (`mailFolder`, `calendar`) do **not** support/need immutable IDs — their regular IDs were already constant. [S10]

### 8.5 Migrating an existing store of non-immutable IDs
`translateExchangeIds` converts already-stored regular IDs to immutable format in batches of up to 1000:
```bash
curl -s -X POST \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/json" \
  -d '{
    "inputIds": ["AQMkAGM2..."],
    "targetIdType": "restImmutableEntryId",
    "sourceIdType": "restId"
  }' \
  "https://graph.microsoft.com/v1.0/me/translateExchangeIds"
```
Response (S10, verbatim):
```json
{
  "value": [
    { "targetId": "AAkALgAA...", "sourceId": "AQMkAGM2..." }
  ]
}
```
This is also the documented path for migrating Exchange Web Services (EWS) IDs into Graph. [S10]

### 8.6 CLI implementation recommendation
Send `Prefer: IdType="ImmutableId"` on every mail-read request (list messages, get message, list attachments) from day one — retrofitting a store of regular IDs later requires the batch `translateExchangeIds` migration step above, and any IDs persisted before the header was adopted are subject to silent invalidation on the next folder move.

---

## 9. Unverified / Not Found (explicit gaps)

The following could **not** be confirmed with exact wording from the 10 assigned docs + the supporting docs fetched, and should not be treated as verified fact by implementers:

1. **Whether `GET /me/messages?$search=...&$orderby=...` (OData form) is rejected outright.** The only "sort + Bad Request" sentence found is scoped to **events** in the `POST /search/query` API (S5): *"Sorting results is not supported for events. A sort clause in the request will return a Bad Request error code in the response."* No sentence in S1/S2/S4/S6/S12 states an analogous rejection for `$search`+`$orderby` on **messages** via the plain OData `GET ?$search=` form. Treat as **unconfirmed** — defensively avoid the combination in the CLI and sort client-side if both are needed.
2. **Whether `$search` combined with `$filter` on `/me/messages` specifically is supported or rejected**, and any exact error text for that combination. S6/S12 only document the AND-conjunction behavior for `$search`+`$filter` generically (using `/groups` examples), not for messages.
3. **Exact maximum value of `$top` for `/me/mailFolders` or `/me/messages/{id}/attachments`** — S1 gives the 1–1000 range explicitly for messages; no equivalent explicit ceiling was found for mail folders or attachments collections in the fetched docs.
4. **Whether `Prefer` headers (e.g. `outlook.body-content-type`) persist automatically across `@odata.nextLink` pagination requests for mail.** S6/S7 state this negatively only for **directory-object** paging (`ConsistencyLevel`); no equivalent statement was found for mail/message paging. Recommendation: resend explicitly on every page request.
5. **`hasAttachment` vs `hasAttachments` naming inconsistency in `$search`.** S12's own property-name column reads `hasAttachment` (singular) but its own worked example uses `$search="hasAttachments:true"` (plural). Quoted verbatim from the source — this is the docs' own inconsistency, not a transcription error on this reference's part. Test both spellings against a live tenant before shipping.
6. **`GET /me/messages/$count` (bare-segment count) against messages specifically** — the `$count`-as-URL-segment pattern is documented generically (S6, using `/users/$count`, `/groups/{id}/members/$count`), not demonstrated against `/me/messages` in the fetched docs.
7. **Whether the list-messages HTML-only-body behavior noted in S1** ("Currently, this operation returns message bodies in only HTML format") **is still current / whether `Prefer: outlook.body-content-type="text"` actually overrides it on `/me/messages` (list) vs. only on `/me/messages/{id}` (get).** S1's body-format note and S2's `Prefer`-header example are on two different operations (list vs. get); the docs don't explicitly reconcile whether `Prefer` is honored on the list endpoint too. Recommend testing directly.
8. **Rate limiting / throttling numeric thresholds** for the mail read surface — not covered by any of the 10 assigned docs or the supporting docs fetched; not included in this reference.
9. **OAuth/token acquisition flow details** — out of scope for this reference (task was read-endpoints only); not fetched.

No sub-agents were dispatched for this research; all fetches above were performed sequentially and directly.
