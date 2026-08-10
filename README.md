# ScopeGuardian — AI Scope Monitor

Event-driven scope creep detection for enterprise projects. Monitors SharePoint task lists via Microsoft Graph change notifications, analyses every change against the project's master scope using Azure OpenAI + vector search, and alerts the project team in Microsoft Teams with actionable Adaptive Cards.

## Architecture

```
SharePoint list change
        │  (Graph change notification, changeType: updated)
        ▼
EventTrigger (HTTP)          ── validates clientState, resolves subscription → project
        │  enqueue
        ▼
scope-analysis (Storage Queue)
        │
        ▼
ProcessScopeAnalysis (Queue) ── Graph delta query → changed items → dedup (eTag)
        │  per item
        ▼
scopeAnalyzer                ── embed → AI Search vector match → map/reduce LLM
        │                       violation detection (gpt-4o-mini map, gpt-4o reduce)
        ├─► Cosmos DB           violations, task events, summaries
        └─► Teams (Bot)         Adaptive Card: Acknowledge / Dismiss / Escalate
                                     │
                                     ▼
                          HandleTeamsFeedback (HTTP)
                                     └─► violation status + feedback log (Cosmos)

Timers:
  GraphSubscriptionManager (12h) — creates/renews Graph subscriptions, primes delta links
  SyncMasterScope (6h)           — re-indexes scope into AI Search, regenerates Markdown summary

Scope onboarding:
  ScopeManagement (HTTP)         — upload/fetch master scope, immediate indexing
```

### Key design decisions

- **Async pipeline:** Graph requires webhook responses within seconds. EventTrigger only validates and enqueues; all AI work happens in the queue worker.
- **Delta queries:** SharePoint notifications carry no item data. The worker runs a Graph delta query per notification and stores the delta link per project (in the subscription record).
- **Idempotency:** every item version (itemId + eTag) is marked in a TTL'd Cosmos container before analysis. Graph redeliveries and queue retries never double-analyse or double-alert.
- **Deterministic risk score:** severity-weighted (low 5 / medium 10 / high 20 / critical 30, capped at 100) — same function everywhere.

## Azure resources required

| Resource | Purpose |
|---|---|
| Function App (Node 20, Consumption/Premium) | All triggers |
| Storage Account | Function runtime + `scope-analysis` queue + master scope blobs |
| Cosmos DB (SQL API) | Violations, events, summaries, subscriptions, dedup, conversation refs (containers auto-created) |
| Azure AI Search (Basic+) | Vector index over scope items (index auto-created) |
| Azure OpenAI | Deployments: chat (`gpt-4o`), mini (`gpt-4o-mini`), embedding (`text-embedding-3-small`, 1536 dims) |
| Azure Bot resource | Teams channel enabled, messaging endpoint → `/api/bot/messages` |

## Setup

### 1. Entra ID app registration (Graph access)

1. Create an app registration; note tenant ID, client ID; create a client secret.
2. API permissions → Microsoft Graph → **Application** → `Sites.Read.All` → grant admin consent.
   (Alternatively omit `AZURE_CLIENT_SECRET` and grant the same Graph app role to the Function App's Managed Identity.)

### 2. Bot registration (Teams)

1. Create an **Azure Bot** resource (multi-tenant). Note the Microsoft App ID + password.
2. Channels → enable **Microsoft Teams**.
3. Messaging endpoint: `https://<function-app>.azurewebsites.net/api/bot/messages`.
4. Package the bot in a Teams app manifest and install it in the target team (required for proactive channel alerts).
5. First time a user talks to the bot, the real regional `serviceUrl` is captured and stored — proactive alerts prefer it over `BOT_SERVICE_URL`.

### 3. Configuration

Copy `local.settings.json.example` → `local.settings.json` (locally) or set the same values as Function App settings in Azure. Notable settings:

| Setting | Notes |
|---|---|
| `PROJECT_CONFIGS` | **Single source of truth** for monitored projects: `projectId:siteId:listId,...` |
| `WEBHOOK_NOTIFICATION_URL` | Public URL of EventTrigger **including function key** (`?code=...`) |
| `WEBHOOK_CLIENT_STATE` | Long random secret — validates notification authenticity |
| `TEAMS_CHANNEL_ID` | Channel that receives violation alerts (`19:...@thread.tacv2`) |
| `AZURE_OPENAI_*_DEPLOYMENT` | Must match your Azure OpenAI deployment names |
| `DEFAULT_PROJECT_ID` | Fallback for bot commands without an explicit project id |

### 4. Onboard a project

1. Add the project to `PROJECT_CONFIGS` (site ID and list ID from Graph or SharePoint).
2. Upload the master scope (the subscription manager picks the project up within 12 hours, or restart the app for immediate bootstrap):

```bash
curl -X PUT "https://<function-app>.azurewebsites.net/api/projects/project-001/scope?code=<function-key>" \
  -H "Content-Type: application/json" \
  -d @scope.json
```

`scope.json` shape:

```json
{
  "projectName": "Customer Portal",
  "version": "1",
  "scopeItems": [
    {
      "id": "S1",
      "title": "User login",
      "description": "Entra ID SSO login for portal users",
      "deliverables": ["Login page", "SSO integration"],
      "acceptanceCriteria": ["Users authenticate via Entra ID"],
      "tags": ["auth"]
    }
  ],
  "outOfScope": ["Mobile app", "Custom reporting"],
  "assumptions": ["Customer provides tenant access"],
  "constraints": ["Go-live before Q3"]
}
```

Endpoints: `PUT/POST` uploads + indexes immediately; `GET .../scope` returns the latest scope; `GET .../scope/summary` returns the stakeholder Markdown summary.

### 5. Deploy

No CI pipeline — deployed directly with Azure Functions Core Tools:

```bash
npm run build
func azure functionapp publish <FUNCTION-APP-NAME>
```

## Teams commands

- `/status <projectId>` — risk score + pending violation count
- `/violations <projectId>` — list pending violations
- Violation cards: **Acknowledge**, **Dismiss (False Positive)**, **Escalate** — with optional comment; all feedback is logged.

## Local development

```bash
npm install
npm run build
npm test
func start          # requires Azurite or a real storage connection
```

Note: Graph webhooks require a publicly reachable notification URL — use a tunnel (e.g. `ngrok`) and set `WEBHOOK_NOTIFICATION_URL` accordingly for end-to-end local testing.

## Operational notes

- **Subscriptions** expire after 3 days; the manager renews anything expiring within 24 h every 12 h run, and re-creates subscriptions Graph has deleted.
- **Poison messages** land in `scope-analysis-poison` after 3 dequeues — monitor it.
- **Dedup records** expire after 7 days (Cosmos TTL).
- **Costs:** each analysed change ≈ 1 embedding call + 1–2 mini calls + 1 gpt-4o call. The 6-hour sync re-embeds all scope items per project.
- **Secrets** live in app settings; move to Key Vault references for production hardening. Rotate `WEBHOOK_CLIENT_STATE` and the client secret periodically.
