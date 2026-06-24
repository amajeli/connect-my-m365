---
name: MCP tool response token efficiency
description: Conventions for keeping MCP connector tool responses lean so they don't burn Claude.ai context tokens.
---

For an MCP connector, every byte a tool returns is fed into the LLM's context. "Token burn" is dominated by tool *response* size, so optimize there.

**Conventions (apply to every tool):**
1. **Emit compact JSON** — `JSON.stringify(x)`, never `JSON.stringify(x, null, 2)`. Indentation/newlines inflate every response with zero information gain.
2. **Cap free-form text fields** — email/chat bodies and file contents must be length-capped with a truncation marker. An uncapped field (e.g. a 512 KB OneDrive text file or a giant marketing email body) can dump 100k+ tokens in one call, directly defeating the connector's usefulness.
3. **Fetch only the fields a given caller needs from the upstream API** — e.g. list/search email tools only use `bodyPreview`, so they request Graph data *without* the full `body`. The detail tool fetches the body. This cuts upstream payload/latency too.

**Why:** the user's explicit standing requirement is "most efficient manner to at all times reduce token burn." These are the highest-leverage, always-applicable levers.

**Constraint to remember:** a Graph fetch helper can be shared between the REST API (web UI) and MCP tools. The web inbox renders the full HTML email body, so the body must stay available for the REST path — make it opt-out (e.g. an `includeBody` flag) rather than removing it outright, and have MCP callers pass the lean option.
