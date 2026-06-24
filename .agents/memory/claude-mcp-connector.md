---
name: Claude.ai MCP connector debugging
description: Non-obvious failure modes when a Replit-hosted MCP connector "won't connect in Claude".
---

When an MCP connector "won't connect in Claude", these are the non-obvious blockers (each can occur independently):

1. **Private deployment blocks server-to-server clients.** A Private Replit deployment forces a Replit login on every request. The owner's browser passes silently, so the app "works for me" but Claude (no Replit account) is bounced to a login page and never reaches any endpoint. Tell: an external fetch of `/.well-known/...` returns a Replit login page / `privateDeployment=true`. Fix: set the deployment Public.

2. **Dev and prod are separate databases.** Data/registrations created while testing the dev URL do not exist in prod. Confirm by comparing row counts across environments before assuming a write "happened".

3. **Claude caches its Dynamic Client Registration per issuer URL** and will reuse a stale `client_id`, skipping re-registration even after the user removes/re-adds the connector. So "remove and re-add" is an unreliable fix.
   **Decision — self-heal on the server instead:** auto-create a client record for an unrecognized previously-issued `client_id` at the authorize step, and trust-on-first-use the presented client secret at the token step (then enforce it thereafter).
   **Why:** the client_id/secret were originally issued by us and PKCE + the upstream identity login still gate the flow, so accepting a former client back is bounded and far more reliable than depending on the client to re-register.

4. **Tier/subscription gating is separate from identity.** Being an admin/owner is not the same as having an *activated* access tier; a status of `none` still rejects MCP calls. Activate the tier explicitly.

**General:** a code fix only reaches production after the user redeploys. If prod still emits an error string your new code no longer produces, the old build is live — have them redeploy before debugging further.

**Security caution:** never log full OAuth/token JSON response bodies — redact token/secret/code fields in any response-logging middleware.
