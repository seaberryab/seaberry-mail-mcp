#!/usr/bin/env bash
# setup.sh — provision the Entra app registration for seaberry-mail-mcp.
#
# Prerequisite: signed into `az` as a tenant admin in the target tenant.
# Idempotent: safe to re-run. Reuses an existing app by display name and
# replaces its permission set atomically — so re-running corrects any prior mistakes.
#
# Output: writes ~/.config/mail-mcp/config.json with
#   { ClientId, TenantId, AllowedInviteDomains, ArchiveFolderName }.
# If config.json already exists, AllowedInviteDomains and ArchiveFolderName are preserved.

set -euo pipefail

APP_DISPLAY_NAME="local-mail-mcp"
REDIRECT_URI="http://localhost"
GRAPH_RESOURCE_ID="00000003-0000-0000-c000-000000000000"

# Scope NAMES are the source of truth. IDs are resolved dynamically from the live
# Graph service principal (avoids the entire class of "I guessed the GUID wrong" bug).
REQUIRED_SCOPES=("Mail.ReadWrite" "Calendars.ReadWrite" "Tasks.ReadWrite")

CONFIG_DIR="$HOME/.config/mail-mcp"
CONFIG_FILE="$CONFIG_DIR/config.json"

log() { printf '\033[36m[setup]\033[0m %s\n' "$*" >&2; }
err() { printf '\033[31m[setup]\033[0m %s\n' "$*" >&2; }

require() {
  command -v "$1" >/dev/null 2>&1 || { err "missing required command: $1"; exit 1; }
}

require az
require jq

# Verify az login.
if ! az account show >/dev/null 2>&1; then
  err "not signed into az. run: az login"
  exit 1
fi

TENANT_ID="$(az account show --query tenantId -o tsv)"
log "tenant: $TENANT_ID"

# ─── Resolve scope IDs from the live Graph SP (authoritative) ────────────────
log "resolving Graph delegated scope IDs from the live service principal"

declare -A SCOPE_ID
for scope_name in "${REQUIRED_SCOPES[@]}"; do
  scope_id="$(az ad sp show --id "$GRAPH_RESOURCE_ID" \
    --query "oauth2PermissionScopes[?value=='${scope_name}'].id | [0]" -o tsv)"
  if [[ -z "$scope_id" || "$scope_id" == "null" ]]; then
    err "could not resolve delegated scope id for '$scope_name'"
    exit 1
  fi
  SCOPE_ID[$scope_name]="$scope_id"
  log "  $scope_name → $scope_id"
done

# ─── Find or create the app ──────────────────────────────────────────────────
EXISTING_APP_ID="$(az ad app list --display-name "$APP_DISPLAY_NAME" --query '[0].appId' -o tsv 2>/dev/null || true)"

if [[ -n "$EXISTING_APP_ID" ]]; then
  APP_ID="$EXISTING_APP_ID"
  log "reusing existing app '$APP_DISPLAY_NAME' (appId=$APP_ID)"
else
  log "creating app '$APP_DISPLAY_NAME'"
  APP_ID="$(az ad app create \
    --display-name "$APP_DISPLAY_NAME" \
    --sign-in-audience AzureADMyOrg \
    --is-fallback-public-client true \
    --public-client-redirect-uris "$REDIRECT_URI" \
    --query appId -o tsv)"
  log "created appId=$APP_ID"
fi

# Ensure the SP exists. Harmless if it already does.
az ad sp create --id "$APP_ID" >/dev/null 2>&1 || true

# ─── Atomically REPLACE the required-resource-accesses ───────────────────────
# This is the key to self-healing: whatever permissions were there (right or wrong)
# get overwritten with exactly the set we want. No need to track-and-remove orphans.
log "replacing app permissions with: ${REQUIRED_SCOPES[*]}"

RESOURCE_ACCESS_JSON="$(jq -n \
  --arg graphId "$GRAPH_RESOURCE_ID" \
  --argjson scopes "$(jq -n '[]')" '
    {resourceAppId: $graphId, resourceAccess: $scopes}
  ')"

for scope_name in "${REQUIRED_SCOPES[@]}"; do
  RESOURCE_ACCESS_JSON="$(jq \
    --arg id "${SCOPE_ID[$scope_name]}" \
    '.resourceAccess += [{id: $id, type: "Scope"}]' \
    <<< "$RESOURCE_ACCESS_JSON")"
done

REQUIRED_ACCESSES="$(jq -n --argjson r "$RESOURCE_ACCESS_JSON" '[$r]')"

az ad app update --id "$APP_ID" --required-resource-accesses "$REQUIRED_ACCESSES" >/dev/null
log "app permissions set"

# ─── Admin-consent (with retry for SP replication race) ──────────────────────
log "granting admin consent (with retry for SP replication)"
CONSENTED=0
for attempt in 1 2 3 4 5; do
  if az ad app permission admin-consent --id "$APP_ID" >/dev/null 2>&1; then
    CONSENTED=1
    log "admin consent granted (attempt $attempt)"
    break
  fi
  log "admin-consent attempt $attempt failed, retrying in $((attempt * 2))s"
  sleep "$((attempt * 2))"
done

if [[ $CONSENTED -eq 0 ]]; then
  err "admin-consent failed after 5 attempts. confirm you are signed in as a Global Admin."
  err "you can retry manually: az ad app permission admin-consent --id $APP_ID"
  exit 1
fi

# ─── Write config.json (preserve user-customized AllowedInviteDomains) ───────
mkdir -p "$CONFIG_DIR"
chmod 700 "$CONFIG_DIR"

if [[ -f "$CONFIG_FILE" ]]; then
  # Preserve whatever the user has configured. Don't clobber a manually-widened list
  # or a renamed archive folder.
  EXISTING_DOMAINS="$(jq -c '.AllowedInviteDomains // []' "$CONFIG_FILE" 2>/dev/null || echo '[]')"
  EXISTING_ARCHIVE="$(jq -r '.ArchiveFolderName // "Archived by AI"' "$CONFIG_FILE" 2>/dev/null || echo 'Archived by AI')"
else
  # Fresh install: empty allowlist. create_meeting will refuse every attendee until
  # the user explicitly opts in by editing AllowedInviteDomains in config.json.
  EXISTING_DOMAINS='[]'
  EXISTING_ARCHIVE='Archived by AI'
fi

jq -n \
  --arg clientId "$APP_ID" \
  --arg tenantId "$TENANT_ID" \
  --argjson allowedDomains "$EXISTING_DOMAINS" \
  --arg archiveFolder "$EXISTING_ARCHIVE" \
  '{ClientId: $clientId, TenantId: $tenantId, AllowedInviteDomains: $allowedDomains, ArchiveFolderName: $archiveFolder}' \
  > "$CONFIG_FILE"
chmod 600 "$CONFIG_FILE"

log "wrote $CONFIG_FILE"
echo
echo "ClientId: $APP_ID"
echo "TenantId: $TENANT_ID"
echo "Scopes:   ${REQUIRED_SCOPES[*]}"
echo
log "done. next: register the MCP with Claude Code and run ensure_setup to complete sign-in."
