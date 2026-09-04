#!/usr/bin/env bash
# Register the Microsoft Entra app that imap-mcp-server uses for Outlook.com /
# Microsoft 365 OAuth 2.0 (device-code flow), using the Azure CLI.
#
# Equivalent to the manual steps in README.md ("Outlook.com / Microsoft 365"):
#   New registration → personal + org accounts → no redirect URI →
#   Allow public client flows = Yes → API permissions: Office 365 Exchange Online,
#   delegated IMAP.AccessAsUser.All + SMTP.Send.
#
# Usage:   ./scripts/register-entra-app.sh [display-name]
# Needs:   az (https://learn.microsoft.com/cli/azure/install-azure-cli)
# Prints:  the Application (client) ID to put in IMAP_MCP_MS_CLIENT_ID.
set -euo pipefail

NAME="${1:-imap-mcp-server}"
EXO_APP_ID="00000002-0000-0ff1-ce00-000000000000"   # Office 365 Exchange Online
# Well-known delegated scope IDs on that resource (used only if the tenant
# cannot be queried for them, e.g. a brand-new personal-account tenant).
IMAP_SCOPE_FALLBACK="652390e4-393a-48de-9484-05f9b1212954"   # IMAP.AccessAsUser.All
SMTP_SCOPE_FALLBACK="258f6531-6087-4cc4-bb90-092c5fb3ed3f"   # SMTP.Send

command -v az >/dev/null || { echo "az CLI not found — install it first: brew install azure-cli" >&2; exit 1; }

if ! az account show >/dev/null 2>&1; then
  echo "Signing in. Use the Microsoft account that owns the mailbox (a personal" >&2
  echo "account is fine — it gets a small default tenant on first sign-in)." >&2
  az login --allow-no-subscriptions --use-device-code >/dev/null
fi

echo "Creating app registration \"$NAME\" ..." >&2
APP_ID=$(az ad app create \
  --display-name "$NAME" \
  --sign-in-audience AzureADandPersonalMicrosoftAccount \
  --is-fallback-public-client true \
  --query appId -o tsv)

# Look the scope IDs up on the Exchange Online service principal when it exists
# in this tenant; otherwise fall back to the well-known IDs.
lookup_scope() {
  az ad sp show --id "$EXO_APP_ID" \
    --query "oauth2PermissionScopes[?value=='$1'].id | [0]" -o tsv 2>/dev/null || true
}
IMAP_SCOPE=$(lookup_scope IMAP.AccessAsUser.All); IMAP_SCOPE="${IMAP_SCOPE:-$IMAP_SCOPE_FALLBACK}"
SMTP_SCOPE=$(lookup_scope SMTP.Send);            SMTP_SCOPE="${SMTP_SCOPE:-$SMTP_SCOPE_FALLBACK}"

echo "Adding delegated permissions IMAP.AccessAsUser.All + SMTP.Send ..." >&2
az ad app permission add --id "$APP_ID" --api "$EXO_APP_ID" \
  --api-permissions "${IMAP_SCOPE}=Scope" "${SMTP_SCOPE}=Scope" >/dev/null

cat <<EOF

Done. Application (client) ID:

  $APP_ID

Put it in your MCP client config, e.g.:

  "env": { "IMAP_MCP_MS_CLIENT_ID": "$APP_ID" }

then ask the assistant to add your Outlook.com mailbox (imap_add_oauth_account →
sign in at https://microsoft.com/devicelogin → imap_complete_oauth_login).
The permissions are consented by you at that first sign-in; no admin grant needed
for a personal account.
EOF
