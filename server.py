import os
import json
import httpx
from msal import ConfidentialClientApplication
from mcp.server.fastmcp import FastMCP
from mcp.server.transport_security import TransportSecuritySettings

# Configuration
TENANT_ID = os.environ["AZURE_TENANT_ID"]
CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
PORT = int(os.environ.get("PORT", 8000))

# Webhook URLs for sending messages to channels (set via env vars)
# Format: WEBHOOK_<CHANNEL_KEY>=<webhook_url>
WEBHOOK_URLS = {}
for key, val in os.environ.items():
    if key.startswith("WEBHOOK_"):
        channel_key = key[8:].lower().replace("_", " ")
        WEBHOOK_URLS[channel_key] = val

# Disable DNS rebinding protection for Railway reverse proxy
security_settings = TransportSecuritySettings(
    enable_dns_rebinding_protection=False,
)

mcp = FastMCP(
    "Microsoft Teams",
    host="0.0.0.0",
    port=PORT,
    json_response=True,
    stateless_http=True,
    transport_security=security_settings,
)

_app = ConfidentialClientApplication(
    CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    client_credential=CLIENT_SECRET,
)


def get_token() -> str:
    result = _app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description')}")
    return result["access_token"]


def graph_headers() -> dict:
    return {
        "Authorization": f"Bearer {get_token()}",
        "Content-Type": "application/json",
    }


async def _resolve_team_id(client, team_name):
    resp = await client.get(
        f"{GRAPH_BASE}/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team') and displayName eq '{team_name}'",
        headers=graph_headers(),
    )
    resp.raise_for_status()
    teams = resp.json().get("value", [])
    if not teams:
        raise ValueError(f"Team '{team_name}' not found")
    return teams[0]["id"]


async def _resolve_channel_id(client, team_id, channel_name):
    resp = await client.get(
        f"{GRAPH_BASE}/teams/{team_id}/channels",
        headers=graph_headers(),
    )
    resp.raise_for_status()
    channels = resp.json().get("value", [])
    for ch in channels:
        if ch.get("displayName", "").lower() == channel_name.lower():
            return ch["id"]
    raise ValueError(f"Channel '{channel_name}' not found in team")


@mcp.tool()
async def list_teams() -> str:
    """List all Microsoft Teams the app has access to."""
    async with httpx.AsyncClient() as client:
        resp = await client.get(
            f"{GRAPH_BASE}/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,description",
            headers=graph_headers(),
        )
        if resp.status_code != 200:
            return json.dumps({"error": resp.status_code, "detail": resp.text})
        teams = resp.json().get("value", [])
        result = [{"id": t["id"], "name": t.get("displayName"), "description": t.get("description")} for t in teams]
        return json.dumps(result, indent=2)


@mcp.tool()
async def list_channels(team_id: str) -> str:
    """List all channels in a team.
    Args:
        team_id: The ID of the team
    """
    async with httpx.AsyncClient() as client:
        resp = await client.get(f"{GRAPH_BASE}/teams/{team_id}/channels", headers=graph_headers())
        if resp.status_code != 200:
            return json.dumps({"error": resp.status_code, "detail": resp.text})
        channels = resp.json().get("value", [])
        result = [{"id": ch["id"], "name": ch.get("displayName"), "description": ch.get("description")} for ch in channels]
        return json.dumps(result, indent=2)


@mcp.tool()
async def find_channel(channel_name: str, team_name: str = "") -> str:
    """Find a channel by name across all teams or within a specific team.
    Args:
        channel_name: Name of the channel to find (case-insensitive)
        team_name: Optional team name to narrow search
    """
    async with httpx.AsyncClient() as client:
        if team_name:
            try:
                tid = await _resolve_team_id(client, team_name)
            except ValueError as e:
                return json.dumps({"error": str(e)})
            teams_to_search = [{"id": tid, "displayName": team_name}]
        else:
            resp = await client.get(
                f"{GRAPH_BASE}/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName",
                headers=graph_headers(),
            )
            if resp.status_code != 200:
                return json.dumps({"error": resp.status_code, "detail": resp.text})
            teams_to_search = resp.json().get("value", [])
        results = []
        for team in teams_to_search:
            resp = await client.get(f"{GRAPH_BASE}/teams/{team['id']}/channels", headers=graph_headers())
            if resp.status_code != 200:
                continue
            for ch in resp.json().get("value", []):
                if channel_name.lower() in ch.get("displayName", "").lower():
                    results.append({
                        "team_id": team["id"], "team_name": team.get("displayName"),
                        "channel_id": ch["id"], "channel_name": ch.get("displayName"),
                    })
        if not results:
            return json.dumps({"error": f"No channel matching '{channel_name}' found"})
        return json.dumps(results, indent=2)


@mcp.tool()
async def read_messages(team_id: str, channel_id: str, top: int = 20) -> str:
    """Read recent messages from a Teams channel.
    Args:
        team_id: The ID of the team
        channel_id: The ID of the channel
        top: Number of messages to retrieve (default 20)
    """
    async with httpx.AsyncClient() as client:
        resp = await client.get(
            f"{GRAPH_BASE}/teams/{team_id}/channels/{channel_id}/messages?$top={top}",
            headers=graph_headers(),
        )
        if resp.status_code != 200:
            return json.dumps({"error": resp.status_code, "detail": resp.text})
        messages = resp.json().get("value", [])
        result = []
        for m in messages:
            result.append({
                "id": m.get("id"),
                "from": m.get("from", {}).get("user", {}).get("displayName", "Unknown") if m.get("from") else "System",
                "body": m.get("body", {}).get("content", ""),
                "createdDateTime": m.get("createdDateTime"),
            })
        return json.dumps(result, indent=2)


@mcp.tool()
async def read_replies(team_id: str, channel_id: str, message_id: str) -> str:
    """Read replies to a specific message.
    Args:
        team_id: The ID of the team
        channel_id: The ID of the channel
        message_id: The ID of the parent message
    """
    async with httpx.AsyncClient() as client:
        resp = await client.get(
            f"{GRAPH_BASE}/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies",
            headers=graph_headers(),
        )
        if resp.status_code != 200:
            return json.dumps({"error": resp.status_code, "detail": resp.text})
        replies = resp.json().get("value", [])
        result = []
        for r in replies:
            result.append({
                "id": r.get("id"),
                "from": r.get("from", {}).get("user", {}).get("displayName", "Unknown") if r.get("from") else "System",
                "body": r.get("body", {}).get("content", ""),
                "createdDateTime": r.get("createdDateTime"),
            })
        return json.dumps(result, indent=2)


@mcp.tool()
async def list_members(team_id: str) -> str:
    """List members of a team.
    Args:
        team_id: The ID of the team
    """
    async with httpx.AsyncClient() as client:
        resp = await client.get(f"{GRAPH_BASE}/teams/{team_id}/members", headers=graph_headers())
        if resp.status_code != 200:
            return json.dumps({"error": resp.status_code, "detail": resp.text})
        return json.dumps(resp.json().get("value", []), indent=2)


@mcp.tool()
async def read_chat_messages(chat_id: str, top: int = 20) -> str:
    """Read messages from a 1:1 or group chat.
    Args:
        chat_id: The ID of the chat
        top: Number of messages to retrieve (default 20)
    """
    async with httpx.AsyncClient() as client:
        resp = await client.get(
            f"{GRAPH_BASE}/chats/{chat_id}/messages?$top={top}",
            headers=graph_headers(),
        )
        if resp.status_code != 200:
            return json.dumps({"error": resp.status_code, "detail": resp.text})
        return json.dumps(resp.json().get("value", []), indent=2)


# WRITE TOOLS (Power Automate webhooks - Graph API app perms are migration-only)

@mcp.tool()
async def list_webhook_channels() -> str:
    """List all channels that have webhook URLs configured for sending messages."""
    if not WEBHOOK_URLS:
        return json.dumps({"error": "No webhook URLs configured. Add WEBHOOK_<CHANNEL_KEY> env vars."})
    return json.dumps([{"channel_key": k, "configured": True} for k in WEBHOOK_URLS], indent=2)


@mcp.tool()
async def send_message(channel_key: str, message: str) -> str:
    """Send a message to a Teams channel via Power Automate webhook.
    Args:
        channel_key: Channel key matching a configured webhook (e.g. 'it helpdesk', 'general', 'team lead reports')
        message: The message content (plain text or HTML)
    """
    key = channel_key.lower()
    webhook_url = WEBHOOK_URLS.get(key)
    if not webhook_url:
        for wk, wurl in WEBHOOK_URLS.items():
            if key in wk or wk in key:
                webhook_url = wurl
                key = wk
                break
    if not webhook_url:
        return json.dumps({
            "error": f"No webhook configured for '{channel_key}'",
            "available_channels": list(WEBHOOK_URLS.keys()),
        })
    async with httpx.AsyncClient() as client:
        resp = await client.post(
            webhook_url,
            content=message,
            headers={"Content-Type": "text/plain"},
            timeout=30.0,
        )
        if resp.status_code not in (200, 202):
            return json.dumps({"error": resp.status_code, "detail": resp.text})
        return json.dumps({"status": "sent", "channel": key})


@mcp.tool()
async def send_message_by_name(team_name: str, channel_name: str, message: str) -> str:
    """Send a message to a Teams channel by team and channel name.
    Args:
        team_name: The display name of the team
        channel_name: The display name of the channel
        message: The message content (plain text or HTML)
    """
    key = channel_name.lower()
    if key not in WEBHOOK_URLS:
        for wk in WEBHOOK_URLS:
            if key in wk or wk in key:
                key = wk
                break
    return await send_message(key, message)


@mcp.tool()
async def create_channel(team_id: str, display_name: str, description: str = "") -> str:
    """Create a new channel in a team.
    Args:
        team_id: The ID of the team
        display_name: Name of the new channel
        description: Optional description
    """
    async with httpx.AsyncClient() as client:
        resp = await client.post(
            f"{GRAPH_BASE}/teams/{team_id}/channels",
            headers=graph_headers(),
            json={"displayName": display_name, "description": description},
        )
        if resp.status_code not in (200, 201):
            return json.dumps({"error": resp.status_code, "detail": resp.text})
        data = resp.json()
        return json.dumps({"status": "created", "channelId": data.get("id"), "displayName": data.get("displayName")})


@mcp.tool()
async def delete_channel(team_id: str, channel_id: str) -> str:
    """Delete a channel from a team.
    Args:
        team_id: The ID of the team
        channel_id: The ID of the channel to delete
    """
    async with httpx.AsyncClient() as client:
        resp = await client.delete(f"{GRAPH_BASE}/teams/{team_id}/channels/{channel_id}", headers=graph_headers())
        if resp.status_code not in (200, 204):
            return json.dumps({"error": resp.status_code, "detail": resp.text})
        return json.dumps({"status": "deleted", "channelId": channel_id})


@mcp.tool()
async def create_team(display_name: str, description: str = "") -> str:
    """Create a new Microsoft Team.
    Args:
        display_name: Name of the new team
        description: Optional description
    """
    body = {
        "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
        "displayName": display_name,
        "description": description,
    }
    async with httpx.AsyncClient() as client:
        resp = await client.post(f"{GRAPH_BASE}/teams", headers=graph_headers(), json=body)
        if resp.status_code == 202:
            return json.dumps({"status": "creating", "location": resp.headers.get("Location", "")})
        if resp.status_code not in (200, 201):
            return json.dumps({"error": resp.status_code, "detail": resp.text})
        return json.dumps({"status": "created"})


@mcp.tool()
async def add_team_member(team_id: str, user_id: str, role: str = "member") -> str:
    """Add a member to a team.
    Args:
        team_id: The ID of the team
        user_id: The ID of the user to add
        role: 'member' or 'owner'
    """
    body = {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        "roles": ["owner"] if role == "owner" else [],
        "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{user_id}')",
    }
    async with httpx.AsyncClient() as client:
        resp = await client.post(f"{GRAPH_BASE}/teams/{team_id}/members", headers=graph_headers(), json=body)
        if resp.status_code not in (200, 201):
            return json.dumps({"error": resp.status_code, "detail": resp.text})
        return json.dumps({"status": "added", "memberId": resp.json().get("id")})


@mcp.tool()
async def remove_team_member(team_id: str, membership_id: str) -> str:
    """Remove a member from a team.
    Args:
        team_id: The ID of the team
        membership_id: The membership ID of the member
    """
    async with httpx.AsyncClient() as client:
        resp = await client.delete(f"{GRAPH_BASE}/teams/{team_id}/members/{membership_id}", headers=graph_headers())
        if resp.status_code not in (200, 204):
            return json.dumps({"error": resp.status_code, "detail": resp.text})
        return json.dumps({"status": "removed", "membershipId": membership_id})


@mcp.tool()
async def update_channel(team_id: str, channel_id: str, display_name: str = "", description: str = "") -> str:
    """Update a channel's name or description.
    Args:
        team_id: The ID of the team
        channel_id: The ID of the channel
        display_name: New name (optional)
        description: New description (optional)
    """
    body = {}
    if display_name:
        body["displayName"] = display_name
    if description:
        body["description"] = description
    if not body:
        return json.dumps({"error": "Provide at least display_name or description"})
    async with httpx.AsyncClient() as client:
        resp = await client.patch(
            f"{GRAPH_BASE}/teams/{team_id}/channels/{channel_id}",
            headers=graph_headers(), json=body,
        )
        if resp.status_code != 200:
            return json.dumps({"error": resp.status_code, "detail": resp.text})
        data = resp.json()
        return json.dumps({"status": "updated", "channelId": data.get("id"), "displayName": data.get("displayName")})


if __name__ == "__main__":
    mcp.run(transport="streamable-http")

# Force redeploy with updated webhook env vars
