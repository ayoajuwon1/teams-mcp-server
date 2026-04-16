import os
import json
import httpx
import uvicorn
from msal import ConfidentialClientApplication
from mcp.server.fastmcp import FastMCP

# Configuration
TENANT_ID = os.environ["AZURE_TENANT_ID"]
CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
PORT = int(os.environ.get("PORT", 8000))

mcp = FastMCP("Microsoft Teams")

# Auth Helper
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


# ââââââââââââââââââââââââââââââââââââââââââââââ
# READ TOOLS
# ââââââââââââââââââââââââââââââââââââââââââââââ

@mcp.tool()
async def list_teams() -> str:
    """List all teams in the organization."""
    async with httpx.AsyncClient() as client:
        resp = await client.get(
            f"{GRAPH_BASE}/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,description",
            headers=graph_headers(),
        )
        resp.raise_for_status()
        teams = resp.json().get("value", [])
        return json.dumps(teams, indent=2)


@mcp.tool()
async def list_channels(team_id: str) -> str:
    """List all channels in a team.
    
    Args:
        team_id: The ID of the team
    """
    async with httpx.AsyncClient() as client:
        resp = await client.get(
            f"{GRAPH_BASE}/teams/{team_id}/channels",
            headers=graph_headers(),
        )
        resp.raise_for_status()
        channels = resp.json().get("value", [])
        return json.dumps(channels, indent=2)


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
        resp.raise_for_status()
        messages = resp.json().get("value", [])
        result = []
        for m in messages:
            result.append({
                "id": m.get("id"),
                "from": m.get("from", {}).get("user", {}).get("displayName", "Unknown"),
                "body": m.get("body", {}).get("content", ""),
                "createdDateTime": m.get("createdDateTime"),
            })
        return json.dumps(result, indent=2)


@mcp.tool()
async def read_replies(team_id: str, channel_id: str, message_id: str) -> str:
    """Read replies to a specific message in a Teams channel.
    
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
        resp.raise_for_status()
        replies = resp.json().get("value", [])
        result = []
        for r in replies:
            result.append({
                "id": r.get("id"),
                "from": r.get("from", {}).get("user", {}).get("displayName", "Unknown"),
                "body": r.get("body", {}).get("content", ""),
                "createdDateTime": r.get("createdDateTime"),
            })
        return json.dumps(result, indent=2)


@mcp.tool()
async def list_members(team_id: str) -> str:
    """List all members of a team.
    
    Args:
        team_id: The ID of the team
    """
    async with httpx.AsyncClient() as client:
        resp = await client.get(
            f"{GRAPH_BASE}/teams/{team_id}/members",
            headers=graph_headers(),
        )
        resp.raise_for_status()
        members = resp.json().get("value", [])
        result = []
        for m in members:
            result.append({
                "id": m.get("id"),
                "displayName": m.get("displayName"),
                "email": m.get("email"),
                "roles": m.get("roles", []),
            })
        return json.dumps(result, indent=2)


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
        resp.raise_for_status()
        messages = resp.json().get("value", [])
        result = []
        for m in messages:
            result.append({
                "id": m.get("id"),
                "from": m.get("from", {}).get("user", {}).get("displayName", "Unknown"),
                "body": m.get("body", {}).get("content", ""),
                "createdDateTime": m.get("createdDateTime"),
            })
        return json.dumps(result, indent=2)


# ââââââââââââââââââââââââââââââââââââââââââââââ
# WRITE TOOLS
# ââââââââââââââââââââââââââââââââââââââââââââââ

@mcp.tool()
async def send_message(team_id: str, channel_id: str, message: str) -> str:
    """Send a message to a Teams channel.
    
    Args:
        team_id: The ID of the team
        channel_id: The ID of the channel
        message: The message content (HTML supported)
    """
    async with httpx.AsyncClient() as client:
        resp = await client.post(
            f"{GRAPH_BASE}/teams/{team_id}/channels/{channel_id}/messages",
            headers=graph_headers(),
            json={"body": {"contentType": "html", "content": message}},
        )
        resp.raise_for_status()
        data = resp.json()
        return json.dumps({"status": "sent", "messageId": data.get("id")})


@mcp.tool()
async def reply_to_message(team_id: str, channel_id: str, message_id: str, reply: str) -> str:
    """Reply to a message in a Teams channel.
    
    Args:
        team_id: The ID of the team
        channel_id: The ID of the channel
        message_id: The ID of the message to reply to
        reply: The reply content (HTML supported)
    """
    async with httpx.AsyncClient() as client:
        resp = await client.post(
            f"{GRAPH_BASE}/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies",
            headers=graph_headers(),
            json={"body": {"contentType": "html", "content": reply}},
        )
        resp.raise_for_status()
        data = resp.json()
        return json.dumps({"status": "replied", "replyId": data.get("id")})


@mcp.tool()
async def create_channel(team_id: str, display_name: str, description: str = "", channel_type: str = "standard") -> str:
    """Create a new channel in a team.
    
    Args:
        team_id: The ID of the team
        display_name: Name of the new channel
        description: Optional description for the channel
        channel_type: Type of channel - 'standard' or 'private' (default: standard)
    """
    body = {
        "displayName": display_name,
        "description": description,
        "membershipType": channel_type,
    }
    async with httpx.AsyncClient() as client:
        resp = await client.post(
            f"{GRAPH_BASE}/teams/{team_id}/channels",
            headers=graph_headers(),
            json=body,
        )
        resp.raise_for_status()
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
        resp = await client.delete(
            f"{GRAPH_BASE}/teams/{team_id}/channels/{channel_id}",
            headers=graph_headers(),
        )
        resp.raise_for_status()
        return json.dumps({"status": "deleted", "channelId": channel_id})


@mcp.tool()
async def create_team(display_name: str, description: str = "", owner_id: str = "") -> str:
    """Create a new team.
    
    Args:
        display_name: Name of the new team
        description: Optional description
        owner_id: User ID of the team owner (required)
    """
    body = {
        "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
        "displayName": display_name,
        "description": description,
    }
    if owner_id:
        body["members"] = [
            {
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": ["owner"],
                "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{owner_id}')",
            }
        ]
    async with httpx.AsyncClient() as client:
        resp = await client.post(
            f"{GRAPH_BASE}/teams",
            headers=graph_headers(),
            json=body,
        )
        resp.raise_for_status()
        location = resp.headers.get("Location", "")
        return json.dumps({"status": "creating", "message": "Team creation initiated", "location": location})


@mcp.tool()
async def add_team_member(team_id: str, user_id: str, role: str = "member") -> str:
    """Add a member to a team.
    
    Args:
        team_id: The ID of the team
        user_id: The user ID to add
        role: Role for the user - 'member' or 'owner' (default: member)
    """
    body = {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        "roles": [role] if role == "owner" else [],
        "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{user_id}')",
    }
    async with httpx.AsyncClient() as client:
        resp = await client.post(
            f"{GRAPH_BASE}/teams/{team_id}/members",
            headers=graph_headers(),
            json=body,
        )
        resp.raise_for_status()
        data = resp.json()
        return json.dumps({"status": "added", "memberId": data.get("id"), "displayName": data.get("displayName")})


@mcp.tool()
async def remove_team_member(team_id: str, membership_id: str) -> str:
    """Remove a member from a team.
    
    Args:
        team_id: The ID of the team
        membership_id: The membership ID of the member to remove (from list_members)
    """
    async with httpx.AsyncClient() as client:
        resp = await client.delete(
            f"{GRAPH_BASE}/teams/{team_id}/members/{membership_id}",
            headers=graph_headers(),
        )
        resp.raise_for_status()
        return json.dumps({"status": "removed", "membershipId": membership_id})


@mcp.tool()
async def send_chat_message(chat_id: str, message: str) -> str:
    """Send a message in a 1:1 or group chat.
    
    Args:
        chat_id: The ID of the chat
        message: The message content (HTML supported)
    """
    async with httpx.AsyncClient() as client:
        resp = await client.post(
            f"{GRAPH_BASE}/chats/{chat_id}/messages",
            headers=graph_headers(),
            json={"body": {"contentType": "html", "content": message}},
        )
        resp.raise_for_status()
        data = resp.json()
        return json.dumps({"status": "sent", "messageId": data.get("id")})


@mcp.tool()
async def update_channel(team_id: str, channel_id: str, display_name: str = "", description: str = "") -> str:
    """Update a channel's name or description.
    
    Args:
        team_id: The ID of the team
        channel_id: The ID of the channel
        display_name: New name for the channel (optional)
        description: New description for the channel (optional)
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
            headers=graph_headers(),
            json=body,
        )
        resp.raise_for_status()
        data = resp.json()
        return json.dumps({"status": "updated", "channelId": data.get("id"), "displayName": data.get("displayName")})


# ââââââââââââââââââââââââââââââââââââââââââââââ
# SERVER ENTRY POINT (Streamable HTTP for remote deployment)
# ââââââââââââââââââââââââââââââââââââââââââââââ

app = mcp.streamable_http_app()

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=PORT)
