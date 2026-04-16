import os
import httpx
from msal import ConfidentialClientApplication
from mcp.server.fastmcp import FastMCP

# Configuration
TENANT_ID = os.environ["AZURE_TENANT_ID"]
CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

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


# READ TOOLS


@mcp.tool()
async def list_teams() -> str:
    """List all Teams the app has access to."""
    async with httpx.AsyncClient() as client:
        r = await client.get(
            f"{GRAPH_BASE}/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')",
            headers=graph_headers(),
        )
        r.raise_for_status()
        teams = r.json().get("value", [])
        return (
            "\n".join(f"- {t['displayName']} (id: {t['id']})" for t in teams)
            or "No teams found."
        )


@mcp.tool()
async def list_channels(team_id: str) -> str:
    """List all channels in a Team.

    Args:
        team_id: The ID of the team
    """
    async with httpx.AsyncClient() as client:
        r = await client.get(
            f"{GRAPH_BASE}/teams/{team_id}/channels",
            headers=graph_headers(),
        )
        r.raise_for_status()
        channels = r.json().get("value", [])
        return (
            "\n".join(
                f"- {c['displayName']} (id: {c['id']})" for c in channels
            )
            or "No channels found."
        )


@mcp.tool()
async def read_messages(team_id: str, channel_id: str, count: int = 10) -> str:
    """Read recent messages from a Teams channel.

    Args:
        team_id: The ID of the team
        channel_id: The ID of the channel
        count: Number of messages to retrieve (default 10)
    """
    async with httpx.AsyncClient() as client:
        r = await client.get(
            f"{GRAPH_BASE}/teams/{team_id}/channels/{channel_id}/messages?$top={count}",
            headers=graph_headers(),
        )
        r.raise_for_status()
        msgs = r.json().get("value", [])
        lines = []
        for m in msgs:
            sender = (
                m.get("from", {}).get("user", {}).get("displayName", "Unknown")
            )
            body = m.get("body", {}).get("content", "")[:200]
            lines.append(f"[{sender}]: {body}")
        return "\n\n".join(lines) or "No messages found."


@mcp.tool()
async def read_replies(
    team_id: str, channel_id: str, message_id: str
) -> str:
    """Read replies to a specific message thread.

    Args:
        team_id: The ID of the team
        channel_id: The ID of the channel
        message_id: The ID of the parent message
    """
    async with httpx.AsyncClient() as client:
        r = await client.get(
            f"{GRAPH_BASE}/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies",
            headers=graph_headers(),
        )
        r.raise_for_status()
        replies = r.json().get("value", [])
        lines = []
        for m in replies:
            sender = (
                m.get("from", {}).get("user", {}).get("displayName", "Unknown")
            )
            body = m.get("body", {}).get("content", "")[:200]
            lines.append(f"[{sender}]: {body}")
        return "\n\n".join(lines) or "No replies found."


@mcp.tool()
async def list_members(team_id: str) -> str:
    """List members of a Team.

    Args:
        team_id: The ID of the team
    """
    async with httpx.AsyncClient() as client:
        r = await client.get(
            f"{GRAPH_BASE}/teams/{team_id}/members",
            headers=graph_headers(),
        )
        r.raise_for_status()
        members = r.json().get("value", [])
        return (
            "\n".join(
                f"- {m.get('displayName', 'Unknown')} ({m.get('email', '')})"
                for m in members
            )
            or "No members found."
        )


# WRITE TOOLS


@mcp.tool()
async def send_message(team_id: str, channel_id: str, message: str) -> str:
    """Send a new message to a Teams channel.

    Args:
        team_id: The ID of the team
        channel_id: The ID of the channel
        message: The message content (supports HTML)
    """
    async with httpx.AsyncClient() as client:
        r = await client.post(
            f"{GRAPH_BASE}/teams/{team_id}/channels/{channel_id}/messages",
            headers=graph_headers(),
            json={"body": {"contentType": "html", "content": message}},
        )
        r.raise_for_status()
        return f"Message sent (id: {r.json().get('id')})"


@mcp.tool()
async def reply_to_message(
    team_id: str, channel_id: str, message_id: str, reply: str
) -> str:
    """Reply to an existing message thread.

    Args:
        team_id: The ID of the team
        channel_id: The ID of the channel
        message_id: The ID of the parent message to reply to
        reply: The reply content (supports HTML)
    """
    async with httpx.AsyncClient() as client:
        r = await client.post(
            f"{GRAPH_BASE}/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies",
            headers=graph_headers(),
            json={"body": {"contentType": "html", "content": reply}},
        )
        r.raise_for_status()
        return f"Reply sent (id: {r.json().get('id')})"


# Run
if __name__ == "__main__":
    mcp.run(transport="stdio")
