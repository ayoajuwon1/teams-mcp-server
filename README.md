# Teams MCP Server

MCP (Model Context Protocol) server for Microsoft Teams integration with Claude. Provides full **read and write** access to Teams via Microsoft Graph API.

## Tools Available

### Read Tools
- **list_teams** - List all teams in the organization
- **list_channels** - List channels in a team
- **read_messages** - Read channel messages
- **read_replies** - Read message replies
- **list_members** - List team members
- **read_chat_messages** - Read 1:1/group chat messages

### Write Tools
- **send_message** - Send a message to a channel
- **reply_to_message** - Reply to a channel message
- **create_channel** - Create a new channel
- **delete_channel** - Delete a channel
- **update_channel** - Update channel name/description
- **create_team** - Create a new team
- **add_team_member** - Add a member to a team
- **remove_team_member** - Remove a member from a team
- **send_chat_message** - Send a chat message

## Prerequisites

- Azure App Registration with Microsoft Graph API permissions
- Python 3.10+

## Azure Permissions Required

The app registration needs these **Application** permissions (with admin consent):

| Permission | Type | Description |
|---|---|---|
| Channel.Create | Application | Create channels |
| Channel.Delete.All | Application | Delete channels |
| Channel.ReadBasic.All | Application | Read channel info |
| ChannelMessage.Read.All | Application | Read channel messages |
| ChannelSettings.ReadWrite.All | Application | Read/write channel settings |
| Chat.ReadWrite.All | Application | Read/write chat messages |
| Group.ReadWrite.All | Application | Read/write groups (send messages) |
| Team.Create | Application | Create teams |
| Team.ReadBasic.All | Application | List teams |
| TeamMember.Read.All | Application | Read team members |
| TeamMember.ReadWrite.All | Application | Add/remove team members |
| User.Read.All | Application | Read user profiles |

## Deployment (Remote)

### Using Docker

```bash
docker build -t teams-mcp-server .
docker run -p 8000:8000 \
  -e AZURE_TENANT_ID=your-tenant-id \
  -e AZURE_CLIENT_ID=your-client-id \
  -e AZURE_CLIENT_SECRET=your-client-secret \
  teams-mcp-server
```

### Deploy to Render

1. Fork this repo
2. Create a new **Web Service** on [render.com](https://render.com)
3. Connect your GitHub repo
4. Set environment: **Docker**
5. Add environment variables: AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET
6. Deploy

### Deploy to Railway

1. Go to [railway.app](https://railway.app)
2. New Project > Deploy from GitHub Repo
3. Select this repository
4. Add environment variables in the Variables tab
5. Railway auto-detects the Dockerfile

## Connect to Claude

Once deployed, configure Claude MCP connector to use the server URL:

```
https://your-server-url.onrender.com/mcp
```

For Claude Desktop, add to your config:

```json
{
  "mcpServers": {
    "teams": {
      "url": "https://your-server-url.onrender.com/mcp"
    }
  }
}
```

## Environment Variables

| Variable | Description |
|---|---|
| AZURE_TENANT_ID | Azure AD Tenant ID |
| AZURE_CLIENT_ID | App Registration Client ID |
| AZURE_CLIENT_SECRET | App Registration Client Secret |
| PORT | Server port (default: 8000) |

## License

MIT
