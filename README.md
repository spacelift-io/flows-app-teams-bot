# Teams Bot

A Microsoft Teams Bot app for interactive messaging and conversation management using the Bot Framework.

## Features

- **Rich Messaging**: Send messages with markdown and Adaptive Cards to channels
- **Message Management**: Update and delete bot messages
- **Conversation Monitoring**: Subscribe to @mentions and thread replies
- **User Information**: Fetch user details from Teams
- **Interactive Events**: Receive reactions and Adaptive Card action events
- **Proactive Messaging**: Send messages to channels without user prompt

## Relationship to Teams Management App

This app handles **interactive bot operations** via the Bot Framework:

- Sending messages to channels
- Responding to @mentions
- Monitoring thread replies
- Receiving reaction and button click events
- Real-time bot interaction handling

For **administrative operations** (channel management, user management, reading message history), use the separate **Teams Management** app which leverages Microsoft Graph API.

## Getting Started

### Prerequisites

- Azure account with permissions to create Azure Bot resources
- Microsoft Teams workspace where you can install custom apps

### Setup

See the app's built-in installation wizard for step-by-step setup instructions. The process includes:

1. Create Azure Bot resource
2. Get credentials (App ID, Client Secret, Tenant ID)
3. Configure the app
4. Set messaging endpoint in Azure
5. Enable Teams channel
6. Download and install the pre-configured Teams app package

### Development scripts

```bash
npm install
npm run typecheck  # Type checking
npm run format     # Code formatting
```

## Configuration

Required settings:

- **Microsoft App ID**: Application ID from Azure Bot registration
- **App Password**: Client secret from Azure Bot registration
- **Service URL**: Bot Framework service URL (default works for all regions)
- **Tenant ID**: Azure AD tenant ID where the bot is registered

The bot automatically validates credentials and populates Bot ID and Bot Name signals during sync.

## Blocks

### Messaging

#### Send Message

Sends a message to a Teams channel with markdown text and/or Adaptive Cards. Optionally reply in a thread by providing a conversation ID.

**Key Inputs:**

- **Channel ID**: Teams channel ID (starts with `19:`)
- **Message Text**: Markdown-formatted text (optional)
- **Adaptive Cards**: Array of card JSON from [Adaptive Cards Designer](https://adaptivecards.microsoft.com/designer) (optional)
- **Conversation ID**: For threaded replies (optional)

**Outputs:**

- Activity ID, Conversation ID, Channel ID, Timestamp
- **Events output**: Emits reaction and Adaptive Card action events on the sent message

#### Update Message

Updates an existing bot message with new text or Adaptive Cards.

**Key Inputs:**

- **Conversation ID**: From Send Message output
- **Activity ID**: From Send Message output
- **New Text**: Updated markdown text (optional)
- **New Adaptive Cards**: Replacement cards from [designer](https://adaptivecards.microsoft.com/designer) (optional)

#### Delete Message

Deletes a bot message from a conversation.

**Key Inputs:**

- **Conversation ID**: From Send Message output
- **Activity ID**: From Send Message output

### Subscriptions

#### Mentions Subscription

Subscribes to @mentions of the bot in Teams.

**Config:**

- **Channel ID Filter**: Optional. Only receive mentions from specific channel.

**Outputs:**

- Mention text, user info, channel ID, conversation ID, activity ID for replying

#### Subscribe to Replies

Subscribes to replies in a conversation thread.

**Inputs:**

- **Conversation ID**: Thread to monitor (from Send Message or Mention event)

**Outputs:**

- Reply text, attachments, user info, activity ID, timestamp

### Users

#### Get User Info

Fetches user details (name, email, Azure AD info) from Teams.

**Inputs:**

- **User ID**: From reaction/mention/reply events
- **Conversation ID**: From the same event

**Outputs:**

- Name, email, Azure AD object ID, user principal name

## Installation

The app includes a built-in installation wizard with complete setup instructions. The wizard provides:

- Step-by-step Azure Bot registration guide
- Credential gathering instructions
- Automatic messaging endpoint URL generation
- Pre-configured Teams app package download (manifest + icons)
- Teams channel setup instructions

Simply follow the wizard to complete the setup process.

## Example Usage

### Responding to Mentions

1. Add **Mentions Subscription** block to receive @mention events
2. Connect to **Send Message** block
3. Pass `conversationId` and `channelId` from mention event to reply in the thread

### Monitoring Thread Replies

1. Send a message with **Send Message** block
2. Add **Subscribe to Replies** block
3. Connect the `conversationId` output to the subscribe input
4. Receive events for each reply in the thread

### Handling Reactions and Card Actions

1. Send a message with **Send Message** block
2. Connect the **Events** output (secondary) to handle:
   - User reactions (like, heart, etc.)
   - Adaptive Card button clicks with submitted data

### Creating Adaptive Cards

1. Design your card at [Adaptive Cards Designer](https://adaptivecards.microsoft.com/designer)
2. Copy the card JSON
3. Pass it in the **Adaptive Cards** input of **Send Message**
4. The app automatically wraps it with the proper `contentType`

### Proactive Channel Messaging

1. Get channel ID from Teams (right-click channel → Get link to channel)
2. Use **Send Message** with the channel ID
3. Leave `conversationId` empty to post a new message (not a reply)

## Architecture

### Block Organization

```
blocks/
├── messaging/
│   ├── sendMessage.ts       # Send messages (new or threaded)
│   ├── updateMessage.ts     # Update existing messages
│   └── deleteMessage.ts     # Delete messages
├── subscriptions/
│   ├── mentionsSubscription.ts     # @mention events
│   └── subscribeToReplies.ts       # Thread reply events
└── users/
    └── getUserInfo.ts       # Fetch user details
```

### Key Components

- **`utils/botHandler.ts`**: Bot Framework activity routing and filtering
- **Message helpers**: Functions for sending, updating, and deleting messages
- **Token management**: Bot Framework authentication
- **Event routing**: Routes reactions and actions back to Send Message blocks

### Design Patterns

- **Event-driven subscriptions**: Blocks use `onInternalMessage` to receive Bot Framework activities
- **Config-based filtering**: Subscription blocks filter by channel/conversation ID
- **Automatic wrapping**: Adaptive Cards automatically get proper `contentType`
- **Multi-output support**: Send Message has default output + secondary events output

## Resources

- [Bot Framework Documentation](https://docs.microsoft.com/azure/bot-service/)
- [Teams Bot Development](https://docs.microsoft.com/microsoftteams/platform/bots/what-are-bots)
- [Adaptive Cards Designer](https://adaptivecards.microsoft.com/designer)
- [Teams Management App](../flows-app-teams-management/) - For administrative operations via Graph API

## Development

See [CLAUDE.md](./CLAUDE.md) for development guidelines and patterns.
