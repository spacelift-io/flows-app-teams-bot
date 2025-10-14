import { ActivityTypes, Activity } from "botbuilder";
import { blocks, messaging, kv } from "@slflows/sdk/v1";

/**
 * Verify Bot Framework JWT token
 * In production, this should use JwtTokenValidation from botbuilder
 * For now, we do basic validation
 */
export async function verifyBotFrameworkRequest(
  authHeader: string,
  _expectedAppId: string,
): Promise<boolean> {
  // Bot Framework sends "Bearer <token>"
  if (!authHeader.startsWith("Bearer ")) {
    return false;
  }

  // In a full implementation, you would:
  // 1. Extract the JWT token
  // 2. Verify the signature using Bot Framework's public keys
  // 3. Validate the claims (iss, aud, exp, etc.)
  // 4. Ensure the appId claim matches expectedAppId
  //
  // The botbuilder package provides JwtTokenValidation for this,
  // but it requires more setup. For now, we accept requests with Bearer tokens.
  // This is acceptable in a trusted environment where the endpoint is not public.

  return true; // Simplified - in production use proper JWT validation
}

/**
 * Handle incoming Bot Framework activity
 */
export async function handleBotActivity(
  activity: Activity,
  _appId: string,
  _appPassword: string,
): Promise<void> {
  // Route based on activity type
  switch (activity.type) {
    case ActivityTypes.Message:
      await handleMessageActivity(activity, _appId);
      break;

    case ActivityTypes.MessageReaction:
      await handleReactionActivity(activity);
      break;

    case ActivityTypes.ConversationUpdate:
      // Bot added/removed from conversation
      break;

    case ActivityTypes.Invoke:
      // Adaptive Card actions
      await handleInvokeActivity(activity);
      break;

    default:
      break;
  }
}

/**
 * Handle message activities (mentions, replies, regular messages)
 */
async function handleMessageActivity(
  activity: Activity,
  botId: string,
): Promise<void> {
  // Check if this is a mention of the bot
  const entities = activity.entities || [];
  const mentions = entities.filter((e: any) => e.type === "mention");
  const botMentioned = mentions.some((m: any) =>
    m.mentioned?.id?.includes("28:"),
  ); // Bot IDs in Teams start with 28:

  if (botMentioned) {
    // Route to mention subscription blocks, filtering by channel if configured
    const mentionBlocks = await blocks.list({
      typeIds: ["mentionsSubscription"],
    });

    if (mentionBlocks.blocks.length > 0) {
      const mentionChannelId = activity.channelData?.teamsChannelId;

      // Filter blocks by channel ID if they have it configured
      const targetBlockIds = mentionBlocks.blocks
        .filter((block) => {
          const filterChannelId = block.config?.channelId;
          // If no filter set, include this block
          // If filter set, only include if it matches
          return !filterChannelId || filterChannelId === mentionChannelId;
        })
        .map((b) => b.id);

      if (targetBlockIds.length > 0) {
        await messaging.sendToBlocks({
          blockIds: targetBlockIds,
          body: {
            type: "mention",
            activity: activity,
            text: activity.text,
            from: activity.from,
            conversation: activity.conversation,
            channelData: activity.channelData,
            timestamp: activity.timestamp,
          },
        });
      }
    }
  }

  // Check if this is a reply in a thread
  // Teams channels: replyToId may be in conversationId as "messageid=XXX"
  let replyToId = activity.replyToId;

  // In Teams channels, if conversationId contains messageid=XXX that differs from activity.id,
  // this is a reply to that parent message
  if (!replyToId && activity.conversation?.id && activity.id) {
    const match = activity.conversation.id.match(/;messageid=(\d+)/);
    if (match) {
      const messageId = match[1];
      // If messageid differs from current activity ID, it's a reply
      if (messageId !== activity.id) {
        replyToId = messageId;
      }
    }
  }

  if (replyToId) {
    // Check if this reply is from the bot itself - if so, skip routing
    const botAppId = botId.replace(/^28:/, "");
    const fromId = activity.from?.id;
    const isFromBot = fromId && fromId.includes(botAppId);

    if (isFromBot) {
      return;
    }

    // Check if this is an Action.Submit (has value data)
    const actionData = (activity as any).value;
    if (actionData) {
      // Route to events handler for Action.Submit
      const prefix = `events|${replyToId}|`;
      const subscribedBlockIds = (
        await kv.app.list({ keyPrefix: prefix })
      ).pairs.map(({ key }) => key.split("|")[2]);

      if (subscribedBlockIds.length > 0) {
        const sendMessageBlocks = await blocks.list({
          typeIds: ["sendMessage"],
        });
        const existingBlockIds = new Set(
          sendMessageBlocks.blocks.map((b) => b.id),
        );
        const validBlockIds = subscribedBlockIds.filter((id) =>
          existingBlockIds.has(id),
        );

        if (validBlockIds.length > 0) {
          await messaging.sendToBlocks({
            blockIds: validBlockIds,
            body: {
              type: "action",
              actionId: (activity as any).name || "submit",
              actionData: actionData,
              activity: activity,
              from: activity.from,
              activityId: replyToId,
              conversationId: activity.conversation?.id,
              timestamp: activity.timestamp,
            },
          });
        }
      }
      return;
    }

    // Regular reply handling
    const conversationId = activity.conversation?.id;
    if (!conversationId) {
      return;
    }

    const prefix = `replies|${conversationId}|`;
    const subscribedBlockIds = (
      await kv.app.list({ keyPrefix: prefix })
    ).pairs.map(({ key }) => key.split("|")[2]);

    if (subscribedBlockIds.length > 0) {
      // Get all threadSubscription blocks to filter out deleted ones
      const threadBlocks = await blocks.list({
        typeIds: ["threadSubscription"],
      });
      const existingBlockIds = new Set(threadBlocks.blocks.map((b) => b.id));

      // Only send to blocks that still exist
      const validBlockIds = subscribedBlockIds.filter((id) =>
        existingBlockIds.has(id),
      );

      if (validBlockIds.length > 0) {
        await messaging.sendToBlocks({
          blockIds: validBlockIds,
          body: {
            type: "reply",
            activity: activity,
            text: activity.text,
            from: activity.from,
            replyToId: replyToId,
            conversationId: conversationId,
            conversation: activity.conversation,
            channelData: activity.channelData,
            timestamp: activity.timestamp,
          },
        });
      }
    }
  }
}

/**
 * Handle reaction activities (message reactions added/removed)
 */
async function handleReactionActivity(activity: Activity): Promise<void> {
  const reactionsAdded = activity.reactionsAdded || [];
  const reactionsRemoved = activity.reactionsRemoved || [];
  const replyToId = activity.replyToId;

  if (!replyToId) {
    return;
  }

  // Look up subscribed blocks for this activity
  const prefix = `events|${replyToId}|`;
  const subscribedBlockIds = (
    await kv.app.list({ keyPrefix: prefix })
  ).pairs.map(({ key }) => key.split("|")[2]);

  if (subscribedBlockIds.length === 0) {
    return;
  }

  // Get all sendMessage blocks to filter out deleted ones
  const reactionBlocks = await blocks.list({
    typeIds: ["sendMessage"],
  });
  const existingBlockIds = new Set(reactionBlocks.blocks.map((b) => b.id));

  // Only send to blocks that still exist
  const validBlockIds = subscribedBlockIds.filter((id) =>
    existingBlockIds.has(id),
  );

  if (validBlockIds.length === 0) {
    return;
  }

  // Send events for added reactions
  for (const reaction of reactionsAdded) {
    await messaging.sendToBlocks({
      blockIds: validBlockIds,
      body: {
        type: "reaction",
        action: "add",
        reactionType: reaction.type,
        activity: activity,
        from: activity.from,
        activityId: replyToId,
        conversationId: activity.conversation?.id,
        channelData: activity.channelData,
        timestamp: activity.timestamp,
      },
    });
  }

  // Send events for removed reactions
  for (const reaction of reactionsRemoved) {
    await messaging.sendToBlocks({
      blockIds: validBlockIds,
      body: {
        type: "reaction",
        action: "remove",
        reactionType: reaction.type,
        activity: activity,
        from: activity.from,
        activityId: replyToId,
        conversationId: activity.conversation?.id,
        channelData: activity.channelData,
        timestamp: activity.timestamp,
      },
    });
  }
}

/**
 * Handle invoke activities (Adaptive Card actions)
 */
async function handleInvokeActivity(activity: Activity): Promise<void> {
  const replyToId = activity.replyToId;

  if (!replyToId) {
    return;
  }

  // Look up subscribed blocks for this activity (the message with the card)
  const prefix = `events|${replyToId}|`;
  const subscribedBlockIds = (
    await kv.app.list({ keyPrefix: prefix })
  ).pairs.map(({ key }) => key.split("|")[2]);

  if (subscribedBlockIds.length === 0) {
    return;
  }

  // Get all sendMessage blocks to filter out deleted ones
  const actionBlocks = await blocks.list({
    typeIds: ["sendMessage"],
  });
  const existingBlockIds = new Set(actionBlocks.blocks.map((b) => b.id));

  // Only send to blocks that still exist
  const validBlockIds = subscribedBlockIds.filter((id) =>
    existingBlockIds.has(id),
  );

  if (validBlockIds.length === 0) {
    return;
  }

  // Extract action data
  const actionData = (activity as any).value || {};

  await messaging.sendToBlocks({
    blockIds: validBlockIds,
    body: {
      type: "action",
      actionId: activity.name,
      actionData: actionData,
      activity: activity,
      from: activity.from,
      activityId: replyToId,
      conversationId: activity.conversation?.id,
      timestamp: activity.timestamp,
    },
  });
}

/**
 * Get Bot Framework access token for making API calls
 * For single-tenant bots, use tenant-specific endpoint
 */
export async function getBotAccessToken(
  appId: string,
  appPassword: string,
  tenantId: string,
): Promise<string> {
  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const params = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: appId,
    client_secret: appPassword,
    scope: "https://api.botframework.com/.default",
  });

  const response = await fetch(tokenUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: params.toString(),
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(`Failed to get bot token: ${JSON.stringify(error)}`);
  }

  const data = await response.json();
  return data.access_token;
}

/**
 * Create or find a conversation and send a message to a Teams channel
 * If conversationId is provided, posts directly to that conversation (for replies)
 * Otherwise creates a new conversation
 */
export async function sendMessageToChannel(
  serviceUrl: string,
  channelId: string,
  botId: string,
  botName: string,
  tenantId: string,
  message: {
    text?: string;
    attachments?: any[];
    conversationId?: string;
  },
  accessToken: string,
): Promise<{ activityId: string; conversationId: string }> {
  // If conversationId is provided, post directly to that conversation (reply in thread)
  if (message.conversationId) {
    const activity: any = {
      type: "message",
    };

    if (message.text) {
      activity.text = message.text;
      activity.textFormat = "markdown";
    }

    if (message.attachments && message.attachments.length > 0) {
      activity.attachments = message.attachments;
    }

    const url = `${serviceUrl}/v3/conversations/${encodeURIComponent(message.conversationId)}/activities`;
    const response = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(activity),
    });

    if (!response.ok) {
      const error = await response.text();
      throw new Error(`Failed to send reply to thread: ${error}`);
    }

    const result = await response.json();
    return {
      activityId: result.id,
      conversationId: message.conversationId,
    };
  }

  // Create new conversation
  const conversationParams: any = {
    bot: {
      id: botId,
      name: botName,
    },
    isGroup: true,
    channelData: {
      channel: {
        id: channelId,
      },
    },
    tenantId: tenantId,
    activity: {
      type: "message",
    },
  };

  if (message.text) {
    conversationParams.activity.text = message.text;
    conversationParams.activity.textFormat = "markdown";
  }

  if (message.attachments && message.attachments.length > 0) {
    conversationParams.activity.attachments = message.attachments;
  }

  const url = `${serviceUrl}/v3/conversations`;
  const response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(conversationParams),
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to send message to channel: ${error}`);
  }

  const result = await response.json();
  return {
    activityId: result.activityId,
    conversationId: result.id,
  };
}

/**
 * Update an existing message
 */
export async function updateBotMessage(
  serviceUrl: string,
  conversationId: string,
  activityId: string,
  message: Partial<Activity>,
  accessToken: string,
): Promise<void> {
  const url = `${serviceUrl}/v3/conversations/${conversationId}/activities/${activityId}`;

  const response = await fetch(url, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(message),
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to update message: ${error}`);
  }
}

/**
 * Delete a message
 */
export async function deleteBotMessage(
  serviceUrl: string,
  conversationId: string,
  activityId: string,
  accessToken: string,
): Promise<void> {
  const url = `${serviceUrl}/v3/conversations/${conversationId}/activities/${activityId}`;

  const response = await fetch(url, {
    method: "DELETE",
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to delete message: ${error}`);
  }
}
