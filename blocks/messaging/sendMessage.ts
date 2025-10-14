import { AppBlock, events, kv } from "@slflows/sdk/v1";
import {
  getBotAccessToken,
  sendMessageToChannel,
} from "../../utils/botHandler.ts";

export const sendMessage: AppBlock = {
  name: "Send Message",
  description:
    "Sends a message to a Teams channel. Optionally reply to a specific message by providing replyToId to create a thread.",
  category: "Messaging",

  inputs: {
    default: {
      name: "Send",
      description: "Trigger sending the message",
      config: {
        channelId: {
          name: "Channel ID",
          description:
            "The Teams channel ID (starts with 19:). Get this from channel settings or from a subscription event.",
          type: "string",
          required: true,
        },
        text: {
          name: "Message Text",
          description:
            "The text content of the message. Supports markdown formatting.",
          type: "string",
          required: false,
        },
        attachments: {
          name: "Adaptive Cards",
          description:
            "Array of Adaptive Card JSON objects. Design cards at https://adaptivecards.microsoft.com/designer and paste the card JSON here. Each card should have type, version, and body properties.",
          type: {
            type: "array",
            items: {
              type: "object",
              description: "An Adaptive Card JSON object",
            },
          },
          required: false,
        },
        conversationId: {
          name: "Conversation ID",
          description:
            "Optional. To reply in a thread, provide the conversation ID from a mention event or previous message. Leave empty to post a new message in the channel.",
          type: "string",
          required: false,
        },
      },
      async onEvent(input) {
        const { appId, appPassword, serviceUrl, tenantId } = input.app.config;
        const { botId, botName } = input.app.signals;
        const { channelId, text, attachments, conversationId } =
          input.event.inputConfig;

        // Ensure we have either text or attachments
        if (!text && (!attachments || attachments.length === 0)) {
          throw new Error(
            "Message must have either text content or adaptive cards",
          );
        }

        // Get Bot Framework access token
        const accessToken = await getBotAccessToken(
          appId,
          appPassword,
          tenantId,
        );

        // Prepare attachments with Adaptive Card wrapper
        const formattedAttachments = attachments
          ? attachments.map((card: any) => ({
              contentType: "application/vnd.microsoft.card.adaptive",
              content: card,
            }))
          : undefined;

        // Send message to channel
        const result = await sendMessageToChannel(
          serviceUrl,
          channelId,
          botId,
          botName || "Flows Bot",
          tenantId,
          {
            text,
            attachments: formattedAttachments,
            conversationId,
          },
          accessToken,
        );

        const eventId = input.event.id;

        // Store subscription for reactions and actions on this message
        const eventsKey = `events|${result.activityId}|${input.block.id}`;
        await Promise.all([
          kv.block.set({ key: result.activityId, value: eventId }),
          kv.app.set({ key: eventsKey, value: true }),
        ]);

        await events.emit({
          activityId: result.activityId,
          conversationId: result.conversationId,
          channelId: channelId,
          timestamp: new Date().toISOString(),
        });
      },
    },
  },

  async onInternalMessage({ message }) {
    const data = message.body;

    const { value: parentEventId } = await kv.block.get(data.activityId);

    if (!parentEventId) {
      return;
    }

    const commonEventData = {
      type: data.type,
      userId: data.from?.id,
      userAadObjectId: data.from?.aadObjectId,
      activityId: data.activityId,
      conversationId: data.conversationId,
      timestamp: data.timestamp,
    };

    // Handle reactions
    if (data.type === "reaction") {
      await events.emit(
        {
          reactionType: data.reactionType,
          action: data.action,
          ...commonEventData,
        },
        { parentEventId, outputKey: "events" },
      );
    }

    // Handle actions
    if (data.type === "action") {
      await events.emit(
        {
          actionId: data.actionId,
          actionData: data.actionData,
          ...commonEventData,
        },
        { parentEventId, outputKey: "events" },
      );
    }
  },

  outputs: {
    default: {
      name: "Message Sent",
      description: "Emitted when the message has been successfully sent",
      default: true,
      type: {
        type: "object",
        properties: {
          activityId: {
            type: "string",
            description: "The ID of the sent message activity",
          },
          conversationId: {
            type: "string",
            description: "The conversation ID for the channel",
          },
          channelId: {
            type: "string",
            description: "The channel ID where the message was sent",
          },
          timestamp: {
            type: "string",
            description: "ISO 8601 timestamp when the message was sent",
          },
        },
        required: ["activityId", "conversationId", "channelId", "timestamp"],
      },
    },
    events: {
      name: "Events",
      description:
        "Emitted when users react to or interact with this message (reactions, button clicks)",
      secondary: true,
      type: {
        type: "object",
        required: [
          "type",
          "userId",
          "userAadObjectId",
          "activityId",
          "conversationId",
          "timestamp",
        ],
        properties: {
          type: {
            type: "string",
            description: "Type of event",
            enum: ["reaction", "action"],
          },
          userId: {
            type: "string",
            description: "The Teams user ID",
          },
          userAadObjectId: {
            type: "string",
            description: "The Azure AD object ID of the user",
          },
          activityId: {
            type: "string",
            description: "The activity ID of the message",
          },
          conversationId: {
            type: "string",
            description: "Conversation ID",
          },
          timestamp: {
            type: "string",
            description: "ISO 8601 timestamp of the event",
          },
        },
        oneOf: [
          {
            type: "object",
            properties: {
              reactionType: {
                type: "string",
                description:
                  "The type of reaction (e.g., 'like', 'heart', 'laugh')",
              },
              action: {
                type: "string",
                description: "Either 'add' or 'remove'",
                enum: ["add", "remove"],
              },
            },
            required: ["reactionType", "action"],
          },
          {
            type: "object",
            properties: {
              actionId: {
                type: "string",
                description:
                  "The ID of the action that was clicked (e.g., 'approve', 'reject')",
              },
              actionData: {
                type: "object",
                description:
                  "Data submitted with the action (from Input fields on the card)",
              },
            },
            required: ["actionId"],
          },
        ],
      },
    },
  },
};
