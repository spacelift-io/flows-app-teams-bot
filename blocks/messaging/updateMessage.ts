import { AppBlock, events } from "@slflows/sdk/v1";
import { getBotAccessToken, updateBotMessage } from "../../utils/botHandler.ts";

export const updateMessage: AppBlock = {
  name: "Update Message",
  description:
    "Updates an existing message in Teams with new text or Adaptive Card content.",
  category: "Messaging",

  inputs: {
    default: {
      name: "Update",
      description: "Trigger updating the message",
      config: {
        conversationId: {
          name: "Conversation ID",
          description:
            "The conversation ID containing the message (from Send Message output).",
          type: "string",
          required: true,
        },
        activityId: {
          name: "Activity ID",
          description:
            "The message activity ID to update (from Send Message output).",
          type: "string",
          required: true,
        },
        text: {
          name: "New Message Text",
          description:
            "The new text content of the message. Supports markdown formatting.",
          type: "string",
          required: false,
        },
        attachments: {
          name: "New Adaptive Cards",
          description:
            "Array of Adaptive Card JSON objects to replace existing content. Design cards at https://adaptivecards.microsoft.com/designer",
          type: {
            type: "array",
            items: {
              type: "object",
              description: "An Adaptive Card JSON object",
            },
          },
          required: false,
        },
      },
      async onEvent(input) {
        const { appId, appPassword, serviceUrl, tenantId } = input.app.config;
        const { conversationId, activityId, text, attachments } =
          input.event.inputConfig;

        // Get Bot Framework access token
        const accessToken = await getBotAccessToken(
          appId,
          appPassword,
          tenantId,
        );

        // Build the updated message activity
        const message: any = {
          type: "message",
        };

        if (text) {
          message.text = text;
          message.textFormat = "markdown";
        }

        if (attachments && attachments.length > 0) {
          message.attachments = attachments.map((card: any) => ({
            contentType: "application/vnd.microsoft.card.adaptive",
            content: card,
          }));
        }

        // Ensure we have either text or attachments
        if (!text && (!attachments || attachments.length === 0)) {
          throw new Error(
            "Updated message must have either text content or adaptive cards",
          );
        }

        // Update the message
        await updateBotMessage(
          serviceUrl,
          conversationId,
          activityId,
          message,
          accessToken,
        );

        await events.emit({
          activityId: activityId,
          conversationId: conversationId,
          timestamp: new Date().toISOString(),
        });
      },
    },
  },

  outputs: {
    default: {
      name: "Message Updated",
      description: "Emitted when the message has been successfully updated",
      possiblePrimaryParents: ["default"],
      type: {
        type: "object",
        properties: {
          activityId: {
            type: "string",
            description: "The ID of the updated message activity",
          },
          conversationId: {
            type: "string",
            description: "The conversation ID where the message was updated",
          },
          timestamp: {
            type: "string",
            description: "ISO 8601 timestamp when the message was updated",
          },
        },
        required: ["activityId", "conversationId", "timestamp"],
      },
    },
  },
};
