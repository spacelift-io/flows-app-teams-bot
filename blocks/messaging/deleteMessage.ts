import { AppBlock, events } from "@slflows/sdk/v1";
import { getBotAccessToken, deleteBotMessage } from "../../utils/botHandler.ts";

export const deleteMessage: AppBlock = {
  name: "Delete Message",
  description: "Deletes a message from a Teams conversation.",
  category: "Messaging",

  inputs: {
    default: {
      name: "Delete",
      description: "Trigger deleting the message",
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
            "The message activity ID to delete (from Send Message output).",
          type: "string",
          required: true,
        },
      },
      async onEvent(input) {
        const { appId, appPassword, serviceUrl, tenantId } = input.app.config;
        const { conversationId, activityId } = input.event.inputConfig;

        // Get Bot Framework access token
        const accessToken = await getBotAccessToken(
          appId,
          appPassword,
          tenantId,
        );

        // Delete the message
        await deleteBotMessage(
          serviceUrl,
          conversationId,
          activityId,
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
      name: "Message Deleted",
      description: "Emitted when the message has been successfully deleted",
      possiblePrimaryParents: ["default"],
      type: {
        type: "object",
        properties: {
          activityId: {
            type: "string",
            description: "The ID of the deleted message activity",
          },
          conversationId: {
            type: "string",
            description: "The conversation ID where the message was deleted",
          },
          timestamp: {
            type: "string",
            description: "ISO 8601 timestamp when the message was deleted",
          },
        },
        required: ["activityId", "conversationId", "timestamp"],
      },
    },
  },
};
