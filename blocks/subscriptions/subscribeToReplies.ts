import { AppBlock, events, kv } from "@slflows/sdk/v1";

export const subscribeToReplies: AppBlock = {
  name: "Subscribe to Replies",
  description:
    "Subscribes to replies in a conversation thread. Emits events for each reply (excluding bot messages).",
  category: "Subscription",

  config: {},

  inputs: {
    subscribe: {
      name: "Subscribe",
      description: "Subscribe this block to thread replies",
      config: {
        conversationId: {
          name: "Conversation ID",
          description:
            "The conversation ID to monitor for replies (from Send Message output or Mention event).",
          type: "string",
          required: true,
        },
      },
      async onEvent(input) {
        const conversationId = input.event.inputConfig.conversationId;
        const appKey = `replies|${conversationId}|${input.block.id}`;

        await Promise.all([
          kv.block.set({ key: conversationId, value: input.event.id }),
          kv.app.set({ key: appKey, value: true }),
        ]);
      },
    },
  },

  async onInternalMessage({ message }) {
    const replyData = message.body;

    const { value: parentEventId } = await kv.block.get(
      replyData.conversationId,
    );

    if (!parentEventId) {
      return;
    }

    await events.emit(
      {
        text: replyData.text,
        attachments: replyData.activity?.attachments,
        userId: replyData.from?.id,
        userName: replyData.from?.name,
        userAadObjectId: replyData.from?.aadObjectId,
        activityId: replyData.activity?.id,
        timestamp: replyData.timestamp,
      },
      { parentEventId },
    );
  },

  outputs: {
    default: {
      name: "Reply Event",
      description: "Emitted when a reply is posted to the subscribed message",
      type: {
        type: "object",
        properties: {
          text: {
            type: "string",
            description: "The reply message text",
          },
          attachments: {
            type: "array",
            description:
              "Attachments in the reply (images, files, adaptive cards, etc.)",
            items: {
              type: "object",
            },
          },
          userId: {
            type: "string",
            description: "The Teams user ID who posted the reply",
          },
          userName: {
            type: "string",
            description: "The name of the user who posted the reply",
          },
          userAadObjectId: {
            type: "string",
            description:
              "The Azure AD object ID of the user who posted the reply",
          },
          activityId: {
            type: "string",
            description: "Activity ID of this reply",
          },
          timestamp: {
            type: "string",
            description: "ISO 8601 timestamp of the reply",
          },
        },
        required: [
          "text",
          "userId",
          "userName",
          "userAadObjectId",
          "activityId",
          "timestamp",
        ],
      },
    },
  },
};
