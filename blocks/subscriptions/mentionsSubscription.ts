import { AppBlock, events } from "@slflows/sdk/v1";

export const mentionsSubscription: AppBlock = {
  name: "Mentions Subscription",
  description:
    "Subscribes to @mentions of the bot in Teams. Emits events with mention text, channel info, and conversation details for replying.",
  category: "Subscription",

  config: {
    channelId: {
      name: "Channel ID Filter",
      description:
        "Optional. Only receive mentions from this specific channel. Leave empty for all channels.",
      type: "string",
      required: false,
    },
  },

  async onInternalMessage({ message }) {
    const mentionData = message.body;

    await events.emit({
      text: mentionData.text,
      activityId: mentionData.activity?.id,
      channelId: mentionData.channelData?.teamsChannelId,
      serviceUrl: mentionData.activity?.serviceUrl,
      conversationId: mentionData.conversation?.id,
      from: mentionData.from,
      timestamp: mentionData.timestamp,
    });
  },

  outputs: {
    default: {
      name: "Mention Event",
      description: "Emitted when the bot is mentioned",
      type: {
        type: "object",
        properties: {
          text: {
            type: "string",
            description: "The mention message text",
          },
          activityId: {
            type: "string",
            description:
              "Activity ID - use this with Thread Subscription block to monitor replies",
          },
          channelId: {
            type: "string",
            description: "The Teams channel ID where the mention occurred",
          },
          serviceUrl: {
            type: "string",
            description: "Service URL for responding (use in Send Message)",
          },
          conversationId: {
            type: "string",
            description: "Conversation ID for responding (use in Send Message)",
          },
          from: {
            type: "object",
            description: "Information about the user who mentioned the bot",
          },
          timestamp: {
            type: "string",
            description: "ISO 8601 timestamp of the mention",
          },
        },
        required: [
          "text",
          "activityId",
          "channelId",
          "serviceUrl",
          "conversationId",
          "from",
          "timestamp",
        ],
      },
    },
  },
};
