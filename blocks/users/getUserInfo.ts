import { AppBlock, events } from "@slflows/sdk/v1";
import { getBotAccessToken } from "../../utils/botHandler.ts";

export const getUserInfo: AppBlock = {
  name: "Get User Info",
  description:
    "Fetches user details (name, email, Azure AD info) from Teams using their user ID. Useful for resolving user information from reaction or mention events.",
  category: "Users",

  inputs: {
    default: {
      name: "Get User",
      description: "Fetch user information",
      config: {
        userId: {
          name: "User ID",
          description:
            "The Teams user ID to look up (from reaction events, mention events, or reply events).",
          type: "string",
          required: true,
        },
        conversationId: {
          name: "Conversation ID",
          description:
            "The conversation ID where this user is a member (from the same event as the user ID).",
          type: "string",
          required: true,
        },
      },
      async onEvent(input) {
        const { appId, appPassword, serviceUrl, tenantId } = input.app.config;
        const { userId, conversationId } = input.event.inputConfig;

        // Get Bot Framework access token
        const accessToken = await getBotAccessToken(
          appId,
          appPassword,
          tenantId,
        );

        // Fetch user details from Bot Framework API
        const url = `${serviceUrl}/v3/conversations/${encodeURIComponent(conversationId)}/members/${encodeURIComponent(userId)}`;
        const response = await fetch(url, {
          method: "GET",
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        });

        if (!response.ok) {
          const error = await response.text();
          throw new Error(`Failed to fetch user info: ${error}`);
        }

        const member = await response.json();

        await events.emit({
          name: member.name,
          email: member.email,
          aadObjectId: member.aadObjectId,
          userPrincipalName: member.userPrincipalName,
        });
      },
    },
  },

  outputs: {
    default: {
      name: "User Info",
      description: "User details from Teams",
      type: {
        type: "object",
        properties: {
          name: {
            type: "string",
            description: "The display name of the user",
          },
          email: {
            type: "string",
            description: "The user's email address",
          },
          aadObjectId: {
            type: "string",
            description: "The Azure AD object ID",
          },
          userPrincipalName: {
            type: "string",
            description: "The user principal name (UPN)",
          },
        },
        required: ["name", "aadObjectId"],
      },
    },
  },
};
