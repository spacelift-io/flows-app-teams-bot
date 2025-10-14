import { defineApp, http } from "@slflows/sdk/v1";
import JSZip from "jszip";
import { blocks } from "./blocks/index.ts";
import {
  handleBotActivity,
  verifyBotFrameworkRequest,
} from "./utils/botHandler.ts";

export const app = defineApp({
  name: "Teams Bot",
  installationInstructions: `
# Microsoft Teams Bot Setup

This app enables interactive bot capabilities for Microsoft Teams using the Bot Framework.

**Note:** For administrative operations (reading message history, managing channels), use the separate Teams Management app.

## Prerequisites

- Azure account with permissions to create Azure Bot resources
- Microsoft Teams workspace where you can install custom apps

## Setup Instructions

### 1. Create Azure Bot Resource

1. Go to [Azure Portal](https://portal.azure.com) → **Create a resource** → Search for **Azure Bot**
2. Configure the bot:
   - **Type of App**: Single Tenant
   - **Microsoft App ID**: Create new
3. Click **Review + Create**, then **Create**

### 2. Get Credentials

1. In the bot resource, go to **Configuration** → Copy the **Microsoft App ID**
2. Click **Manage** next to Microsoft App ID → **Certificates & secrets**
3. Create a new client secret and copy the **Value** (not the Secret ID)
4. Go to **Azure Active Directory** and copy your **Tenant ID**

### 3. Configure This App

Fill in the configuration below with your credentials, then click **Confirm**.

### 4. Set Azure Bot Messaging Endpoint

After confirming the configuration above, your messaging endpoint will be:

**\`{appEndpointUrl}/messages\`**

Go to your Azure Bot resource → **Configuration** → Set **Messaging endpoint** to this URL → **Apply**

### 5. Enable Teams Channel

In Azure Bot resource → **Channels** → Click **Microsoft Teams** icon → Accept terms → **Apply**

### 6. Install in Teams

Download the
**[pre-configured Teams app package]({appEndpointUrl}/manifest.zip)**

Upload to Teams: **Apps** → **Manage your apps** → **Upload an app** → Select the downloaded zip file
`,

  config: {
    appId: {
      name: "Microsoft App ID",
      description:
        "The Microsoft App ID (Application ID) for your Azure Bot registration.",
      type: "string",
      required: true,
      sensitive: false,
    },
    appPassword: {
      name: "App Password",
      description:
        "The client secret (app password) created for your Azure Bot registration.",
      type: "string",
      required: true,
      sensitive: true,
    },
    serviceUrl: {
      name: "Service URL",
      description:
        "The Bot Framework service URL for Teams. The default value works for all regions.",
      type: "string",
      required: true,
      default: "https://smba.trafficmanager.net/teams/",
    },
    tenantId: {
      name: "Tenant ID",
      description: "The Azure AD tenant ID where your bot is registered.",
      type: "string",
      required: true,
      sensitive: false,
    },
  },

  signals: {
    botId: {
      name: "Bot ID",
      description: "The ID of the bot user in Teams",
    },
    botName: {
      name: "Bot Name",
      description: "The name of the bot",
    },
  },

  async onSync(input) {
    const { appId, appPassword, tenantId } = input.app.config;

    // Validate credentials by attempting to get a token
    try {
      // Use tenant-specific endpoint for single-tenant apps
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
        console.error("Bot Framework auth failed:", error);
        return {
          newStatus: "failed",
          customStatusDescription: `Authentication failed: ${error.error_description || "check credentials"}`,
        };
      }

      await response.json(); // Validate token works

      return {
        newStatus: "ready",
        signalUpdates: {
          botId: appId,
          botName: "Flows Bot",
        },
      };
    } catch (error: any) {
      console.error("Error during bot authentication:", error.message);
      return {
        newStatus: "failed",
        customStatusDescription: `Authentication error: ${error.message}`,
      };
    }
  },

  http: {
    async onRequest(input) {
      const requestPath = input.request.path;

      // Bot Framework messages endpoint
      if (requestPath === "/messages" || requestPath.endsWith("/messages")) {
        const { appId, appPassword } = input.app.config;

        // Verify Bot Framework JWT token in Authorization header
        // HTTP headers are case-insensitive, check both common casings
        const authHeader =
          input.request.headers["authorization"] ||
          input.request.headers["Authorization"];
        if (!authHeader) {
          console.error(
            "Missing authorization header. Received headers:",
            Object.keys(input.request.headers),
          );
          await http.respond(input.request.requestId, {
            statusCode: 401,
            body: { error: "Missing authorization" },
          });
          return;
        }

        // Verify the request is from Bot Framework
        const isValid = await verifyBotFrameworkRequest(authHeader, appId);
        if (!isValid) {
          console.error("Invalid Bot Framework authorization");
          await http.respond(input.request.requestId, {
            statusCode: 403,
            body: { error: "Invalid authorization" },
          });
          return;
        }

        // Handle the bot activity
        const activity = input.request.body;
        await handleBotActivity(activity, appId, appPassword);

        // Bot Framework expects 200 OK response
        await http.respond(input.request.requestId, {
          statusCode: 200,
          body: {},
        });
        return;
      }

      // Teams app manifest download endpoint
      if (
        requestPath === "/manifest.zip" ||
        requestPath.endsWith("/manifest.zip")
      ) {
        const { appId } = input.app.config;

        // Generate manifest.json
        const manifest = {
          $schema:
            "https://developer.microsoft.com/json-schemas/teams/v1.22/MicrosoftTeams.schema.json",
          manifestVersion: "1.22",
          version: "1.0.0",
          id: appId,
          developer: {
            name: "Flows Bot",
            websiteUrl: input.app.installationUrl,
            privacyUrl: input.app.installationUrl,
            termsOfUseUrl: input.app.installationUrl,
          },
          name: {
            short: "Flows Bot",
            full: "Flows Teams Bot",
          },
          description: {
            short: "Interactive bot for Teams",
            full: "Bot for interactive conversations and automation in Microsoft Teams",
          },
          icons: {
            outline: "outline.png",
            color: "color.png",
          },
          accentColor: "#2E2D2D",
          bots: [
            {
              botId: appId,
              scopes: ["personal", "team", "groupChat"],
              supportsFiles: false,
              isNotificationOnly: false,
              commandLists: [
                {
                  scopes: ["personal", "team", "groupChat"],
                  commands: [],
                },
              ],
            },
          ],
          permissions: ["identity", "messageTeamMembers"],
          validDomains: [],
          webApplicationInfo: {
            id: appId,
            resource: `api://botid-${appId}`,
          },
          authorization: {
            permissions: {
              resourceSpecific: [
                {
                  name: "ChannelMessage.Read.Group",
                  type: "Application",
                },
                {
                  name: "ChatMessage.Read.Chat",
                  type: "Application",
                },
              ],
            },
          },
        };

        // Download PNG icons from GitHub
        const colorIconUrl =
          "https://raw.githubusercontent.com/spacelift-io/flows-app-teams-bot/main/icons/color.png";
        const outlineIconUrl =
          "https://raw.githubusercontent.com/spacelift-io/flows-app-teams-bot/main/icons/outline.png";

        const [colorIconResponse, outlineIconResponse] = await Promise.all([
          fetch(colorIconUrl),
          fetch(outlineIconUrl),
        ]);

        if (!colorIconResponse.ok || !outlineIconResponse.ok) {
          throw new Error("Failed to download icons from GitHub");
        }

        const colorIconBuffer = await colorIconResponse.arrayBuffer();
        const outlineIconBuffer = await outlineIconResponse.arrayBuffer();

        // Create ZIP file with manifest and icons
        const zip = new JSZip();
        zip.file("manifest.json", JSON.stringify(manifest, null, 2));
        zip.file("color.png", colorIconBuffer);
        zip.file("outline.png", outlineIconBuffer);

        const zipUint8Array = await zip.generateAsync({
          type: "uint8array",
          compression: "DEFLATE",
        });

        await http.respond(input.request.requestId, {
          statusCode: 200,
          headers: {
            "Content-Type": "application/octet-stream",
            "Content-Disposition":
              "attachment; filename=teams-bot-manifest.zip",
            "Content-Encoding": "identity",
          },
          body: zipUint8Array,
        });
        return;
      }

      // Unknown endpoint
      console.warn("Received request on unhandled HTTP path:", requestPath);
      await http.respond(input.request.requestId, {
        statusCode: 404,
        body: { error: "Endpoint not found" },
      });
    },
  },

  blocks,
});
