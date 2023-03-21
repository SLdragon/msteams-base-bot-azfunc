import { ConversationBot, DefaultBotSsoExecutionActivityHandler } from "../sdk";
import { FormSubmitActionHandler } from "../actions/formSubmit";
import { ProfileSsoActionHandler } from "../actions/profileSsoActionHandler";
import { FormCommandHandler } from "../commands/form";
import { BlobsStorage } from "./blobsStorage";
import { conversationState, userState } from "./state";
import { ProfileHandler } from "../commands/profile";

const storage = new BlobsStorage(
  process.env.BLOB_CONNECTION_STRING,
  process.env.BLOB_CONTAINER_NAME_NOTIFICATIONS
);
// initialise the bot
export const bot = new ConversationBot({
  adapterConfig: {
    appId: process.env.BOT_ID,
    appPassword: process.env.BOT_PASSWORD,
  },
  notification: {
    enabled: true,
    storage: storage,
  },
  ssoConfig: {
    aad :{
      scopes:["User.Read"],
      clientId: process.env.M365_CLIENT_ID,
      clientSecret: process.env.M365_CLIENT_SECRET,
      tenantId: process.env.M365_TENANT_ID,
      authorityHost: process.env.M365_AUTHORITY_HOST,
      initiateLoginEndpoint: process.env.INITIATE_LOGIN_ENDPOINT.replace("auth-start.html", "api/public/auth-start.html"),
      applicationIdUri: process.env.M365_APPLICATION_ID_URI
    },
    dialog: {
      conversationState: conversationState,
      userState: userState,
      ssoPromptConfig: {
        timeout: 900000,
        endOnInvalidMessage: true
      }
    }
  },
  command: {
    enabled: true,
    commands: [
      new FormCommandHandler(),
      new ProfileHandler()
    ]
  },
  cardAction: {
    enabled: true,
    actions: [
      new FormSubmitActionHandler()
    ],
    ssoActions: [new ProfileSsoActionHandler()]
  }
});
