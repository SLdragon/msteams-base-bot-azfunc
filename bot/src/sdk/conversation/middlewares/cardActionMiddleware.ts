import {
  Activity,
  ActivityTypes,
  Attachment,
  CardFactory,
  ConversationState,
  InvokeResponse,
  MemoryStorage,
  MessageFactory,
  Middleware,
  StatePropertyAccessor,
  TurnContext,
  UserState,
} from "botbuilder";
import { TeamsBotSsoPromptSettings } from "../../bot/teamsBotSsoPrompt";
import { TeamsBotSsoPromptTokenResponse } from "../../bot/teamsBotSsoPromptTokenResponse";
import { ErrorCode, ErrorMessage, ErrorWithCode } from "../../core/errors";
import { TeamsFx } from "../../core/teamsfx";
import { IdentityType } from "../../models/identityType";
import { internalLogger } from "../../util/logger";
import {
  AdaptiveCardResponse,
  BotSsoConfig,
  BotSsoExecutionActivityHandler,
  CommandMessage,
  InvokeResponseErrorCode,
  TeamsFxAdaptiveCardActionHandler,
  TeamsFxSsoAdaptiveCardActionHandler,
} from "../interface";
import { InvokeResponseFactory, InvokeResponseType } from "../invokeResponseFactory";
import { BotSsoExecutionDialog } from "../sso/botSsoExecutionDialog";

/**
 * @internal
 */
export class CardActionMiddleware implements Middleware {
  public readonly actionHandlers: (TeamsFxAdaptiveCardActionHandler |TeamsFxSsoAdaptiveCardActionHandler)[] = [];
  private readonly ssoActionHandlers: TeamsFxSsoAdaptiveCardActionHandler[] = [];

  private readonly defaultMessage: string = "Your response was sent to the app";
  public hasSsoCommand: boolean;

  private verifyStateOperationName = 'signin/verifyState';
  private tokenExchangeOperationName = 'signin/tokenExchange';

  private ssoExecutionDialog:BotSsoExecutionDialog;
  private conversationState:ConversationState;
  private userState:UserState;
  private dialogState: StatePropertyAccessor;
  private store: Map<string,string>

  constructor(handlers?: TeamsFxAdaptiveCardActionHandler[], ssoHandlers?: TeamsFxSsoAdaptiveCardActionHandler[],
    ssoConfig?: BotSsoConfig) {
    handlers = handlers ?? [];
    ssoHandlers = ssoHandlers ?? [];
    this.hasSsoCommand = ssoHandlers.length > 0;
    this.store = new Map<string,string>();
    // this.ssoActivityHandler = activityHandler;

    if (this.hasSsoCommand && !ssoConfig) {
      internalLogger.error(ErrorMessage.SsoActivityHandlerIsNull);
      throw new ErrorWithCode(
        ErrorMessage.SsoActivityHandlerIsNull,
        ErrorCode.SsoActivityHandlerIsUndefined
      );
    }

    const memoryStorage = new MemoryStorage();
    const userState = ssoConfig?.dialog?.userState ?? new UserState(memoryStorage);
    const conversationState =
      ssoConfig?.dialog?.conversationState ?? new ConversationState(memoryStorage);
    const dedupStorage = ssoConfig?.dialog?.dedupStorage ?? memoryStorage;
    const scopes = ssoConfig?.aad.scopes ?? [".default"];
    const settings: TeamsBotSsoPromptSettings = {
      scopes: scopes,
      timeout: ssoConfig?.dialog?.ssoPromptConfig?.timeout,
      endOnInvalidMessage: ssoConfig?.dialog?.ssoPromptConfig?.endOnInvalidMessage,
    };

    let teamsfx: TeamsFx;
    if (ssoConfig) {
      const { scopes, ...customConfig } = ssoConfig?.aad;
      teamsfx = new TeamsFx(IdentityType.User, { ...customConfig });
    } else {
      teamsfx = new TeamsFx(IdentityType.User);
    }

    this.ssoExecutionDialog = new BotSsoExecutionDialog(dedupStorage, settings, teamsfx);
    this.conversationState = conversationState;

    this.dialogState = conversationState.createProperty("DialogState");
    this.userState = userState;

    this.actionHandlers.push(...handlers);

    for (const ssoHandler of ssoHandlers) {
      this.addSsoCommand(ssoHandler);
    }
  }

  public addSsoCommand(ssoHandler: TeamsFxSsoAdaptiveCardActionHandler) {
    this.ssoExecutionDialog?.addCommand(
      async (
        context: TurnContext,
        tokenResponse: TeamsBotSsoPromptTokenResponse,
        message: CommandMessage
      ) => {
        const response = await ssoHandler.handleActionInvoked(context, message, tokenResponse);
        await this.processResponse(context, response, ssoHandler);
      },
      ssoHandler.triggerVerb
    );
    this.ssoActionHandlers.push(ssoHandler);
    this.actionHandlers.push(ssoHandler);
    this.hasSsoCommand = true;
  }

  private async processResponse(context: TurnContext, response: InvokeResponse<any>, handler: TeamsFxAdaptiveCardActionHandler |TeamsFxSsoAdaptiveCardActionHandler) {
    const responseType = response.body?.type;
    switch (responseType) {
      case InvokeResponseType.AdaptiveCard:
        const card = response.body?.value;
        if (!card) {
          const errorMessage = "Adaptive card content cannot be found in the response body";
          await this.sendInvokeResponse(
            context,
            InvokeResponseFactory.errorResponse(
              InvokeResponseErrorCode.InternalServerError,
              errorMessage
            )
          );
          throw new Error(errorMessage);
        }

        if (card.refresh && handler.adaptiveCardResponse !== AdaptiveCardResponse.NewForAll) {
          // Card won't be refreshed with AdaptiveCardResponse.ReplaceForInteractor.
          // So set to AdaptiveCardResponse.ReplaceForAll here.
          handler.adaptiveCardResponse = AdaptiveCardResponse.ReplaceForAll;
        }

        if (this.isSsoExecutionHandler(handler) && (handler.adaptiveCardResponse === AdaptiveCardResponse.ReplaceForInteractor||!handler.adaptiveCardResponse)) {
          // Bot sso only support one-on-one chat: https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/authentication/auth-aad-sso-bots
          handler.adaptiveCardResponse = AdaptiveCardResponse.ReplaceForAll;
        }

        const activity = MessageFactory.attachment(CardFactory.adaptiveCard(card));

        if (handler.adaptiveCardResponse === AdaptiveCardResponse.NewForAll) {
          await this.sendInvokeResponse(
            context,
            InvokeResponseFactory.textMessage(this.defaultMessage)
          );
          await context.sendActivity(activity);
        } else if (handler.adaptiveCardResponse === AdaptiveCardResponse.ReplaceForAll) {
          if (this.isSsoExecutionHandler(handler)) {
            activity.id = this.store.get(context.activity.from.aadObjectId);
          }
          else{
            activity.id = context.activity.replyToId;
          }
          await context.updateActivity(activity);
          await this.sendInvokeResponse(context, response);
        } else {
            await this.sendInvokeResponse(context, response, activity.id);
        }
        break;
      case InvokeResponseType.Message:
      case InvokeResponseType.Error:
      default:
        await this.sendInvokeResponse(context, response);
        break;
    }
  }

  async onTurn(context: TurnContext, next: () => Promise<void>): Promise<void> {
    if (context.activity.name === this.verifyStateOperationName || context.activity.name === this.tokenExchangeOperationName || context.activity.name === ActivityTypes.Message) {
      try {
        await this.ssoExecutionDialog.run(context, this.dialogState);
      }
      finally{
        await this.conversationState.saveChanges(context, false);
        await this.userState.saveChanges(context, false);
      }
    }
    else if (context.activity.name === "adaptiveCard/action") {
      const action = context.activity.value.action;
      const actionVerb = action.verb;

      for (const handler of this.actionHandlers) {
        if (handler.triggerVerb?.toLowerCase() === actionVerb?.toLowerCase()) {
          if (this.isSsoExecutionHandler(handler)) {
            await this.sendInvokeResponse(context, {
              status: 200,
              body: {
                statusCode: 200,
              },
            });
           
            try{
              await this.ssoExecutionDialog?.run(context,this.dialogState);
              await this.store.set(context.activity.from.aadObjectId, context.activity.replyToId);
            }
            finally{
              await this.conversationState.saveChanges(context, false);
              await this.userState.saveChanges(context, false);
            }
/*
            const card = CardFactory.oauthCard(
              "",
              "Teams SSO Sign In",
              "Sign In",
              `${process.env.INITIATE_LOGIN_ENDPOINT}?scope=User.Read&clientId=${process.env.M365_CLIENT_ID}&tenantId=${process.env.M365_TENANT_ID}&loginHint=rentu@vschinateamsdev.onmicrosoft.com`,
              {
                "id":process.env.M365_CLIENT_ID,
                "uri":process.env.M365_APPLICATION_ID_URI
              }
            );

            
              {
                "statusCode": 401,
                "value": {
                    "code": "401",
                    "message": "Unexpected response type for Status Code: 401 Expected responseTypes: application/vnd.microsoft.activity.loginRequest or application/vnd.microsoft.error.invalidAuthCode"
                },
                "type": "application/vnd.microsoft.error",
                "text": null,
                "attachmentLayout": null,
                "channelData": null
              }
            
            await this.sendInvokeResponse(context, {
              status: 200,
              body: {
                statusCode: 401,
                type: "application/vnd.microsoft.activity.loginRequest",
                value:card.content
              },
            });
*/

/*
{
    "statusCode": 412,
    "value": {
        "code": "412",
        "message": "Unexpected response type for Status Code: 412 Expected responseType: application/vnd.microsoft.error.preconditionFailed"
    },
    "type": "application/vnd.microsoft.error",
    "text": null,
    "attachmentLayout": null,
    "channelData": null
}


            await this.sendInvokeResponse(context, {
              status: 200,
              body: {
                statusCode: 412,
                type: "application/vnd.microsoft.error.preconditionFailed",
              },
            });
*/

/**works fine 

            await this.sendInvokeResponse(context, {
              status: 200,
              body: {
                statusCode: 200,
                type: "application/vnd.microsoft.activity.message",
                value: "my"
              }
            })
*/
          } else {
            let response: InvokeResponse;
            try {
              response = await (handler as TeamsFxAdaptiveCardActionHandler).handleActionInvoked(context, action.data);
            } catch (error: any) {
              const errorResponse = InvokeResponseFactory.errorResponse(
                InvokeResponseErrorCode.InternalServerError,
                error.message
              );
              await this.sendInvokeResponse(context, errorResponse);
              throw error;
            }
            await this.processResponse(context, response, handler);
  
          }
          break;
        }
      }
    }

    await next();
  }

  private isSsoExecutionHandler(
    handler: TeamsFxAdaptiveCardActionHandler |TeamsFxSsoAdaptiveCardActionHandler
  ): boolean {
    return this.ssoActionHandlers.indexOf(handler) >= 0;
  }

  private async sendInvokeResponse(context: TurnContext, response: InvokeResponse, replyId?: string) {
    const activity :Partial<Activity> = {
      type: ActivityTypes.InvokeResponse,
      value: response,
    };
    if (replyId) {
      activity.id = replyId;
      context.activity.id = replyId;
    }
    await context.sendActivity(activity);
  }
}
