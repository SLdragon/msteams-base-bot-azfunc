import { AdaptiveCards } from "@microsoft/adaptivecards-tools";
import { CommandMessage, TeamsFxBotCommandHandler, TriggerPatterns } from "../sdk";
import { TurnContext, Activity, MessageFactory, CardFactory } from "botbuilder";
import formCard from "../cards/profile.json";
import { FormCard } from "../models/formCard";

export class ProfileHandler implements TeamsFxBotCommandHandler {
  triggerPatterns: TriggerPatterns = "profile";

  // handle the command
  async handleCommandReceived(context: TurnContext, message: CommandMessage): Promise<string | void | Partial<Activity>> {
    const card = AdaptiveCards.declare<FormCard>(formCard).render({
      title: "Use sso to show your profile",
    });
    
    // send the card
    await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(card)));
  }
}