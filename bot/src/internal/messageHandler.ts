import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import { BotActivityHandler } from "./activityHandler";
import { bot } from "./initialize";
import { ResponseWrapper } from "./responseWrapper";
import { conversationState, userState } from "./state";

const httpTrigger: AzureFunction = async function (
  context: Context,
  req: HttpRequest
): Promise<any> {
  // create a bot activity handler
  const botActivityHandler = new BotActivityHandler(conversationState, userState);

  // process the request
  const res = new ResponseWrapper(context.res);
  try{
    await bot.adapter.process(req, res, (context) => botActivityHandler.run(context))
  }
  catch(err){
    if(err.message.includes("412")) {
      // Error message including "412" means it is waiting for user's consent, which is a normal process of SSO, sholdn't throw this error.
    }
    else {
      throw err;
    }
  }
  
  // return the response
  return res.body;
};

export default httpTrigger;
