import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import * as fs from "fs";
import * as path from "path";

const httpTrigger: AzureFunction = async function (
  context: Context,
  req: HttpRequest
): Promise<any> {
  let body;
  if(req.url.includes("auth-start")) {
    body = fs.readFileSync(path.resolve(__dirname, "../public/auth-start.html"), "utf-8");
  }
  else if(req.url.includes("auth-end")){
    body = fs.readFileSync(path.resolve(__dirname, "../public/auth-end.html"), "utf-8");
  }

  var res = {
    body,
    headers: {
        "Content-Type": "text/html"
    }
  };

  context.res = res;
};

export default httpTrigger;
