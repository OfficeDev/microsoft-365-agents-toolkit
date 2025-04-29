import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import { teamsBot } from "../teamsBot";
import { notificationApp } from "./initialize";

const httpTrigger: AzureFunction = async function (
  context: Context,
  req: HttpRequest
): Promise<any> {
  const res = {
    status: (code: number) => {
      context.res.status = code;
      return res;
    },
    send: (body: any) => {
      context.res.body = body;
      return res;
    },
    json: (body: any) => {
      context.res.body = body;
      return res;
    },
    body: context.res.body,
  };
  await notificationApp.requestHandler(req, res, async (context) => {
    await teamsBot.run(context);
  });
  return res.body;
};

export default httpTrigger;
