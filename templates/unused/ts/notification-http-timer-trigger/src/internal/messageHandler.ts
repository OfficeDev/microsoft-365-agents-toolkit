import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import { teamsBot } from "../teamsBot";
import { adapter } from "./initialize";

const httpTrigger: AzureFunction = async function (
  context: Context,
  req: HttpRequest
): Promise<any> {
  let status = 200;
  let return_body: unknown = null;
  const res = {
    status: (code: number) => {
      status = code;
      context.res.status = code;
    },
    send: (body: unknown) => {
      return_body = body;
    },
    setHeader: () => {},
    end: () => {},
  };
  await adapter.process(req, res as any, async (context) => {
    await teamsBot.run(context);
  });
  context.res = {
    status,
    body: return_body,
  };
  return return_body;
};

export default httpTrigger;
