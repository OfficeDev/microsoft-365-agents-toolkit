const { teamsBot } = require("../teamsBot");
const { notificationApp } = require("./initialize");

module.exports = async function (context, req) {
  let status = 200;
  let return_body = null;
  const res = {
    status: (code) => {
      status = code;
      context.res.status = code;
    },
    send: (body) => {
      return_body = body;
    },
    setHeader: () => {},
    end: () => {},
  };
  await notificationApp.requestHandler(req, res, async (context) => {
    await teamsBot.run(context);
  });
  context.res = {
    status,
    body: return_body,
  };
  return return_body;
};
