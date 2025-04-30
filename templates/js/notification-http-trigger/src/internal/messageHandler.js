const { teamsBot } = require("../teamsBot");
const { notificationApp } = require("./initialize");

module.exports = async function (context, req) {
  const res = {
    status: (code) => {
      context.res.status = code;
      return res;
    },
    send: (body) => {
      context.res.body = body;
      return res;
    },
    json: (body) => {
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
