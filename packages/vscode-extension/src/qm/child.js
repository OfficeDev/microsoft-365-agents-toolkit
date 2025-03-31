const SwaggerParser = require("@apidevtools/swagger-parser");
process.on("message", async (data) => {
  console.log(data);
  // Simulate a long-running synchronous task
  //   for (let i = 0; i < 10000000000000; i++) {
  //         // Do nothing
  //     }
  let res;
  try {
    let api = await SwaggerParser.dereference(data);
    res = JSON.stringify(api);
  } catch (err) {
    console.error(err);
    res = err;
  }

  process.send("Result of long-running task" + " " + res);
});
