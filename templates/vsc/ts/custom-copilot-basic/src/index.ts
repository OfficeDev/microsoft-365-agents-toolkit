// Import the Teams AI v2 app
import app from "./app/app";

// Start the application - the App class handles the server hosting
(async () => {
  console.log(`\nAgent started, app listening to`, process.env.PORT || 3978);
  await app.start(+(process.env.PORT || 3978));
})();
