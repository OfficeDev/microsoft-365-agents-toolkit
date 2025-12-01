const express = require("express");
const fs = require("fs");
const https = require("https");
const path = require("path");
const send = require("send");

const app = express();

const sslOptions = {
  key: process.env.SSL_KEY_FILE ? fs.readFileSync(process.env.SSL_KEY_FILE) : undefined,
  cert: process.env.SSL_CRT_FILE ? fs.readFileSync(process.env.SSL_CRT_FILE) : undefined,
};

app.use("/static", express.static(path.join(__dirname, "static")));

// Adding tabs to our app. This will setup routes to various views
// Setup home page
app.get("/", (req, res) => {
  send(req, path.join(__dirname, "views", "hello.html")).pipe(res);
});

// Setup the static tab
app.get("/tab", (req, res) => {
  send(req, path.join(__dirname, "views", "hello.html")).pipe(res);
});

// Create HTTP server
const port = process.env.port || process.env.PORT || 3333;

if (sslOptions.key && sslOptions.cert) {
  https.createServer(sslOptions, app).listen(port, () => {
    console.log(`Express server listening on port ${port}`);
  });
} else {
  app.listen(port, () => {
    console.log(`Express server listening on port ${port}`);
  });
}
