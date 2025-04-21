{
  "name": "{{SafeProjectNameLowerCase}}",
  "version": "1.0.0",
  "msteams": {
    "teamsAppId": null
  },
  "description": "Microsoft 365 Agents Toolkit message extension Bot sample",
  "engines": {
    "node": "18 || 20 || 22"
  },
  "author": "Microsoft",
  "license": "MIT",
  "main": "./src/index.js",
  "scripts": {
    "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run dev",
    "dev:teamsfx:testtool": "env-cmd --silent -f .localConfigs.testTool npm run dev",
    "dev:teamsfx:launch-testtool": "env-cmd --silent -f env/.env.testtool teamsapptester start",
    "dev": "nodemon --inspect=9239 --signal SIGINT ./src/index.js",
    "start": "node .src/index.js",
    "watch": "nodemon ./src/index.js",
    "test": "echo \"Error: no test specified\" && exit 1"
  },
  "dependencies": {
    "@microsoft/agents-hosting": "^0.1.49",
    "@microsoft/agents-hosting-teams": "^0.1.49",
    "adaptive-expressions": "^4.20.0",
    "adaptivecards-templating": "^2.3.1",
    "adaptivecards": "^3.0.1",
    "isomorphic-fetch": "^3.0.0",
    "express": "^5.0.1"
  },
  "devDependencies": {
    "env-cmd": "^10.1.0",
    "nodemon": "^3.1.7"
  }
}