{
  "name": "{{SafeProjectNameLowerCase}}",
  "version": "1.0.0",
  "msteams": {
    "teamsAppId": null
  },
  "description": "Microsoft 365 Agents Toolkit hello world Bot sample",
  "engines": {
    "node": "18 || 20 || 22"
  },
  "author": "Microsoft",
  "license": "MIT",
  "main": "index.js",
  "scripts": {
    "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run dev",
    "dev": "nodemon --inspect=9239 --signal SIGINT ./index.js",
    "start": "node ./index.js",
    "watch": "nodemon ./index.js",
    "test": "echo \"Error: no test specified\" && exit 1"
  },
  "dependencies": {
    "adaptivecards-templating": "^2.3.1",
    "adaptive-expressions": "^4.22.3",
    "botbuilder": "^4.23.1",
    "express": "^5.0.1"
  },
  "devDependencies": {
    "env-cmd": "^10.1.0",
    "nodemon": "^3.1.7"
  }
}
