{
  "name": "{{SafeProjectNameLowerCase}}",
  "version": "1.0.0",
  "description": "Microsoft 365 Agents Toolkit hello world Bot sample",
  "engines": {
    "node": "18 || 20 || 22"
  },
  "author": "Microsoft",
  "license": "MIT",
  "main": "./lib/index.js",
  "scripts": {
    "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run dev",
    "dev": "nodemon --exec node --inspect=9239 --signal SIGINT -r ts-node/register ./index.ts",
    "build": "tsc --build && shx cp -r ./adaptiveCards ./lib/",
    "start": "node ./lib/index.js",
    "watch": "nodemon --exec \"npm run start\"",
    "test": "echo \"Error: no test specified\" && exit 1"
  },
  "repository": {
    "type": "git",
    "url": "https://github.com"
  },
  "dependencies": {
    "adaptivecards-templating": "^2.3.1",
    "adaptive-expressions": "^4.22.3",
    "botbuilder": "^4.23.1",
    "express": "^5.0.1"
  },
  "devDependencies": {
    "@types/express": "^5.0.0",
    "@types/json-schema": "^7.0.15",
    "@types/node": "^18.0.0",
    "env-cmd": "^10.1.0",
    "ts-node": "^10.4.0",
    "typescript": "^4.4.4",
    "nodemon": "^3.1.7",
    "shx": "^0.3.3"
  }
}