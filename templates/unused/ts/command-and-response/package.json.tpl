{
    "name": "{{SafeProjectNameLowerCase}}",
    "version": "1.0.0",
    "description": "Microsoft 365 Agents Toolkit Command and Response Bot Sample",
    "engines": {
        "node": "20 || 22"
    },
    "author": "Microsoft",
    "license": "MIT",
    "main": "./lib/index.js",
    "scripts": {
        "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run dev",
        "dev:teamsfx:testtool": "env-cmd --silent -f .localConfigs.playground npm run dev",
        "dev:teamsfx:launch-testtool": "env-cmd --silent -f env/.env.playground teamsapptester start",
        "dev": "nodemon --watch ./src --exec node --inspect=9239 --signal SIGINT -r ts-node/register ./src/index.ts",
        "build": "tsc --build && shx cp -r ./src/adaptiveCards ./lib/src",
        "start": "node ./lib/src/index.js",
        "watch": "nodemon --watch ./src --exec \"npm run start\"",
        "test": "echo \"Error: no test specified\" && exit 1"
    },
    "repository": {
        "type": "git",
        "url": "https://github.com"
    },
    "dependencies": {
        "@azure/identity": "^4.11.1",
        "adaptive-expressions": "^4.23.1",
        "adaptivecards-templating": "^2.3.1",
        "adaptivecards": "^3.0.5",
        "@microsoft/teams.apps": "^2.0.0",
        "@microsoft/teams.common": "^2.0.0"
    },
    "devDependencies": {
        "@types/json-schema": "^7.0.15",
        "@types/node": "^20.0.0",
        "env-cmd": "^10.1.0",
        "nodemon": "^3.1.7",
        "shx": "^0.3.4",
        "ts-node": "^10.4.0",
        "typescript": "~5.8.3"
    }
}
