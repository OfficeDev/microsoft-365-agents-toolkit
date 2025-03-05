{
    "name": "{{SafeProjectNameLowerCase}}",
    "version": "1.0.0",
    "msteams": {
        "teamsAppId": null
    },
    "description": "Microsoft Teams Toolkit AI Chat Bot Sample with Teams AI Library",
    "engines": {
        "node": "18 || 20 || 22"
    },
    "author": "Microsoft",
    "license": "MIT",
    "main": "./lib/src/index.js",
    "scripts": {
        "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run dev",
        "dev:teamsfx:testtool": "env-cmd --silent -f .localConfigs.testTool npm run dev",
        "dev:teamsfx:launch-testtool": "env-cmd --silent -f env/.env.testtool teamsapptester start",
        "dev": "nodemon --exec node --inspect=9239 --signal SIGINT -r ts-node/register ./src/index.ts",
        "build": "tsc --build && shx cp -r ./src/prompts ./lib/src && shx cp -r ./appPackage ./lib/appPackage && shx cp -r src/adaptiveCards ./lib/src",
        "start": "node ./lib/src/index.js",
        "test": "echo \"Error: no test specified\" && exit 1",
        "watch": "nodemon --exec \"npm run start\""
    },
    "repository": {
        "type": "git",
        "url": "https://github.com"
    },
    "dependencies": {
        "@microsoft/teams-ai": "^1.1.0",
        "botbuilder": "^4.23.1",
        "express": "^5.0.1",
        "fs-extra": "^11.2.0",
        "js-yaml": "^4.1.0",
        "adaptivecards-templating": "^2.3.1",
        "openapi-client-axios": "^7.4.0"
    },
    "devDependencies": {
        "@types/express": "^5.0.0",
        "@types/json-schema": "^7.0.15",
        "@types/node": "^18.0.0",
        "env-cmd": "^10.1.0",
        "ts-node": "^10.4.0",
        "typescript": "^5.5.4",
        "nodemon": "^3.1.7",
        "shx": "^0.3.3"
    }
}