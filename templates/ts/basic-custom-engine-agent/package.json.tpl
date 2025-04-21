{
    "name": "{{SafeProjectNameLowerCase}}",
    "version": "1.0.0",
    "msteams": {
        "teamsAppId": null
    },
    "description": "Basic Custome Engine Agent with Microsoft 365 Agents SDK",
    "engines": {
        "node": "18 || 20 || 22"
    },
    "author": "Microsoft",
    "license": "MIT",
    "main": "./lib/src/index.js",
    "scripts": {
        "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run dev",
        "dev:teamsfx:testtool": "env-cmd --silent -f .localConfigs.playground npm run dev",
        "dev:teamsfx:launch-testtool": "env-cmd --silent -f env/.env.playground teamsapptester start",
        "dev": "nodemon --exec node --inspect=9239 --signal SIGINT -r ts-node/register ./src/index.ts",
        "build": "tsc --build && shx cp -r ./src/prompts ./lib/src",
        "start": "node ./lib/src/index.js",
        "test": "echo \"Error: no test specified\" && exit 1",
        "watch": "nodemon --exec \"npm run start\""
    },
    "repository": {
        "type": "git",
        "url": "https://github.com"
    },
    "dependencies": {
        "@azure/identity": "^4.8.0",
        "@azure/openai": "^2.0.0",
        "@microsoft/agents-hosting": "^0.1.49",
        "express": "^5.0.1",
        "openai": "^4.94.0"
    },
    "devDependencies": {
        "@types/express": "^5.0.0",
        "@types/node": "^18.0.0",
        "env-cmd": "^10.1.0",
        "nodemon": "^3.1.7",
        "shx": "^0.3.3",
        "ts-node": "^10.4.0",
        "typescript": "^5.5.4"
    }
}