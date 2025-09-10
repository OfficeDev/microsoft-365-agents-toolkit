{
    "name": "{{SafeProjectNameLowerCase}}",
    "version": "1.0.0",
    "description": "Microsoft 365 Agents Toolkit echo bot sample",
    "author": "Microsoft",
    "license": "MIT",
    "main": "./lib/index.js",
    "scripts": {
        "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run dev",
        "dev:teamsfx:testtool": "env-cmd --silent -f .localConfigs.playground npm run dev",
        "dev:teamsfx:launch-testtool": "env-cmd --silent -f env/.env.playground teamsapptester start",
        "dev": "nodemon --exec node --inspect=9239 --signal SIGINT -r ts-node/register ./index.ts",
        "build": "tsc --build",
        "start": "node ./lib/index.js",
        "watch": "nodemon --exec \"npm run start\"",
        "test": "echo \"Error: no test specified\" && exit 1"
    },
    "repository": {
        "type": "git",
        "url": "https://github.com"
    },
    "dependencies": {
        "@azure/identity": "^4.11.1",
        "@microsoft/teams.apps": "preview",
        "@microsoft/teams.common": "preview"
    },
    "devDependencies": {
        "@types/node": "^20.0.0",
        "env-cmd": "^10.1.0",
        "nodemon": "^3.1.7",
        "ts-node": "^10.4.0",
        "typescript": "~5.8.3"
    }
}
