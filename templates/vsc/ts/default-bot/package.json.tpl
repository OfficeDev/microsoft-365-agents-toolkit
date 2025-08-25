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
        "@microsoft/agents-hosting-express": "^1.0.0"
    },
    "devvDependencies": {
        "@types/express": "^5.0.0",
        "@types/node": "^22.0.0",
        "env-cmd": "^10.1.0",
        "nodemon": "^3.1.10",
        "shx": "^0.3.3",
        "ts-node": "^10.9.2",
        "typescript": "^5.8.3"
    }
}
