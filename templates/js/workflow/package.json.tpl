{
    "name": "{{SafeProjectNameLowerCase}}",
    "version": "1.0.0",
    "description": "Microsoft Teams Toolkit Workflow Bot Sample",
    "engines": {
        "node": "18 || 20"
    },
    "author": "Microsoft",
    "license": "MIT",
    "main": "./src/index.js",
    "scripts": {
        "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run dev",
        "dev:teamsfx:testtool": "env-cmd --silent -f .localConfigs.testTool npm run dev",
        "dev:teamsfx:launch-testtool": "env-cmd --silent -f env/.env.testtool teamsapptester start",
        "dev": "nodemon --inspect=9239 --signal SIGINT ./src/index.js",
        "start": "node ./src/index.js",
        "watch": "nodemon ./src/index.js",
        "test": "echo \"Error: no test specified\" && exit 1"
    },
    "repository": {
        "type": "git",
        "url": "https://github.com"
    },
    "dependencies": {
        "@microsoft/teams-ai": "^1.7.0",
        "adaptive-expressions": "^4.23.1",
        "adaptivecards-templating": "^2.3.1",
        "botbuilder": "^4.23.1",
        "express": "^5.0.1"
    },
    "devDependencies": {
        "env-cmd": "^10.1.0",
        "nodemon": "^3.1.7"
    }
}