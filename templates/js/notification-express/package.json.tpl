{
    "name": "{{SafeProjectNameLowerCase}}",
    "version": "1.0.0",
    "description": "Microsoft 365 Agents Toolkit Notification Bot Sample (Express)",
    "engines": {
        "node": "18 || 20 || 22"
    },
    "author": "Microsoft",
    "license": "MIT",
    "main": "./src/index.js",
    "scripts": {
        "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run dev",
        "dev:teamsfx:testtool": "env-cmd --silent -f .localConfigs.playground npm run dev",
        "dev:teamsfx:launch-testtool": "env-cmd --silent -f env/.env.playground teamsapptester start",
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
        "@microsoft/agents-hosting": "^0.2.14",
        "@microsoft/teamsfx": "^4.0.0",
        "adaptive-expressions": "^4.23.1",
        "adaptivecards": "^3.0.5",
        "adaptivecards-templating": "^2.3.1",
        "express": "^5.0.1"
    },
    "devDependencies": {
        "nodemon": "^3.1.7",
        "env-cmd": "^10.1.0"
    }
}