{
    "name": "{{SafeProjectNameLowerCase}}",
    "version": "0.1.0",
    "engines": {
        "node": "18 || 20"
    },
    "private": true,
    "main": "./lib/app.js",
    "dependencies": {
        "express": "^4.21.1",
        "send": "^0.18.0"
    },
    "devDependencies": {
        "@microsoft/teams-js": "^2.31.1",
        "@types/node": "^18.0.0",
        "@types/express": "^5.0.0",
        "@types/send": "^0.17.1",
        "env-cmd": "^10.1.0",
        "nodemon": "^3.1.7",
        "ts-node": "^10.9.1",
        "typescript": "^4.1.2",
        "shx": "^0.3.4",
        "vite": "^6.0.2"
    },
    "scripts": {
        "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run start",
        "start": "nodemon",
        "prestart": "npm run build:frontend",
        "build": "npm run build:frontend && npm run build:backend",
        "build:frontend": "vite build",
        "build:backend": "tsc --build && shx cp -r ./src/views ./src/static ./lib/",
        "test": "echo \"Error: no test specified\" && exit 1"
    },
    "homepage": "."
}
