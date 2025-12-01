{
    "name": "{{SafeProjectNameLowerCase}}",
    "version": "0.1.0",
    "engines": {
        "node": "18 || 20 || 22"
    },
    "private": true,
    "dependencies": {
        "express": "^4.21.1",
        "send": "^0.18.0"
    },
    "devDependencies": {
        "env-cmd": "^10.1.0",
        "nodemon": "^2.0.21"
    },
    "scripts": {
        "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run start",
        "start": "nodemon --inspect=9239 --signal SIGINT src/app.js",
        "test": "echo \"Error: no test specified\" && exit 1"
    },
    "homepage": "."
}
