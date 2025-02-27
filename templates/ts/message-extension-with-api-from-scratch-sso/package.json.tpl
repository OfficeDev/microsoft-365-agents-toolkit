{
    "name": "{{SafeProjectNameLowerCase}}",
    "version": "1.0.0",
    "scripts": {
        "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run dev",
        "dev": "func start --typescript --language-worker=\"--inspect=9229\" --port \"7071\" --cors \"*\"",
        "build": "tsc",
        "watch:teamsfx": "tsc --watch",
        "watch": "tsc -w",
        "prestart": "npm run build",
        "start": "npx func start",
        "test": "echo \"Error: no test specified\" && exit 1"
    },
    "dependencies": {
        "@azure/functions": "^4.3.0",
        "jsonwebtoken": "^9.0.2",
        "jwks-rsa": "^3.1.0",
        "lru-memoizer": "^2.3.0"
    },
    "devDependencies": {
        "@types/node": "^20.11.26",
        "@types/debug": "^4.1.12",
        "env-cmd": "^10.1.0",
        "typescript": "^5.4.2"
    },
    "main": "dist/src/functions/*.js"
}
