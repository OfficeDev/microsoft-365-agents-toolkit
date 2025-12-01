{
    "name": "{{SafeProjectNameLowerCase}}",
    "version": "0.0.1",
    "engines": {
        "node": "18 || 20 || 22"
    },
    "author": "Microsoft",
    "license": "MIT",
    "scripts": {
        "test": "echo \"Error: no test specified\" && exit 1",
        "install:bot": "cd bot && npm install",
        "install:tab": "cd tab && npm install",
        "install": "concurrently \"npm run install:bot\" \"npm run install:tab\"",
        "dev:bot": "cd bot && npm run dev",
        "start:tab": "cd tab && npm run start",
        "build:tab": "cd tab && npm run build"
    },
    "dependencies": {
        "concurrently": "^7.6.0"
    },
    "license": "MIT"
}