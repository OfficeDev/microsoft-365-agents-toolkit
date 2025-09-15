{
    "name": "{{SafeProjectNameLowerCase}}",
    "version": "0.1.0",
    "private": true,
    "main": "dist/index",
    "types": "dist/index",
    "files": [
        "dist",
        "README.md"
    ],
    "dependencies": {
        "@microsoft/teams.api": "preview",
        "@microsoft/teams.apps": "preview",
        "@microsoft/teams.cards": "preview",
        "@microsoft/teams.client": "preview",
        "@microsoft/teams.common": "preview",
        "@microsoft/teams.dev": "preview",
        "@microsoft/teams.graph": "preview",
        "@microsoft/teams.graph-endpoints": "preview",
        "react": "^19.0.0",
        "react-dom": "^19.0.0"
    },
    "devDependencies": {
        "@types/node": "^22.5.4",
        "@types/react": "^19.1.12",
        "@types/react-dom": "^19.1.9",
        "@vitejs/plugin-react": "^4.3.4",
        "dotenv": "^16.4.5",
        "env-cmd": "^11.0.0",
        "nodemon": "^3.1.4",
        "rimraf": "^6.0.1",
        "tsup": "^8.4.0",
        "typescript": "^5.4.5",
        "vite": "^6.2.0"
    },
    "scripts": {
        "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run start",
        "clean": "rimraf ./dist",
        "start": "nodemon",
        "prestart": "npm run build:frontend",
        "build": "tsup && npm run build:frontend",
        "build:frontend": "vite build --outDir dist/client"
    }
}
