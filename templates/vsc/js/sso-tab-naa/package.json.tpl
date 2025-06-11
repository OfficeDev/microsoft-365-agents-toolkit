{
    "name": "{{SafeProjectNameLowerCase}}",
    "version": "0.1.0",
    "engines": {
        "node": "18 || 20 || 22"
    },
    "type": "module",
    "private": true,
    "dependencies": {
        "@azure/msal-browser": "^4.12.0",
        "@fluentui/react-components": "^9.55.1",
        "@microsoft/teams-js": "^2.31.1",
        "axios": "^0.21.1",
        "react": "^18.2.0",
        "react-dom": "^18.2.0",
        "react-router-dom": "^6.8.0"
    },
    "devDependencies": {
        "@types/node": "^18.0.0",
        "@types/react": "^18.0.0",
        "@types/react-dom": "^18.0.0",
        "@types/react-router-dom": "^5.3.3",
        "concurrently": "^8.2.2",
        "env-cmd": "^10.1.0",
        "vite": "^5.4.2",
        "@vitejs/plugin-basic-ssl": "^1.1.0",
        "@vitejs/plugin-react": "^4.3.1"
    },
    "scripts": {
        "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run start",
        "start": "vite",
        "build": "vite build",
        "test": "echo \"Error: no test specified\" && exit 1",
        "serve": "vite preview"
    },
    "eslintConfig": {
        "extends": [
            "react-app",
            "react-app/jest"
        ]
    },
    "browserslist": {
        "production": [
            ">0.2%",
            "not dead",
            "not op_mini all"
        ],
        "development": [
            "last 1 chrome version",
            "last 1 firefox version",
            "last 1 safari version"
        ]
    },
    "homepage": "."
}