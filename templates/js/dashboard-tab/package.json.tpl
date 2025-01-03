{
    "name": "{{SafeProjectNameLowerCase}}",
    "version": "0.1.0",
    "engines": {
        "node": "18 || 20"
    },
    "type": "module",
    "private": true,
    "dependencies": {
        "@fluentui/react-charting": "^5.14.10",
        "@fluentui/react-components": "^9.55.1",
        "@fluentui/react-icons": "^2.0.186",
        "@microsoft/teams-js": "^2.31.1",
        "@microsoft/teamsfx": "^3.0.0",
        "@microsoft/teamsfx-react": "^4.0.0",
        "react": "^18.2.0",
        "react-dom": "^18.2.0",
        "react-router-dom": "^6.8.0"
    },
    "devDependencies": {
        "@vitejs/plugin-react": "^4.3.1",
        "env-cmd": "^10.1.0",
        "vite": "^5.4.0"
    },
    "scripts": {
        "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run start",
        "start": "vite",
        "build": "vite build",
        "test": "echo \"Error: no test specified\" && exit 1"
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