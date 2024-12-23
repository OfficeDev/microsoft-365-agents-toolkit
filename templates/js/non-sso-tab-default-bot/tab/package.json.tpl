{
    "name": "{{SafeProjectNameLowerCase}}",
    "version": "0.1.0",
    "engines": {
        "node": "18 || 20"
    },
    "private": true,
    "dependencies": {
        "@fluentui/react-components": "^9.55.1",
        "@microsoft/teams-js": "^2.31.1",
        "@microsoft/teamsfx": "^3.0.0",
        "@microsoft/teamsfx-react": "^4.0.0",
        "axios": "^0.21.1",
        "react": "^18.2.0",
        "react-dom": "^18.2.0",
        "react-router-dom": "^6.8.0",
        "react-scripts": "^5.0.1"
    },
    "devDependencies": {
        "env-cmd": "^10.1.0"
    },
    "scripts": {
        "dev:teamsfx": "env-cmd --silent -f .localConfigs npm run start",
        "start": "react-scripts start",
        "build": "react-scripts build",
        "eject": "react-scripts eject",
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