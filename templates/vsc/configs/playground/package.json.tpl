{
    "devDependencies": {
        "env-cmd": "^11.0.0"
    },
    "scripts": {
        "dev:teamsfx:testtool": "env-cmd --silent -f .localConfigs.playground npm run start",
        "dev:teamsfx:launch-testtool": "env-cmd --silent -f env/.env.playground teamsapptester start"
    }
}
