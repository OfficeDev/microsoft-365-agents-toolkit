{
    "name": "{{SafeProjectNameLowerCase}}",
    "version": "1.0.0",
    "description": "Microsoft 365 Agents Toolkit echo bot sample (Python)",
    "author": "Microsoft",
    "license": "MIT",
    "scripts": {
        "dev:teamsfx:testtool": "env-cmd --silent -f .localConfigs.playground python src/app.py",
        "dev:teamsfx:launch-testtool": "env-cmd --silent -f env/.env.playground teamsapptester start"
    },
    "repository": {
        "type": "git",
        "url": "https://github.com"
    },
    "devDependencies": {
        "env-cmd": "^10.1.0"
    }
}
