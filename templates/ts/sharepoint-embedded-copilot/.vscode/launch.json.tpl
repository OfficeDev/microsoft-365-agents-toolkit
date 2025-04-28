{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Attach to Local Service",
            "type": "node",
            "request": "attach",
            "port": 9239,
            "restart": true,
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Launch chrome against localhost",
            "type": "chrome",
            "request": "launch",
            "url": "http://localhost:3000",
            "webRoot": "${workspaceFolder}/react-client",
            "preLaunchTask": "Start Teams App Locally",
        },
    ],
    "compounds": [
    ]
}
