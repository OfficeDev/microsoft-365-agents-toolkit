{
  "version": "2.0",
  "logging": {
    "applicationInsights": {
      "samplingSettings": {
        "isEnabled": true,
        "excludedTypes": "Request"
      }
    },
    "logLevel": {
      "{{SafeProjectName}}": "Information"
    }
  },
  "extensions": {
    "http": {
      "routePrefix": ""
    }
  }
}
