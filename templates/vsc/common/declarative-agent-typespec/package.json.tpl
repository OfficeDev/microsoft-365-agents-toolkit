{
  "name": "{{SafeProjectNameLowerCase}}",
  "version": "1.0.0",
  "scripts": {
    "compile": "tsp compile ./src/agent/main.tsp --config ./tspconfig.yaml",
    "generate:env": "node scripts/generate-env.js"
  },
  "devDependencies": {
    "@microsoft/typespec-m365-copilot": "1.0.0-rc.6",
    "@typespec/compiler": "1.5.0",
    "@typespec/http": "1.5.0",
    "@typespec/openapi": "1.5.0",
    "@typespec/openapi3": "1.5.0"
  }
}