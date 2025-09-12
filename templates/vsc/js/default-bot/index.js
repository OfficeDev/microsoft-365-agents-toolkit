const { startServer } = require('@microsoft/teams.apps');
const { createTeamsApp } = require('./app');
const { createConfig } = require('./config');

async function main() {
  const config = createConfig();
  const app = await createTeamsApp();
  
  startServer(app, {
    port: process.env.PORT ? parseInt(process.env.PORT) : 3978,
    credentials: config.credentials
  });
}

main().catch(console.error);
