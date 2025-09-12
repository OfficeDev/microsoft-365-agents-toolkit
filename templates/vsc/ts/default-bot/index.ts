import { startServer } from '@microsoft/teams.apps';
import { createTeamsApp } from './app';
import { createConfig } from './config';

async function main() {
  const config = createConfig();
  const app = await createTeamsApp();
  
  startServer(app, {
    port: process.env.PORT ? parseInt(process.env.PORT) : 3978,
    credentials: config.credentials
  });
}

main().catch(console.error);
