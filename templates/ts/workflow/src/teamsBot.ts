import { TeamsApplication } from "@microsoft/agents-hosting-teams";
import { MemoryStorage } from "@microsoft/agents-hosting";
import { ApplicationTurnState } from "./internal/interface";

// Define storage and application
const storage = new MemoryStorage();
export const app = new TeamsApplication<ApplicationTurnState>({
  storage,
});