import { createContext } from "react";

export const TeamsFxContext = createContext({
  theme: undefined,
  themeString: "",
  teamsUserCredential: undefined,
});
