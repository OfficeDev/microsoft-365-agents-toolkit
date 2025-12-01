import { createContext } from "react";

import { Theme } from "@fluentui/react-components";

export const TeamsFxContext = createContext<{
  theme?: Theme;
  themeString: string;
}>({
  theme: undefined,
  themeString: "",
});
