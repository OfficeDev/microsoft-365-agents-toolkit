import { useEffect, useState } from "react";
import { unstable_batchedUpdates as batchedUpdates } from "react-dom";
import { app, pages } from "@microsoft/teams-js";
import {
  teamsLightTheme,
  teamsDarkTheme,
  teamsHighContrastTheme,
} from "@fluentui/react-components";

const getTheme = () => {
  const urlParams = new URLSearchParams(window.location.search);
  const theme = urlParams.get("theme");
  return theme == null ? undefined : theme;
};

/**
 * Microsoft Teams React hook
 * @param {Object} options optional options
 * @returns A tuple with properties and methods
 * properties:
 *  - inTeams: boolean = true if inside Microsoft Teams
 *  - fullscreen: boolean = true if in full screen mode
 *  - theme: Fluent UI Theme
 *  - themeString: string - representation of the theme (default, dark or contrast)
 *  - context - the Microsoft Teams JS SDK context
 * methods:
 *  - setTheme - manually set the theme
 */
export function useTeams(options) {
  const [loading, setLoading] = useState(undefined);
  const [inTeams, setInTeams] = useState(undefined);
  const [fullScreen, setFullScreen] = useState(undefined);
  const [theme, setTheme] = useState(teamsLightTheme);
  const [themeString, setThemeString] = useState("default");
  const [initialTheme] = useState(
    options && options.initialTheme ? options.initialTheme : getTheme()
  );
  const [context, setContext] = useState(undefined);

  const themeChangeHandler = (theme) => {
    setThemeString(theme || "default");
    switch (theme) {
      case "dark":
        setTheme(teamsDarkTheme);
        break;
      case "contrast":
        setTheme(teamsHighContrastTheme);
        break;
      case "default":
      default:
        setTheme(teamsLightTheme);
    }
  };

  const overrideThemeHandler = options?.setThemeHandler
    ? options.setThemeHandler
    : themeChangeHandler;

  useEffect(() => {
    // set initial theme based on options or query string
    if (initialTheme) {
      overrideThemeHandler(initialTheme);
    }

    app
      .initialize()
      .then(() => {
        app
          .getContext()
          .then((context) => {
            batchedUpdates(() => {
              setInTeams(true);
              setContext(context);
              setFullScreen(context.page.isFullScreen);
            });
            overrideThemeHandler(context.app.theme);
            app.registerOnThemeChangeHandler(overrideThemeHandler);
            pages.registerFullScreenHandler((isFullScreen) => {
              setFullScreen(isFullScreen);
            });
            setLoading(false);
          })
          .catch(() => {
            setLoading(false);
            setInTeams(false);
          });
      })
      .catch(() => {
        setLoading(false);
        setInTeams(false);
      });
  }, [initialTheme, overrideThemeHandler]);

  return [
    {
      inTeams,
      fullScreen,
      theme,
      themeString,
      context,
      loading,
    },
    {
      setTheme: (theme) => {
        overrideThemeHandler(theme);
      },
    },
  ];
}
