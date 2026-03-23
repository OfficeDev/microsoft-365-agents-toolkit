import { getContext, shareDeepLink, uninitializeCommunication, initialize as init  } from "@microsoft/teams-js";

getContext();

getContext(() => {});

init();

shareDeepLink();

uninitializeCommunication();
