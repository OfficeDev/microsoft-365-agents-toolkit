import { getContext, shareDeepLink, uninitializeCommunication } from "@microsoft/teams-js";
import { initialize as init } from "@microsoft/teams-js";

getContext();

getContext(() => {});

init();

shareDeepLink();

uninitializeCommunication();
