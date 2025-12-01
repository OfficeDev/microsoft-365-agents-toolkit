import { app } from "@microsoft/teams-js";

(function () {
  "use strict";

  // Call the initialize API first
  app.initialize().then(() => {
    app.getContext().then((context: app.Context) => {
      if (context?.app?.host?.name) {
        updateHubState(context.app.host.name);
      }
    });
  });

  function updateHubState(hubName: string) {
    if (hubName) {
      const hubStateElement = document.getElementById("hubState");
      if (hubStateElement) {
        hubStateElement.innerHTML = "in " + hubName;
      }
    }
  }

  // Notify success when the DOM content is fully loaded
  document.addEventListener("DOMContentLoaded", () => {
    app.notifySuccess();
  });
})();
