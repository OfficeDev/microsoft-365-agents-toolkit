import React from "react";
import "./Deploy.css";
import { Image } from "@fluentui/react-components";

export function Deploy(props: { docsUrl?: string }) {
  const { docsUrl } = {
    docsUrl: "https://aka.ms/teamsfx-docs",
    ...props,
  };
  return (
    <div className="deploy page">
      <h2>Deploy to the Cloud</h2>
      <p>
        Before publishing your app to App Catalog, you may want to provision and deploy your app's
        resources to the cloud to make sure your app will be running smoothly!
      </p>
      <p>
        To provision your resources, you can either use our CLI command "teamsapp provision" or
        apply "Microsoft 365 Agents: Provision" in Command palette.
      </p>
      <p>
        To deploy your app, you can either use our CLI command "teamsapp deploy" or apply "Teams:
        Deploy" in Command palette.
      </p>
      <Image src="deploy.png" />
      <p>
        For more information, see the{" "}
        <a href={docsUrl} target="_blank" rel="noreferrer">
          docs
        </a>
        .
      </p>
    </div>
  );
}
