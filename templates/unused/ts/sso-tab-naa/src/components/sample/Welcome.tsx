import { useState } from "react";
import {
  Image,
  TabList,
  Tab,
  SelectTabEvent,
  SelectTabData,
  TabValue,
} from "@fluentui/react-components";
import "./Welcome.css";
import { EditCode } from "./EditCode";
import { CurrentUser } from "./CurrentUser";
import { Deploy } from "./Deploy";
import { Publish } from "./Publish";
import { app } from "@microsoft/teams-js";
import { useData } from "./lib/useData";
import { NestedAppAuth } from "./NestedAppAuth";

export function Welcome(props: { environment?: string }) {
  const { environment } = {
    environment: window.location.hostname === "localhost" ? "local" : "azure",
    ...props,
  };
  const friendlyEnvironmentName =
    {
      local: "local environment",
      azure: "Azure environment",
    }[environment] || "local environment";

  const [selectedValue, setSelectedValue] = useState<TabValue>("local");

  const onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
    setSelectedValue(data.value);
  };
  const { loading, data, error } = useData(async () => {
    await app.initialize();
    const context = await app.getContext();
    if (context.user) {
      return {
        displayName: context.user.displayName || "",
      };
    }
  });
  const userName = loading || error ? "" : data!.displayName;
  const hubName = useData(async () => {
    await app.initialize();
    const context = await app.getContext();
    return context.app.host.name;
  })?.data;
  return (
    <div className="welcome page">
      <div className="narrow page-padding">
        <Image src="hello.png" />
        <h1 className="center">Congratulations{userName ? ", " + userName : ""}!</h1>
        {hubName && <p className="center">Your app is running in {hubName}</p>}
        <p className="center">Your app is running in your {friendlyEnvironmentName}</p>

        <div className="tabList">
          <TabList selectedValue={selectedValue} onTabSelect={onTabSelect}>
            <Tab id="Local" value="local">
              1. Build your app locally
            </Tab>
            <Tab id="Azure" value="azure">
              2. Provision and Deploy to the Cloud
            </Tab>
            <Tab id="Publish" value="publish">
              3. Publish to Teams
            </Tab>
          </TabList>
          <div>
            {selectedValue === "local" && (
              <div>
                <EditCode />
                <CurrentUser userName={userName} />
                <NestedAppAuth />
              </div>
            )}
            {selectedValue === "azure" && (
              <div>
                <Deploy />
              </div>
            )}
            {selectedValue === "publish" && (
              <div>
                <Publish />
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
