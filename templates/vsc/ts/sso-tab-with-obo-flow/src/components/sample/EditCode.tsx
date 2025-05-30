import React from "react";
import config from "./lib/config";

const functionName = config.apiName || "myFunc";

export function EditCode(props: {
  showFunction?: boolean;
  tabCodeEntry?: string;
  functionCodePath?: string;
}) {
  const { showFunction, tabCodeEntry, functionCodePath } = {
    showFunction: true,
    tabCodeEntry: "src/index.tsx",
    functionCodePath: `api/src/functions/${functionName}.ts`,
    ...props,
  };
  return (
    <div>
      <h2>Change this code</h2>
      <p>
        The front end is a <code>create-react-app</code>. The entry point is{" "}
        <code>{tabCodeEntry}</code>. Just save any file and this page will reload automatically.
      </p>
      {showFunction && (
        <p>
          This app contains an Azure Functions backend. Find the code in{" "}
          <code>{functionCodePath}</code>
        </p>
      )}
    </div>
  );
}
