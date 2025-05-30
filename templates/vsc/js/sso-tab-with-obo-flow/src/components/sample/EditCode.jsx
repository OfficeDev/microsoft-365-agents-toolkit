import React from "react";
import config from "./lib/config";

var functionName = config.apiName || "myFunc";

export function EditCode(props) {
  const { showFunction, tabCodeEntry, functionCodePath } = {
    showFunction: true,
    tabCodeEntry: "src/index.jsx",
    functionCodePath: `api/src/functions/${functionName}.js`,
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
