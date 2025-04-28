import React from "react";
import ReactDOM from "react-dom/client";
import "./index.css";
import App from "./routes/App";
import reportWebVitals from "./reportWebVitals";
import { createBrowserRouter, RouterProvider } from "react-router-dom";
import { LoginType, Providers, ProviderState } from "@microsoft/mgt-element";
import { Msal2Provider, PromptType } from "@microsoft/mgt-msal2-provider";
import * as Constants from "./common/Constants";
import * as Scopes from "./common/Scopes";
import { initializeFileTypeIcons } from "@fluentui/react-file-type-icons";
import { initializeIcons } from "@fluentui/react/lib/Icons";
import ErrorPage from "./ErrorPage";
import { CustomAppApiAuthProvider } from "./providers/CustomAppApiAuthProvider";
import { Home } from "./routes/Home";
import {
  Containers,
  loader as containersLoader,
  action as createContainerAction,
} from "./routes/Containers";
import { ContainerBrowser, loader as containerLoader } from "./components/ContainerBrowser";
import { GraphAuthProvider } from "./providers/GraphAuthProvider";

// Register icons and pull the fonts from the default Microsoft Fluent CDN:
initializeFileTypeIcons();
initializeIcons();

const provider = new Msal2Provider({
  clientId: Constants.REACT_APP_AZURE_SERVER_APP_ID,
  authority: Constants.AUTH_AUTHORITY,
  scopes: Scopes.GRAPH_SCOPES,
  redirectUri: window.location.origin,
  loginType: LoginType.Redirect,
  prompt: PromptType.SELECT_ACCOUNT,
});
Providers.globalProvider = provider;

provider.onStateChanged(() => {
  if (provider.state === ProviderState.SignedOut) {
    CustomAppApiAuthProvider.instance.client.clearCache();
    GraphAuthProvider.instance._client.clearCache();
  }
});

const router = createBrowserRouter([
  {
    path: "/",
    element: <App />,
    errorElement: <ErrorPage />,
    children: [
      {
        path: "/",
        element: <Home />,
      },
      {
        path: "/containers",
        element: <Containers />,
        loader: containersLoader,
        action: createContainerAction,
      },
      {
        path: "/containers/:containerId/:itemId?",
        element: <ContainerBrowser />,
        loader: containerLoader,
      },
    ],
  },
]);

const root = ReactDOM.createRoot(document.getElementById("root") as HTMLElement);
root.render(
  <React.StrictMode>
    <RouterProvider router={router} />
  </React.StrictMode>
);

// If you want to start measuring performance in your app, pass a function
// to log results (for example: reportWebVitals(console.log))
// or send to an analytics endpoint. Learn more: https://bit.ly/CRA-vitals
reportWebVitals();
