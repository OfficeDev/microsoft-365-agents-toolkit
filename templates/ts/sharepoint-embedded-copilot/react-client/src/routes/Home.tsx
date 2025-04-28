import { Link as FluentLink } from "@fluentui/react-components";
import { Login, ProviderState, Providers } from "@microsoft/mgt-react";
import { useEffect, useState } from "react";
import * as Constants from "../common/Constants";
import { Link } from "react-router-dom";

const useIsSignedIn = () => {
  const [isSignedIn, setIsSignedIn] = useState<boolean>(false);

  useEffect(() => {
    const updateIsSignedIn = () => {
      setIsSignedIn(Providers.globalProvider.state === ProviderState.SignedIn);
    };
    updateIsSignedIn();
    Providers.globalProvider.onStateChanged(updateIsSignedIn);
    return () => {
      Providers.globalProvider.removeStateChangedHandler(updateIsSignedIn);
    };
  }, []);
  return isSignedIn;
};

export const Home: React.FunctionComponent = () => {
  const isSignedIn = useIsSignedIn();

  let tenantId = "";
  let adminConsentLink = "";
  if (isSignedIn) {
    tenantId = Providers.globalProvider.getActiveAccount!()?.tenantId || "";
    adminConsentLink = `https://login.microsoftonline.com/${tenantId}/adminconsent?client_id=${Constants.REACT_APP_AZURE_SERVER_APP_ID}&redirect_uri=${window.location.origin}`;
  }

  return (
    <div>
      <h2>Getting started with the SharePoint Embedded sample app</h2>
      <p>Follow the steps below to get your tenant setup for this sample app</p>
      <ol className="setup-steps">
        {!isSignedIn && (
          <li>
            <Login /> using a gloabl admin account
          </li>
        )}
        {isSignedIn && (
          <li>
            <FluentLink href={adminConsentLink}>Grant admin consent to this demo app</FluentLink>
          </li>
        )}
        {isSignedIn && (
          <li>
            Visit the <Link to="/containers">Containers</Link> page to use the demo app
          </li>
        )}
      </ol>
    </div>
  );
};
