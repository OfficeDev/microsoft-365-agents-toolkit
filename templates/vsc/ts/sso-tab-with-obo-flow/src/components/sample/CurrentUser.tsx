import React from "react";

export function CurrentUser(props: { userName?: string }) {
  const { userName } = {
    userName: "",
    ...props,
  };
  return (
    <div>
      <h2>Get the current user</h2>
      <p>Access basic information about the user like this:</p>
      <pre>
        {`const authConfig: TeamsUserCredentialAuthConfig = {\n  clientId: process.env.REACT_APP_CLIENT_ID,\n  initiateLoginEndpoint: process.env.REACT_APP_START_LOGIN_PAGE_URL,\n};\n\nconst credential = new TeamsUserCredential(authConfig);\nconst user = await credential.getUserInfo();`}
      </pre>
      {!!userName && (
        <p>
          The currently logged in user's name is <b>{userName}</b>
        </p>
      )}
    </div>
  );
}
