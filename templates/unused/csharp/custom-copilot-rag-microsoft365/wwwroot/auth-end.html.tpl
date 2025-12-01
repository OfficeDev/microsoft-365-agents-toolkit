﻿<html>
<head>
    <title>Login End Page</title>
    <meta charset="utf-8" />
</head>

<body>
    <script src="https://statics.teams.cdn.office.net/sdk/v1.6.0/js/MicrosoftTeams.min.js"
            integrity="sha384-mhp2E+BLMiZLe7rDIzj19WjgXJeI32NkPvrvvZBrMi5IvWup/1NUfS5xuYN5S3VT"
            crossorigin="anonymous"></script>
    <div id="divError"></div>
    <script type="text/javascript">
      microsoftTeams.initialize();
      let hashParams = getHashParameters();

      if (hashParams["error"]) {
        // Authentication failed
        handleAuthError(hashParams["error"], hashParams);
      } else if (hashParams["code"]) {
        // Get the stored state parameter and compare with incoming state
        let expectedState = localStorage.getItem("state");
        if (expectedState !== hashParams["state"]) {
          // State does not match, report error
          handleAuthError("StateDoesNotMatch", hashParams);
        } else {
          microsoftTeams.authentication.notifySuccess();
        }
      } else {
        // Unexpected condition: hash does not contain error or access_token parameter
        handleAuthError("UnexpectedFailure", hashParams);
      }

      // Parse hash parameters into key-value pairs
      function getHashParameters() {
        let hashParams = {};
        location.hash
          .substr(1)
          .split("&")
          .forEach(function (item) {
            let s = item.split("="),
              k = s[0],
              v = s[1] && decodeURIComponent(s[1]);
            hashParams[k] = v;
          });
        return hashParams;
      }

      // Show error information
      function handleAuthError(errorType, errorMessage) {
        const err = JSON.stringify({
          error: errorType,
          message: JSON.stringify(errorMessage),
        });
        let para = document.createElement("p");
        let node = document.createTextNode(err);
        para.appendChild(node);

        let element = document.getElementById("divError");
        element.appendChild(para);
      }
    </script>
</body>
</html>