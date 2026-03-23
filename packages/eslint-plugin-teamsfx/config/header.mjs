import headers from "eslint-plugin-headers";

export default {
  plugins: { headers },
  rules: {
    "headers/header-format": [
      "error",
      {
        source: "string",
        style: "line",
        content: "Copyright (c) Microsoft Corporation.\nLicensed under the MIT license.",
      },
    ],
  },
};
