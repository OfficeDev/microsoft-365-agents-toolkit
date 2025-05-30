const config = {
  initiateLoginEndpoint: import.meta.env.VITE_START_LOGIN_PAGE_URL,
  clientId: import.meta.env.VITE_CLIENT_ID,
  apiEndpoint: import.meta.env.VITE_FUNC_ENDPOINT,
  apiName: import.meta.env.VITE_FUNC_NAME,
};

export default config;
