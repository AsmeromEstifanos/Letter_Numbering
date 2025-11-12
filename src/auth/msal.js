import { PublicClientApplication } from "@azure/msal-browser";

// Resolve values from .env (CRA exposes only REACT_APP_* at build time)
const envClientId =
  process.env.REACT_APP_AZURE_CLIENT_ID ||
  process.env.REACT_APP_AAD_CLIENT_ID ||
  process.env.REACT_APP_MSAL_CLIENT_ID ||
  process.env.REACT_APP_CLIENT_ID;

const envAuthority =
  process.env.REACT_APP_AZURE_AUTHORITY ||
  process.env.REACT_APP_AAD_AUTHORITY ||
  process.env.REACT_APP_MSAL_AUTHORITY ||
  undefined;

const envTenant =
  process.env.REACT_APP_AAD_TENANT_ID ||
  process.env.REACT_APP_TENANT_ID ||
  process.env.REACT_APP_AAD_TENANT_DOMAIN ||
  undefined; // e.g. contoso.onmicrosoft.com or tenant guid

// Build authority if a full one wasn't provided
let authority = envAuthority;
if (!authority) {
  if (
    envTenant &&
    typeof envTenant === "string" &&
    envTenant.trim().length > 0
  ) {
    authority = `https://login.microsoftonline.com/${envTenant.trim()}`;
  } else {
    console.warn(
      "[MSAL] No tenant id/domain provided via .env (REACT_APP_AAD_TENANT_ID or REACT_APP_AZURE_AUTHORITY). Falling back to 'common'."
    );
    authority = "https://login.microsoftonline.com/common";
  }
}

if (!envClientId) {
  console.error(
    "[MSAL] Missing client id. Set REACT_APP_AZURE_CLIENT_ID (preferred) or REACT_APP_AAD_CLIENT_ID in .env."
  );
}

const redirectUri =
  process.env.REACT_APP_REDIRECT_URI ||
  process.env.REACT_APP_AAD_REDIRECT_URI ||
  window.location.origin;
const postLogoutRedirectUri =
  process.env.REACT_APP_POST_LOGOUT_REDIRECT_URI ||
  process.env.REACT_APP_AAD_POST_LOGOUT_REDIRECT_URI ||
  window.location.origin;

const msalConfig = {
  auth: {
    clientId: envClientId || "",
    authority,
    redirectUri,
    postLogoutRedirectUri,
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: true,
  },
};

export const msalInstance = new PublicClientApplication(msalConfig);

export const loginRequest = {
  scopes: (
    process.env.REACT_APP_LOGIN_SCOPES ||
    "openid profile email User.Read"
  )
    .split(/[\s,]+/)
    .filter(Boolean),
};
