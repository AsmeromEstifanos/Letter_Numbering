import React from "react";
import { useIsAuthenticated } from "@azure/msal-react";

const AuthenticatedView = ({ children, fallback = null }) => {
  const isAuthenticated = useIsAuthenticated();
  return isAuthenticated ? children : fallback;
};

export default AuthenticatedView;
