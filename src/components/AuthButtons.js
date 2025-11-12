import React from "react";
import { useMsal } from "@azure/msal-react";
import { loginRequest } from "../auth/msal";
import { LogIn, LogOut } from "lucide-react";

const AuthButtons = ({ compact = false }) => {
  const { instance, accounts } = useMsal();

  const handleLogin = async (event) => {
    event?.preventDefault();
    try {
      await instance.loginPopup(loginRequest);
    } catch (error) {
      console.error("Login failed:", error);
    }
  };

  const handleLogout = () => {
    instance.logoutPopup().catch((error) => {
      console.error("Logout failed:", error);
    });
  };

  const isAuthenticated = accounts && accounts.length > 0;

  if (isAuthenticated) {
    return (
      <button
        onClick={handleLogout}
        className={`flex items-center gap-2 ${
          compact ? "px-2 py-1 text-xs" : "px-3 py-2 text-sm"
        } text-slate-300 hover:text-white hover:bg-white/10 rounded-md transition-colors`}
      >
        <LogOut size={16} />
        Sign Out
      </button>
    );
  }

  return (
    <button
      onClick={handleLogin}
      className={`flex items-center gap-2 ${
        compact ? "px-2 py-1 text-xs" : "px-3 py-2 text-sm"
      } text-slate-300 hover:text-white hover:bg-white/10 rounded-md transition-colors`}
    >
      <LogIn size={16} />
      Sign In
    </button>
  );
};

export default AuthButtons;

