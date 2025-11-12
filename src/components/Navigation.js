import React, { useState, useEffect } from "react";
import { Link, useLocation } from "react-router-dom";
import {
  Home,
  Hash,
  Menu,
  X,
  FileDigit,
  Shield,
  Users,
  RefreshCw,
} from "lucide-react";
import AuthButtons from "./AuthButtons";
import { useLetters } from "../context/LetterContext";

const Navigation = ({ isMobileMenuOpen, setIsMobileMenuOpen }) => {
  const [isCollapsed, setIsCollapsed] = useState(false);
  const [isMobile, setIsMobile] = useState(
    typeof window !== "undefined" ? window.innerWidth < 1024 : false
  );
  const location = useLocation();
  const { refreshAll, loading, isAdmin, userAccessResolved } = useLetters();

  useEffect(() => {
    const onResize = () => {
      const mobile = window.innerWidth < 1024;
      setIsMobile(mobile);
      if (!mobile) {
        setIsMobileMenuOpen(false);
      }
    };
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, [setIsMobileMenuOpen]);

  // On mobile, we want to show the full navigation when menu is open
  // On desktop, we respect the collapsed state
  const isNavCollapsed = isMobile ? !isMobileMenuOpen : isCollapsed;
  const widthClass = isMobile
    ? isMobileMenuOpen
      ? "w-72"
      : "w-0"
    : isNavCollapsed
    ? "w-[72px]"
    : "w-72";

  const baseNavItems = [
    { path: "/", label: "Dashboard", icon: Home },
    { path: "/letters", label: "Letter Numbers", icon: Hash },
  ];
  const navItems =
    userAccessResolved && isAdmin
      ? [
          ...baseNavItems,
          { path: "/admin", label: "Admin Panel", icon: Shield },
        ]
      : baseNavItems;
  const isRefreshing = loading.companies || loading.letters;

  const isActive = (path) => location.pathname === path;

  return (
    <>
      {/* Mobile overlay */}
      {isMobile && isMobileMenuOpen && (
        <div
          className="fixed inset-0 bg-black/50 z-[999] lg:hidden"
          onClick={() => setIsMobileMenuOpen(false)}
        />
      )}

      <nav
        className={`fixed h-screen z-[1000] shadow-lg bg-gradient-to-br from-slate-800 to-slate-700 text-white transition-[width] duration-300 ${
          isMobile ? "left-0" : ""
        } ${widthClass}`}
      >
        <div className="flex items-center justify-between px-4 sm:px-6 py-5 border-b border-white/10">
          <div className="flex items-center gap-3">
            {!isNavCollapsed && (
              <>
                <FileDigit className="w-8 h-8 text-blue-400" />
                <span className="text-xl font-bold tracking-tight">
                  Letter Numbering
                </span>
              </>
            )}
          </div>
          <div className="flex items-center gap-2">
            {!isNavCollapsed && (
              <button
                onClick={refreshAll}
                disabled={isRefreshing}
                className="p-2 rounded-md text-slate-300 hover:text-white hover:bg-white/10 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
                title="Refresh Data"
              >
                <RefreshCw
                  size={18}
                  className={isRefreshing ? "animate-spin" : ""}
                />
              </button>
            )}
            <button
              className="p-2 rounded-md text-slate-300 hover:text-white hover:bg-white/10"
              onClick={() => {
                if (isMobile) {
                  setIsMobileMenuOpen(!isMobileMenuOpen);
                } else {
                  setIsCollapsed(!isCollapsed);
                }
              }}
              title={
                isMobile
                  ? isMobileMenuOpen
                    ? "Close menu"
                    : "Open menu"
                  : isNavCollapsed
                  ? "Expand sidebar"
                  : "Collapse sidebar"
              }
            >
              {isMobile ? (
                isMobileMenuOpen ? (
                  <X size={20} />
                ) : (
                  <></>
                  // <Menu size={20} />
                )
              ) : isNavCollapsed ? (
                <Menu size={20} />
              ) : (
                <X size={20} />
              )}
            </button>
          </div>
        </div>

        <div className="flex flex-col py-4 h-[calc(100vh-76px)]">
          <ul className="space-y-1 flex-1 overflow-y-auto">
            {navItems.map((item) => {
              const Icon = item.icon;
              const active = isActive(item.path);
              return (
                <li key={item.path}>
                  <Link
                    to={item.path}
                    className={`mx-2 flex items-center gap-3 rounded-md px-2 sm:px-4 py-2 text-slate-300 hover:text-white transition-colors ${
                      active ? "bg-blue-500 text-white" : "hover:bg-white/10"
                    } ${isNavCollapsed ? "justify-center" : ""}`}
                    title={isNavCollapsed ? item.label : ""}
                    onClick={() => {
                      if (isMobile) {
                        setIsMobileMenuOpen(false);
                      }
                    }}
                  >
                    <Icon size={20} className="shrink-0" />
                    {!isNavCollapsed && (
                      <span className="font-medium">{item.label}</span>
                    )}
                  </Link>
                </li>
              );
            })}
          </ul>

          {!isNavCollapsed && (
            <div className="mt-auto px-4 pt-4 border-t border-white/10">
              <div className="space-y-2 text-slate-300 mb-3">
                <div className="flex items-center gap-2 text-xs">
                  <Shield size={14} />
                  <span>SharePoint Linked</span>
                </div>
                <div className="flex items-center gap-2 text-xs">
                  <Users size={14} />
                  <span>Multi-Company</span>
                </div>
              </div>
              <AuthButtons />
            </div>
          )}
        </div>
      </nav>
    </>
  );
};

export default Navigation;
