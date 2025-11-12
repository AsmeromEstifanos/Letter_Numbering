import React, { useEffect, useMemo, useState } from "react";
import { Building2, RefreshCw, UserPlus, Trash2, Edit3 } from "lucide-react";
import LoadingSpinner from "./LoadingSpinner";
import { useLetters } from "../context/LetterContext";

const inputClass =
  "w-full rounded-md border border-slate-200 bg-white px-3 py-2 text-sm text-slate-900 shadow-sm focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-500";
const labelClass =
  "block text-xs font-semibold uppercase tracking-wide text-slate-600 mb-1";
const primaryButtonClass =
  "inline-flex items-center justify-center rounded-md bg-blue-600 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2";

const roleOptions = ["Admin", "Editor", "Viewer"];

const AdminPanel = () => {
  const {
    companies,
    letters,
    loading,
    error,
    addCompany,
    updateCompany,
    deleteCompany,
    loadCompanies,
    userAccess,
    loadUserAccess,
    addUserAccessEntry,
    updateUserAccessEntry,
    deleteUserAccessEntry,
    getReferencePreview,
    isAdmin,
    userAccessResolved,
    searchUsers,
  } = useLetters();

  const [companyForm, setCompanyForm] = useState({
    name: "",
    abbreviation: "",
    startingNumber: 1,
  });
  const [companyFormError, setCompanyFormError] = useState(null);
  const [companySuccess, setCompanySuccess] = useState(null);
  const [editingCompany, setEditingCompany] = useState(null);

  const [userForm, setUserForm] = useState({
    userPrincipalName: "",
    role: "Editor",
    companyIds: [],
  });
  const [userFormError, setUserFormError] = useState(null);
  const [userSuccess, setUserSuccess] = useState(null);
  const [isEmailFocused, setIsEmailFocused] = useState(false);
  const [editingUserEntry, setEditingUserEntry] = useState(null);

  useEffect(() => {
    if (companies.length === 0) {
      loadCompanies();
    }
    loadUserAccess();
  }, [companies.length, loadCompanies, loadUserAccess]);

  const stats = useMemo(() => {
    const companyMap = new Map();
    letters.forEach((letter) => {
      const entry = companyMap.get(letter.companyId) || 0;
      companyMap.set(letter.companyId, entry + 1);
    });
    return companyMap;
  }, [letters]);

  const sortedUserAccess = useMemo(() => {
    return [...userAccess].sort((a, b) =>
      (a.userPrincipalName || "").localeCompare(
        b.userPrincipalName || "",
        undefined,
        { sensitivity: "base" }
      )
    );
  }, [userAccess]);

  const emailSuggestions = useMemo(() => {
    const emails = new Set();
    userAccess.forEach((entry) => {
      if (entry.userPrincipalName) {
        emails.add(entry.userPrincipalName.toLowerCase());
      }
    });
    return Array.from(emails).sort();
  }, [userAccess]);

  const [liveEmailSuggestions, setLiveEmailSuggestions] = useState([]);
  const [isSearchingEmails, setIsSearchingEmails] = useState(false);

  useEffect(() => {
    let active = true;
    const term = (userForm.userPrincipalName || "").trim();
    if (term.length < 2) {
      setLiveEmailSuggestions([]);
      setIsSearchingEmails(false);
      return;
    }
    setIsSearchingEmails(true);
    const handle = setTimeout(() => {
      searchUsers(term)
        .then((results) => {
          if (!active) return;
          setLiveEmailSuggestions(results);
        })
        .catch(() => {
          if (active) setLiveEmailSuggestions([]);
        })
        .finally(() => {
          if (active) setIsSearchingEmails(false);
        });
    }, 300);
    return () => {
      active = false;
      clearTimeout(handle);
    };
  }, [userForm.userPrincipalName, searchUsers]);

  const combinedEmailSuggestions = useMemo(() => {
    const map = new Map();
    const term = (userForm.userPrincipalName || "").trim().toLowerCase();
    const matchesTerm = (value, label) => {
      if (!term) return false;
      return (
        value.toLowerCase().includes(term) ||
        (label && label.toLowerCase().includes(term))
      );
    };
    emailSuggestions.forEach((email) => {
      if (!email) return;
      if (term && !email.includes(term)) return;
      map.set(email, { value: email, label: email });
    });
    liveEmailSuggestions.forEach((user) => {
      const rawEmail = (
        user.email ||
        user.userPrincipalName ||
        ""
      ).toLowerCase();
      if (!rawEmail) return;
      const label = user.displayName
        ? `${user.displayName} (${rawEmail})`
        : rawEmail;
      if (!matchesTerm(rawEmail, label)) return;
      map.set(rawEmail, { value: rawEmail, label });
    });
    return Array.from(map.values()).slice(0, 20);
  }, [emailSuggestions, liveEmailSuggestions, userForm.userPrincipalName]);

  const searchTerm = (userForm.userPrincipalName || "").trim().toLowerCase();
  const shouldShowSuggestionPanel =
    isEmailFocused &&
    searchTerm.length > 0 &&
    combinedEmailSuggestions.length > 0;

  const resetCompanyForm = () =>
    setCompanyForm({ name: "", abbreviation: "", startingNumber: 1 });

  const handleCompanySubmit = async (event) => {
    event.preventDefault();
    setCompanyFormError(null);
    setCompanySuccess(null);
    try {
      if (editingCompany) {
        await updateCompany(editingCompany.id, {
          name: companyForm.name.trim(),
          abbreviation: companyForm.abbreviation.trim().toUpperCase(),
          startingNumber: Number(companyForm.startingNumber) || 1,
        });
        setCompanySuccess(`Updated ${companyForm.name.trim()}`);
        setEditingCompany(null);
      } else {
        const created = await addCompany({
          name: companyForm.name.trim(),
          abbreviation: companyForm.abbreviation.trim().toUpperCase(),
          startingNumber: Number(companyForm.startingNumber) || 1,
        });
        setCompanySuccess(`Added ${created.name}`);
      }
      resetCompanyForm();
    } catch (err) {
      setCompanyFormError(err.message);
    }
  };

  const handleEditCompany = (company) => {
    setCompanyForm({
      name: company.name,
      abbreviation: company.abbreviation,
      startingNumber: company.startingNumber || 1,
    });
    setCompanySuccess(null);
    setCompanyFormError(null);
    setEditingCompany(company);
  };

  const handleDeleteCompany = async (company) => {
    if (!window.confirm(`Delete ${company.name}? This cannot be undone.`))
      return;
    try {
      await deleteCompany(company.id);
      if (editingCompany?.id === company.id) {
        cancelCompanyEdit();
      }
    } catch (err) {
      setCompanyFormError(err.message);
    }
  };

  const cancelCompanyEdit = () => {
    setEditingCompany(null);
    resetCompanyForm();
    setCompanyFormError(null);
    setCompanySuccess(null);
  };

  const toggleCompanySelection = (companyId) => {
    setUserForm((prev) => {
      const exists = prev.companyIds.includes(companyId);
      return {
        ...prev,
        companyIds: exists
          ? prev.companyIds.filter((id) => id !== companyId)
          : [...prev.companyIds, companyId],
      };
    });
  };

  const resetUserForm = () =>
    setUserForm({
      userPrincipalName: "",
      role: "Editor",
      companyIds: [],
    });

  const handleUserAccessSubmit = async (event) => {
    event.preventDefault();
    setUserFormError(null);
    setUserSuccess(null);
    try {
      if (editingUserEntry) {
        await updateUserAccessEntry(editingUserEntry.id, {
          userPrincipalName: userForm.userPrincipalName
            .trim()
            .toLowerCase(),
          role: userForm.role,
          companyIds: userForm.companyIds,
        });
        setUserSuccess("User access updated.");
        setEditingUserEntry(null);
      } else {
        await addUserAccessEntry({
          userPrincipalName: userForm.userPrincipalName
            .trim()
            .toLowerCase(),
          role: userForm.role,
          companyIds: userForm.companyIds,
        });
        setUserSuccess("User access saved.");
      }
      resetUserForm();
    } catch (err) {
      setUserFormError(err.message);
    }
  };

  const handleEditUserAccess = (entry) => {
    setEditingUserEntry(entry);
    setUserForm({
      userPrincipalName: entry.userPrincipalName || "",
      role: entry.role || "Editor",
      companyIds: entry.companyIds || [],
    });
    setUserSuccess(null);
    setUserFormError(null);
  };

  const cancelUserEdit = () => {
    setEditingUserEntry(null);
    resetUserForm();
    setUserFormError(null);
    setUserSuccess(null);
  };

  const handleDeleteUserAccess = async (itemId) => {
    if (!window.confirm("Remove this user access entry?")) return;
    try {
      await deleteUserAccessEntry(itemId);
      if (editingUserEntry?.id === itemId) {
        cancelUserEdit();
      }
    } catch (err) {
      setUserFormError(err.message);
    }
  };

  if (!userAccessResolved) {
    return (
      <div className="space-y-6">
        <div className="card p-6">
          <h1 className="page-title">Admin Panel</h1>
          <p className="page-subtitle">Checking your administrator access…</p>
          <LoadingSpinner size="small" text="Verifying permissions..." />
        </div>
      </div>
    );
  }

  if (!isAdmin) {
    return (
      <div className="space-y-6">
        <div className="card p-6 space-y-3">
          <h1 className="page-title">Admin Panel</h1>
          <p className="page-subtitle">
            You need administrator rights to configure companies and user
            access.
          </p>
          <p className="text-sm text-slate-500">
            Please contact an administrator if you believe this is an error.
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <h1 className="page-title">Admin Panel</h1>
          <p className="page-subtitle">
            Configure companies and manage user roles for letter numbering.
          </p>
        </div>
        <button
          onClick={() => {
            loadCompanies();
            loadUserAccess();
          }}
          className="inline-flex items-center gap-2 rounded-md border border-slate-200 px-3 py-2 text-sm text-slate-600 hover:bg-slate-50"
        >
          <RefreshCw size={16} />
          Refresh Data
        </button>
      </div>

      {error && (
        <div className="p-3 rounded-md bg-rose-50 border border-rose-200 text-rose-800">
          {error}
        </div>
      )}

      <div className="card p-6 space-y-4">
        <div>
          <h2 className="text-lg font-semibold text-slate-900">
            Manage Companies
          </h2>
          <p className="text-sm text-slate-500">
            Add a new company and abbreviation to issue references.
          </p>
        </div>
        <form onSubmit={handleCompanySubmit} className="space-y-4">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div>
              <label className={labelClass}>Company name</label>
              <input
                type="text"
                className={inputClass}
                value={companyForm.name}
                onChange={(e) =>
                  setCompanyForm((prev) => ({ ...prev, name: e.target.value }))
                }
                placeholder="Ease Engineering"
                required
              />
            </div>
            <div>
              <label className={labelClass}>Abbreviation</label>
              <input
                type="text"
                className={inputClass}
                value={companyForm.abbreviation}
                onChange={(e) =>
                  setCompanyForm((prev) => ({
                    ...prev,
                    abbreviation: e.target.value,
                  }))
                }
                placeholder="EASE"
                maxLength={8}
                required
              />
            </div>
            <div>
              <label className={labelClass}>Starting sequence</label>
              <input
                type="number"
                className={inputClass}
                min={1}
                value={companyForm.startingNumber}
                onChange={(e) =>
                  setCompanyForm((prev) => ({
                    ...prev,
                    startingNumber: e.target.value,
                  }))
                }
              />
            </div>
          </div>
          {companyFormError && (
            <div className="rounded-md bg-rose-50 border border-rose-200 px-3 py-2 text-sm text-rose-700">
              {companyFormError}
            </div>
          )}
          {companySuccess && (
            <div className="rounded-md bg-green-50 border border-green-200 px-3 py-2 text-sm text-green-700">
              {companySuccess}
            </div>
          )}
          <div className="flex items-center justify-end gap-3">
            {editingCompany && (
              <button
                type="button"
                className="text-sm text-slate-500 hover:text-slate-700"
                onClick={cancelCompanyEdit}
              >
                Cancel edit
              </button>
            )}
            <button
              type="submit"
              className={`${primaryButtonClass} disabled:opacity-50`}
              disabled={
                loading.creatingCompany ||
                loading.updatingCompany ||
                loading.deletingCompany
              }
            >
              {editingCompany
                ? loading.updatingCompany
                  ? "Updating..."
                  : "Save changes"
                : loading.creatingCompany
                ? "Saving..."
                : "Save company"}
            </button>
          </div>
        </form>

        <div className="mt-6 overflow-x-auto">
          <table className="min-w-full divide-y divide-slate-200 text-sm">
            <thead className="bg-slate-50 text-xs uppercase text-slate-500">
              <tr>
                <th className="px-3 py-2 text-left">Company</th>
                <th className="px-3 py-2 text-left">Abbreviation</th>
                <th className="px-3 py-2 text-left">Total Letters</th>
                <th className="px-3 py-2 text-left">Current Preview</th>
                <th className="px-3 py-2 text-right">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200 bg-white">
              {companies.map((company) => {
                const total = stats.get(company.id) || 0;
                const preview = getReferencePreview(
                  company.id,
                  new Date().getFullYear()
                );
                return (
                  <tr key={company.id}>
                    <td className="px-3 py-2">{company.name}</td>
                    <td className="px-3 py-2 font-mono text-slate-600">
                      {company.abbreviation || "--"}
                    </td>
                    <td className="px-3 py-2">
                      <span className="inline-flex items-center gap-1 rounded-full bg-slate-100 px-2 py-0.5 text-xs font-medium text-slate-700">
                        <Building2 size={12} />
                        {total}
                      </span>
                    </td>
                    <td className="px-3 py-2 font-mono text-blue-700">
                      {preview || "--"}
                    </td>
                    <td className="px-3 py-2 text-right space-x-2">
                      <button
                        type="button"
                        className="inline-flex items-center gap-1 rounded-md border border-slate-200 px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50"
                        onClick={() => handleEditCompany(company)}
                      >
                        <Edit3 size={14} />
                        Edit
                      </button>
                      <button
                        type="button"
                        className="inline-flex items-center gap-1 rounded-md border border-rose-200 px-2 py-1 text-xs font-medium text-rose-600 hover:bg-rose-50"
                        onClick={() => handleDeleteCompany(company)}
                        disabled={loading.deletingCompany}
                      >
                        <Trash2 size={14} />
                        Delete
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      <div className="card p-6 space-y-4">
        <div>
          <h2 className="text-lg font-semibold text-slate-900">
            Manage User Roles
          </h2>
          <p className="text-sm text-slate-500">
            Control who can administer the system and which companies they can
            access.
          </p>
        </div>
        <form onSubmit={handleUserAccessSubmit} className="space-y-4">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="relative">
              <label className={`${labelClass} flex items-center gap-2`}>
                Email
                {isSearchingEmails && (
                  <span className="text-[10px] font-normal text-slate-500 uppercase tracking-normal">
                    searching...
                  </span>
                )}
              </label>
              <input
                type="email"
                className={inputClass}
                value={userForm.userPrincipalName}
                onChange={(e) =>
                  setUserForm((prev) => ({
                    ...prev,
                    userPrincipalName: e.target.value,
                  }))
                }
                placeholder="user@contoso.com"
                onFocus={() => setIsEmailFocused(true)}
                onBlur={() => setTimeout(() => setIsEmailFocused(false), 120)}
                required
              />
              {shouldShowSuggestionPanel && (
                <div className="absolute z-10 mt-1 max-h-56 w-full overflow-y-auto rounded-md border border-slate-200 bg-white shadow-lg">
                  {combinedEmailSuggestions.map((item) => (
                    <button
                      type="button"
                      key={item.value}
                      className="w-full px-3 py-2 text-left text-sm hover:bg-slate-50"
                      onMouseDown={(event) => event.preventDefault()}
                      onClick={() => {
                        setUserForm((prev) => ({
                          ...prev,
                          userPrincipalName: item.value,
                        }));
                        setIsEmailFocused(false);
                      }}
                    >
                      <span className="font-medium text-slate-800">
                        {item.value}
                      </span>
                      {item.label !== item.value && (
                        <span className="block text-xs text-slate-500">
                          {item.label}
                        </span>
                      )}
                    </button>
                  ))}
                </div>
              )}
            </div>
            <div>
              <label className={labelClass}>Role</label>
              <select
                className={inputClass}
                value={userForm.role}
                onChange={(e) =>
                  setUserForm((prev) => ({ ...prev, role: e.target.value }))
                }
              >
                {roleOptions.map((role) => (
                  <option key={role} value={role}>
                    {role}
                  </option>
                ))}
              </select>
            </div>
            <div className="md:col-span-3">
              <label className={labelClass}>Allowed companies</label>
              <div className="rounded-md border border-slate-200 p-3 max-h-60 overflow-y-auto space-y-2 bg-white">
                {companies.length === 0 && (
                  <p className="text-sm text-slate-500">
                    No companies available yet.
                  </p>
                )}
                {companies.map((company) => {
                  const checked = userForm.companyIds.includes(company.id);
                  return (
                    <label
                      key={company.id}
                      className="flex items-center gap-2 text-sm text-slate-700"
                    >
                      <input
                        type="checkbox"
                        className="rounded border-slate-300 text-blue-600 focus:ring-blue-500"
                        checked={checked}
                        onChange={() => toggleCompanySelection(company.id)}
                      />
                      <span>
                        {company.name}{" "}
                        <span className="text-xs text-slate-500">
                          ({company.abbreviation || "--"})
                        </span>
                      </span>
                    </label>
                  );
                })}
              </div>
              <p className="mt-1 text-xs text-slate-500">
                Leave blank to grant access to every company.
              </p>
            </div>
          </div>
          {userFormError && (
            <div className="rounded-md bg-rose-50 border border-rose-200 px-3 py-2 text-sm text-rose-700">
              {userFormError}
            </div>
          )}
          {userSuccess && (
            <div className="rounded-md bg-green-50 border border-green-200 px-3 py-2 text-sm text-green-700">
              {userSuccess}
            </div>
          )}
          <div className="flex items-center justify-end gap-3">
            {editingUserEntry && (
              <button
                type="button"
                className="text-sm text-slate-500 hover:text-slate-700"
                onClick={cancelUserEdit}
              >
                Cancel edit
              </button>
            )}
            <button
              type="submit"
              className={`${primaryButtonClass} disabled:opacity-50`}
              disabled={
                loading.creatingUserAccess || loading.updatingUserAccess
              }
            >
              {loading.creatingUserAccess || loading.updatingUserAccess ? (
                loading.updatingUserAccess ? "Updating..." : "Saving..."
              ) : editingUserEntry ? (
                "Save changes"
              ) : (
                <span className="inline-flex items-center gap-2">
                  <UserPlus size={16} />
                  Save user access
                </span>
              )}
            </button>
          </div>
        </form>

        <div className="mt-6">
          {loading.userAccess ? (
            <LoadingSpinner size="small" text="Loading user assignments..." />
          ) : sortedUserAccess.length === 0 ? (
            <p className="text-sm text-slate-500">
              No user assignments have been created yet.
            </p>
          ) : (
            <div className="overflow-x-auto">
              <table className="min-w-full divide-y divide-slate-200 text-sm">
                <thead className="bg-slate-50 text-xs uppercase text-slate-500">
                  <tr>
                    <th className="px-3 py-2 text-left">User</th>
                    <th className="px-3 py-2 text-left">Role</th>
                    <th className="px-3 py-2 text-left">Companies</th>
                    <th className="px-3 py-2 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 bg-white">
                  {sortedUserAccess.map((entry) => (
                    <tr key={entry.id}>
                      <td className="px-3 py-2">
                        <div className="font-medium text-slate-900">
                          {entry.userPrincipalName || entry.title}
                        </div>
                        <div className="text-xs text-slate-500">
                          {entry.title}
                        </div>
                      </td>
                      <td className="px-3 py-2">
                        <span className="inline-flex items-center rounded-full bg-blue-50 px-2 py-0.5 text-xs font-semibold text-blue-700">
                          {entry.role}
                        </span>
                      </td>
                      <td className="px-3 py-2 text-sm text-slate-600">
                        {entry.companyIds.length === 0
                          ? "All companies"
                          : entry.companyNames.join(", ")}
                      </td>
                      <td className="px-3 py-2 text-right">
                        <button
                          type="button"
                          onClick={() => handleEditUserAccess(entry)}
                          className="inline-flex items-center gap-1 rounded-md border border-slate-200 px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50 mr-2"
                        >
                          <Edit3 size={14} />
                          Edit
                        </button>
                        <button
                          type="button"
                          onClick={() => handleDeleteUserAccess(entry.id)}
                          className="inline-flex items-center gap-1 rounded-md border border-rose-200 px-2 py-1 text-xs font-medium text-rose-600 hover:bg-rose-50"
                        >
                          <Trash2 size={14} />
                          Remove
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default AdminPanel;
