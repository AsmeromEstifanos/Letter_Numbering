import React, { useMemo, useState } from "react";
import { Building2, Hash, RefreshCw } from "lucide-react";
import LoadingSpinner from "./LoadingSpinner";
import { useLetters } from "../context/LetterContext";

const inputClass =
  "w-full rounded-md border border-slate-200 bg-white px-3 py-2 text-sm text-slate-900 shadow-sm focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-500";
const labelClass =
  "block text-xs font-semibold uppercase tracking-wide text-slate-600 mb-1";
const primaryButtonClass =
  "inline-flex items-center justify-center rounded-md bg-blue-600 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2";

const Companies = () => {
  const {
    companies,
    letters,
    loading,
    error,
    addCompany,
    refreshAll,
    getReferencePreview,
    canManageCompanies,
  } = useLetters();

  const [form, setForm] = useState({
    name: "",
    abbreviation: "",
    startingNumber: 1,
  });
  const [formError, setFormError] = useState(null);
  const [success, setSuccess] = useState(null);

  const stats = useMemo(() => {
    const companyMap = new Map();
    letters.forEach((letter) => {
      const entry = companyMap.get(letter.companyId) || {
        total: 0,
        lastReference: null,
        lastDate: null,
      };
      entry.total += 1;
      if (
        !entry.lastDate ||
        new Date(letter.letterDate) > new Date(entry.lastDate)
      ) {
        entry.lastReference = letter.referenceNumber;
        entry.lastDate = letter.letterDate;
      }
      companyMap.set(letter.companyId, entry);
    });
    return companyMap;
  }, [letters]);

  const handleSubmit = async (event) => {
    event.preventDefault();
    setFormError(null);
    setSuccess(null);
    try {
      const created = await addCompany({
        name: form.name.trim(),
        abbreviation: form.abbreviation.trim().toUpperCase(),
        startingNumber: Number(form.startingNumber) || 1,
      });
      setForm({ name: "", abbreviation: "", startingNumber: 1 });
      setSuccess(`Added ${created.name}`);
    } catch (err) {
      setFormError(err.message);
    }
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <h1 className="page-title">Companies & Prefixes</h1>
          <p className="page-subtitle">
            Maintain the master list of companies that drive letter numbering
            prefixes.
          </p>
        </div>
        <button
          onClick={refreshAll}
          className="inline-flex items-center gap-2 rounded-md border border-slate-200 px-3 py-2 text-sm text-slate-600 hover:bg-slate-50"
        >
          <RefreshCw size={16} />
          Refresh SharePoint
        </button>
      </div>

      {error && (
        <div className="rounded-md border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">
          {error}
        </div>
      )}

      <div className="card p-6 space-y-4">
        <div>
          <h2 className="text-lg font-semibold text-slate-900">Add Company</h2>
          <p className="text-sm text-slate-500">
            The abbreviation becomes the reference prefix (e.g., EASEM/0001/24).
          </p>
        </div>
        {canManageCompanies ? (
          <form onSubmit={handleSubmit} className="space-y-4">
            <div>
              <label className={labelClass}>Company name</label>
              <input
                type="text"
                className={inputClass}
                value={form.name}
                onChange={(e) =>
                  setForm((prev) => ({ ...prev, name: e.target.value }))
                }
                placeholder="Ease Management PLC"
                required
              />
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div>
                <label className={labelClass}>Abbreviation</label>
                <input
                  type="text"
                  className={inputClass}
                  value={form.abbreviation}
                  onChange={(e) =>
                    setForm((prev) => ({
                      ...prev,
                      abbreviation: e.target.value,
                    }))
                  }
                  placeholder="EASEM"
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
                  value={form.startingNumber}
                  onChange={(e) =>
                    setForm((prev) => ({
                      ...prev,
                      startingNumber: e.target.value,
                    }))
                  }
                />
              </div>
            </div>
            {formError && (
              <div className="rounded-md border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">
                {formError}
              </div>
            )}
            {success && (
              <div className="rounded-md border border-green-200 bg-green-50 px-3 py-2 text-sm text-green-700">
                {success}
              </div>
            )}
            <div className="flex justify-end">
              <button
                type="submit"
                className={`${primaryButtonClass} disabled:opacity-50`}
                disabled={loading.creatingCompany}
              >
                {loading.creatingCompany ? "Saving..." : "Save company"}
              </button>
            </div>
          </form>
        ) : (
          <div className="rounded-md border border-amber-200 bg-amber-50 px-3 py-2 text-sm text-amber-800">
            You have read-only access. Contact an administrator to add or edit
            company prefixes.
          </div>
        )}
      </div>

      <div className="card p-6">
        <div className="flex items-center justify-between mb-4">
          <h3 className="text-lg font-semibold text-slate-900">
            Registered Companies ({companies.length})
          </h3>
          <span className="text-xs text-slate-500">
            {letters.length} letter records synced
          </span>
        </div>
        {loading.companies && companies.length === 0 ? (
          <LoadingSpinner text="Loading companies..." />
        ) : companies.length === 0 ? (
          <p className="text-sm text-slate-500">
            No companies found. Use the form above to add one.
          </p>
        ) : (
          <div className="overflow-x-auto">
            <table className="min-w-full divide-y divide-slate-200 text-sm">
              <thead className="bg-slate-50 text-xs uppercase text-slate-500">
                <tr>
                  <th className="px-3 py-2 text-left font-medium">Company</th>
                  <th className="px-3 py-2 text-left font-medium">
                    Abbreviation
                  </th>
                  <th className="px-3 py-2 text-left font-medium">
                    Starting #
                  </th>
                  <th className="px-3 py-2 text-left font-medium">Letters</th>
                  <th className="px-3 py-2 text-left font-medium">
                    Current Preview
                  </th>
                  <th className="px-3 py-2 text-left font-medium">
                    Last Reference
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 bg-white">
                {companies.map((company) => {
                  const companyStat = stats.get(company.id) || {
                    total: 0,
                    lastReference: "-",
                    lastDate: null,
                  };
                  const preview = getReferencePreview(
                    company.id,
                    new Date().getFullYear()
                  );
                  return (
                    <tr key={company.id} className="hover:bg-slate-50">
                      <td className="px-3 py-2 flex items-center gap-2">
                        <Building2 size={14} className="text-slate-400" />
                        <span className="font-medium text-slate-800">
                          {company.name}
                        </span>
                      </td>
                      <td className="px-3 py-2 font-mono text-slate-600">
                        {company.abbreviation || "—"}
                      </td>
                      <td className="px-3 py-2">{company.startingNumber}</td>
                      <td className="px-3 py-2">
                        <span className="inline-flex items-center gap-1 rounded-full bg-slate-100 px-2 py-0.5 text-xs font-medium text-slate-700">
                          <Hash size={12} />
                          {companyStat.total}
                        </span>
                      </td>
                      <td className="px-3 py-2 font-mono text-blue-700">
                        {preview || "Configure"}
                      </td>
                      <td className="px-3 py-2">
                        {companyStat.lastReference ? (
                          <div>
                            <p className="font-mono text-slate-800">
                              {companyStat.lastReference}
                            </p>
                            <p className="text-xs text-slate-500">
                              {companyStat.lastDate
                                ? new Date(
                                    companyStat.lastDate
                                  ).toLocaleDateString()
                                : ""}
                            </p>
                          </div>
                        ) : (
                          <span className="text-xs text-slate-400">—</span>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
};

export default Companies;
