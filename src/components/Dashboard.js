import React, { useMemo } from "react";
import { useLetters } from "../context/LetterContext";
import LoadingSpinner from "./LoadingSpinner";
import {
  Hash,
  Building2,
  FileText,
  Calendar,
  RefreshCw,
  Shield,
} from "lucide-react";

const Dashboard = () => {
  const {
    companies,
    letters,
    loading,
    error,
    refreshAll,
    getReferencePreview,
    isAdmin,
  } = useLetters();

  const currentYear = new Date().getFullYear();
  const stats = useMemo(() => {
    const totalLetters = letters.length;
    const thisYear = letters.filter(
      (letter) => Number(letter.year) === currentYear
    ).length;
    const uniqueRecipients = new Set(
      letters
        .map((letter) => letter.recipientCompany)
        .filter((recipient) => recipient)
    ).size;

    return {
      totalLetters,
      thisYear,
      uniqueRecipients,
    };
  }, [letters, currentYear]);

  const nextReferences = useMemo(
    () =>
      companies.slice(0, 4).map((company) => ({
        ...company,
        nextReference: getReferencePreview(company.id, currentYear),
      })),
    [companies, currentYear, getReferencePreview]
  );

  const extractSequence = (reference) => {
    if (!reference) return 0;
    const match = reference.match(/^[^/]+\/(\d+)\/\d{2}$/);
    return match ? Number(match[1]) : 0;
  };

  const latestLetters = useMemo(() => {
    if (letters.length === 0) return [];
    const sorted = [...letters].sort((a, b) => {
      const dateA = new Date(a.letterDate || a.createdDateTime || 0);
      const dateB = new Date(b.letterDate || b.createdDateTime || 0);
      const dateDiff = dateB - dateA;
      if (dateDiff !== 0) return dateDiff;
      return extractSequence(b.referenceNumber) - extractSequence(a.referenceNumber);
    });
    return sorted.slice(0, 5);
  }, [letters]);

  if (loading.letters && letters.length === 0) {
    return <LoadingSpinner text="Loading dashboard..." />;
  }

  return (
    <div className="space-y-6">
      <div className="mb-2 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <h1 className="page-title">Letter Numbering Dashboard</h1>
          <p className="page-subtitle">
            Generate, track, and audit outbound reference numbers across every
            company.
          </p>
        </div>
        <button
          onClick={refreshAll}
          className="inline-flex items-center gap-2 rounded-md border border-slate-200 px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-100"
        >
          <RefreshCw size={16} />
          Refresh data
        </button>
      </div>

      {error && (
        <div className="p-3 rounded-md bg-rose-50 border border-rose-200 text-rose-800">
          {error}
        </div>
      )}

      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <div className="card p-6">
          <div className="flex items-center">
            <div className="p-2 bg-blue-100 rounded-lg">
              <Building2 className="h-6 w-6 text-blue-600" />
            </div>
            <div className="ml-4">
              <p className="text-sm font-medium text-slate-600">Companies</p>
              <p className="text-2xl font-bold text-slate-900">
                {companies.length}
              </p>
            </div>
          </div>
        </div>

        <div className="card p-6">
          <div className="flex items-center">
            <div className="p-2 bg-green-100 rounded-lg">
              <Hash className="h-6 w-6 text-green-600" />
            </div>
            <div className="ml-4">
              <p className="text-sm font-medium text-slate-600">
                Letters This Year
              </p>
              <p className="text-2xl font-bold text-slate-900">
                {stats.thisYear}
              </p>
            </div>
          </div>
        </div>

        <div className="card p-6">
          <div className="flex items-center">
            <div className="p-2 bg-purple-100 rounded-lg">
              <FileText className="h-6 w-6 text-purple-600" />
            </div>
            <div className="ml-4">
              <p className="text-sm font-medium text-slate-600">
                Unique Recipients
              </p>
              <p className="text-2xl font-bold text-slate-900">
                {stats.uniqueRecipients}
              </p>
            </div>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="card p-6">
          <h3 className="text-lg font-semibold text-slate-900 mb-4">
            Quick Actions
          </h3>
          <div className="space-y-3">
            <button
              onClick={() => (window.location.href = "/letters")}
              className="w-full flex items-center justify-center px-4 py-2 border border-transparent rounded-md shadow-sm text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500"
            >
              <Hash className="h-4 w-4 mr-2" />
              Generate Reference Number
            </button>
            {isAdmin ? (
              <button
                onClick={() => (window.location.href = "/admin")}
                className="w-full flex items-center justify-center px-4 py-2 border border-slate-300 rounded-md shadow-sm text-sm font-medium text-slate-700 bg-white hover:bg-slate-50 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500"
              >
                <Shield className="h-4 w-4 mr-2" />
                Open Admin Panel
              </button>
            ) : (
              <div className="w-full rounded-md border border-amber-200 bg-amber-50 px-3 py-2 text-sm text-amber-800 text-center">
                Need admin rights to manage companies.
              </div>
            )}
          </div>
        </div>

        <div className="card p-6">
          <h3 className="text-lg font-semibold text-slate-900 mb-4">
            Next Reference Numbers
          </h3>
          <div className="space-y-2">
            {nextReferences.length > 0 ? (
              nextReferences.map((company) => (
                <div
                  key={company.id}
                  className="flex items-center justify-between rounded-md border border-slate-100 bg-slate-50 px-3 py-2"
                >
                  <div>
                    <p className="text-sm font-semibold text-slate-800">
                      {company.name}
                    </p>
                    <p className="text-xs text-slate-500">
                      Abbreviation: {company.abbreviation || "N/A"}
                    </p>
                  </div>
                  <span className="text-sm font-mono text-blue-700 bg-blue-50 px-2 py-1 rounded">
                    {company.nextReference || "Configure"}
                  </span>
                </div>
              ))
            ) : (
              <p className="text-sm text-slate-500">
                Add a company to start generating reference numbers.
              </p>
            )}
          </div>
        </div>
      </div>

      <div className="card p-6">
        <div className="flex items-center justify-between mb-4">
          <h3 className="text-lg font-semibold text-slate-900">
            Latest Letters
          </h3>
          <div className="flex items-center text-sm text-slate-500 gap-2">
            <Calendar size={16} />
            Tracking {letters.length} total records
          </div>
        </div>
        {letters.length === 0 ? (
          <p className="text-sm text-slate-500">
            No letters found. Create your first reference number from the
            Letters page.
          </p>
        ) : (
          <div className="space-y-3">
            {latestLetters.map((letter) => (
              <div
                key={letter.id}
                className="border border-slate-100 rounded-lg px-3 py-2 flex flex-col sm:flex-row sm:items-center sm:justify-between"
              >
                <div>
                  <p className="text-sm font-semibold text-slate-800">
                    {letter.referenceNumber}
                  </p>
                  <p className="text-xs text-slate-500">
                    {letter.recipientCompany || "Recipient TBD"}
                  </p>
                </div>
                <p className="text-xs text-slate-500 mt-1 sm:mt-0">
                  {new Date(letter.letterDate || letter.createdDateTime).toLocaleDateString()}
                </p>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};

export default Dashboard;
