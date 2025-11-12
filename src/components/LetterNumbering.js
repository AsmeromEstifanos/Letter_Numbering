import React, { useEffect, useMemo, useState } from "react";
import { PlusCircle, Paperclip, Download, Edit3, Trash2, X } from "lucide-react";
import { useLetters } from "../context/LetterContext";
import LoadingSpinner from "./LoadingSpinner";

const formatDisplayDate = (value) => {
  if (!value) return "—";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleDateString("en-US", {
    year: "numeric",
    month: "short",
    day: "numeric",
  });
};

const inputClass =
  "w-full rounded-md border border-slate-200 bg-white px-3 py-2 text-sm text-slate-900 shadow-sm focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-500";
const textareaClass = `${inputClass} min-h-[90px]`;
const labelClass =
  "block text-xs font-semibold uppercase tracking-wide text-slate-600 mb-1";
const secondaryButtonClass =
  "inline-flex items-center gap-2 text-sm text-blue-600 hover:text-blue-700";
const primaryButtonClass =
  "inline-flex items-center justify-center rounded-md bg-blue-600 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2";

const LetterNumbering = () => {
  const {
    companies,
    letters,
    loading,
    error,
    addCompany,
    createLetter,
    updateLetter,
    getReferencePreview,
    loadLetterAttachments,
    downloadAttachment,
    canEditLetters,
    canManageCompanies,
  } = useLetters();

  const currentYear = new Date().getFullYear();
  const [formState, setFormState] = useState({
    companyId: "",
    letterDate: new Date().toISOString().slice(0, 10),
    recipientCompany: "",
    subject: "",
    preparedBy: "",
    notes: "",
    year: currentYear,
  });
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [formError, setFormError] = useState(null);
  const [successRef, setSuccessRef] = useState(null);
  const [showQuickAdd, setShowQuickAdd] = useState(false);
  const [newCompany, setNewCompany] = useState({
    name: "",
    abbreviation: "",
    startingNumber: 1,
  });
  const [attachments, setAttachments] = useState([]);
  const [expandedLetterId, setExpandedLetterId] = useState(null);
  const [downloadingId, setDownloadingId] = useState(null);
  const [editModalOpen, setEditModalOpen] = useState(false);
  const [editingLetter, setEditingLetter] = useState(null);
  const [editForm, setEditForm] = useState({
    letterDate: "",
    recipientCompany: "",
    subject: "",
    preparedBy: "",
    notes: "",
  });
  const [editExistingAttachments, setEditExistingAttachments] = useState([]);
  const [editAttachmentsToRemove, setEditAttachmentsToRemove] = useState([]);
  const [editNewAttachments, setEditNewAttachments] = useState([]);
  const [isEditLoadingAttachments, setIsEditLoadingAttachments] =
    useState(false);
  const [isEditSaving, setIsEditSaving] = useState(false);
  const [editError, setEditError] = useState(null);
  useEffect(() => {
    if (!canManageCompanies && showQuickAdd) {
      setShowQuickAdd(false);
    }
  }, [canManageCompanies, showQuickAdd]);

  useEffect(() => {
    if (!formState.companyId && companies.length > 0) {
      setFormState((prev) => ({
        ...prev,
        companyId: prev.companyId || companies[0].id,
      }));
    }
  }, [companies, formState.companyId]);

  const referencePreview = getReferencePreview(
    formState.companyId,
    formState.year
  );

  const filteredLetters = useMemo(() => {
    if (!formState.companyId) return letters;
    return letters.filter(
      (letter) => letter.companyId === formState.companyId
    );
  }, [letters, formState.companyId]);

  const handleChange = (field) => (event) => {
    const value = event.target.value;
    setFormState((prev) => ({
      ...prev,
      [field]: field === "year" ? Number(value) : value,
    }));
  };

  const handleFileChange = (event) => {
    const files = Array.from(event.target.files || []);
    if (!files.length) return;
    setAttachments((prev) => [...prev, ...files]);
    event.target.value = "";
  };

  const removeAttachment = (index) => {
    setAttachments((prev) => prev.filter((_, idx) => idx !== index));
  };

  const handleQuickAddCompany = async (event) => {
    event.preventDefault();
    if (!canManageCompanies) {
      setFormError("You do not have permission to add companies.");
      return;
    }
    try {
      setFormError(null);
      const created = await addCompany({
        name: newCompany.name.trim(),
        abbreviation: newCompany.abbreviation.trim().toUpperCase(),
        startingNumber: Number(newCompany.startingNumber) || 1,
      });
      setNewCompany({ name: "", abbreviation: "", startingNumber: 1 });
      setShowQuickAdd(false);
      setFormState((prev) => ({ ...prev, companyId: created.id }));
    } catch (err) {
      setFormError(err.message);
    }
  };

  const handleSubmit = async (event) => {
    event.preventDefault();
    if (!canEditLetters) {
      setFormError("You do not have permission to create letters.");
      return;
    }
    if (!formState.companyId) {
      setFormError("Please select a company.");
      return;
    }
    if (!formState.subject) {
      setFormError("Please enter a subject.");
      return;
    }
    setFormError(null);
    setSuccessRef(null);
    setIsSubmitting(true);
    try {
      const result = await createLetter({
        ...formState,
        attachments,
      });
      setSuccessRef(result.referenceNumber);
      setFormState((prev) => ({
        ...prev,
        recipientCompany: "",
        subject: "",
        preparedBy: "",
        notes: "",
      }));
      setAttachments([]);
    } catch (err) {
      setFormError(err.message);
    } finally {
      setIsSubmitting(false);
    }
  };

  const toggleAttachments = async (letter) => {
    if (expandedLetterId === letter.id) {
      setExpandedLetterId(null);
      return;
    }
    setExpandedLetterId(letter.id);
    if (
      !letter.attachmentsLoaded ||
      !letter.attachments ||
      letter.attachments.length === 0
    ) {
      try {
        await loadLetterAttachments(letter.id);
      } catch (err) {
        setFormError(err.message);
      }
    }
  };

  const handleDownloadAttachment = async (letterId, attachment) => {
    if (!attachment) return;
    const key = attachment.path || attachment.id || attachment.name;
    setDownloadingId(`${letterId}:${key}`);
    try {
      const url = await downloadAttachment(letterId, attachment);
      if (!url) throw new Error("Unable to open document.");
      window.open(url, "_blank", "noopener,noreferrer");
    } catch (err) {
      setFormError(err.message);
    } finally {
      setDownloadingId(null);
    }
  };

  const openEditModal = async (letter) => {
    setEditError(null);
    setEditingLetter(letter);
    setEditForm({
      letterDate: letter.letterDate
        ? letter.letterDate.slice(0, 10)
        : new Date().toISOString().slice(0, 10),
      recipientCompany: letter.recipientCompany || "",
      subject: letter.subject || "",
      preparedBy: letter.preparedBy || "",
      notes: letter.notes || "",
    });
    setEditAttachmentsToRemove([]);
    setEditNewAttachments([]);
    if (letter.attachmentsLoaded) {
      setEditExistingAttachments(letter.attachments || []);
    } else {
      setEditExistingAttachments([]);
    }
    setEditModalOpen(true);
    if (!letter.attachmentsLoaded) {
      setIsEditLoadingAttachments(true);
      try {
        const attachments = await loadLetterAttachments(letter.id);
        setEditExistingAttachments(attachments || []);
      } catch (err) {
        setEditError(err.message);
      } finally {
        setIsEditLoadingAttachments(false);
      }
    }
  };

  const closeEditModal = () => {
    setEditModalOpen(false);
    setEditingLetter(null);
    setEditAttachmentsToRemove([]);
    setEditNewAttachments([]);
    setEditExistingAttachments([]);
    setIsEditSaving(false);
    setEditError(null);
  };

  const handleEditFieldChange = (field) => (event) => {
    const value = event.target.value;
    setEditForm((prev) => ({
      ...prev,
      [field]: value,
    }));
  };

  const handleEditFileChange = (event) => {
    const files = Array.from(event.target.files || []);
    if (!files.length) return;
    setEditNewAttachments((prev) => [...prev, ...files]);
    event.target.value = "";
  };

  const removeNewEditAttachment = (index) =>
    setEditNewAttachments((prev) => prev.filter((_, idx) => idx !== index));

  const toggleAttachmentRemoval = (attachment) => {
    const key = attachment.path || attachment.id || attachment.name;
    setEditAttachmentsToRemove((prev) => {
      const exists = prev.some(
        (item) => (item.path || item.id || item.name) === key
      );
      if (exists) {
        return prev.filter(
          (item) => (item.path || item.id || item.name) !== key
        );
      }
      return [...prev, attachment];
    });
  };

  const handleEditSubmit = async (event) => {
    event.preventDefault();
    if (!editingLetter) return;
    setEditError(null);
    setIsEditSaving(true);
    try {
      await updateLetter({
        letterId: editingLetter.id,
        updates: {
          letterDate: editForm.letterDate,
          recipientCompany: editForm.recipientCompany,
          subject: editForm.subject,
          preparedBy: editForm.preparedBy,
          notes: editForm.notes,
        },
        attachmentsToAdd: editNewAttachments,
        attachmentsToRemove: editAttachmentsToRemove,
      });
      closeEditModal();
    } catch (err) {
      setEditError(err.message);
    } finally {
      setIsEditSaving(false);
    }
  };

  const letterTableLoading = loading.letters && letters.length === 0;

  return (
    <div className="space-y-6">
      <div className="mb-2">
        <h1 className="page-title">Letter Numbering</h1>
        <p className="page-subtitle">
          Create sequential reference numbers, capture recipients, and attach
          the final letter file.
        </p>
      </div>

      {(loading.letters || loading.companies) && letters.length === 0 ? (
        <LoadingSpinner text="Loading SharePoint data..." />
      ) : null}

      {error && (
        <div className="p-3 rounded-md bg-rose-50 border border-rose-200 text-rose-800">
          {error}
        </div>
      )}

      <div className="card p-6">
          <div className="flex items-center justify-between mb-4">
            <div>
              <h2 className="text-xl font-semibold text-slate-900">
                Generate New Reference Number
              </h2>
              <p className="text-sm text-slate-500">
                Select a company, review the next sequence, and capture the
                letter details.
              </p>
            </div>
            {canManageCompanies && (
              <button
                className={secondaryButtonClass}
                onClick={() => setShowQuickAdd((prev) => !prev)}
              >
                <PlusCircle size={16} />
                {showQuickAdd ? "Hide quick add" : "Quick add company"}
              </button>
            )}
          </div>

          {!canEditLetters && (
            <div className="mb-4 rounded-md border border-amber-200 bg-amber-50 px-3 py-2 text-sm text-amber-800">
              You currently have read-only access. Contact an administrator if
              you need to issue new references.
            </div>
          )}

          {canManageCompanies && showQuickAdd && (
            <form
              onSubmit={handleQuickAddCompany}
              className="mb-6 rounded-lg border border-slate-200 p-4 bg-slate-50 space-y-3"
            >
              <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                <div>
                  <label className={labelClass}>Company name</label>
                  <input
                    type="text"
                    className={inputClass}
                    value={newCompany.name}
                    onChange={(e) =>
                      setNewCompany((prev) => ({
                        ...prev,
                        name: e.target.value,
                      }))
                    }
                    required
                  />
                </div>
                <div>
                  <label className={labelClass}>Abbreviation</label>
                  <input
                    type="text"
                    className={inputClass}
                    value={newCompany.abbreviation}
                    onChange={(e) =>
                      setNewCompany((prev) => ({
                        ...prev,
                        abbreviation: e.target.value,
                      }))
                    }
                    maxLength={8}
                    required
                  />
                </div>
                <div>
                  <label className={labelClass}>Starting number</label>
                  <input
                    type="number"
                    className={inputClass}
                    min={1}
                    value={newCompany.startingNumber}
                    onChange={(e) =>
                      setNewCompany((prev) => ({
                        ...prev,
                        startingNumber: e.target.value,
                      }))
                    }
                  />
                </div>
              </div>
              <div className="flex items-center justify-end gap-3">
                <button
                  type="button"
                  className="text-sm text-slate-500 hover:text-slate-700"
                  onClick={() => setShowQuickAdd(false)}
                >
                  Cancel
                </button>
                <button type="submit" className={primaryButtonClass}>
                  Save company
                </button>
              </div>
            </form>
          )}

          <form className="space-y-4" onSubmit={handleSubmit}>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div>
                <label className={labelClass}>Company</label>
                <select
                  className={inputClass}
                  value={formState.companyId}
                  onChange={handleChange("companyId")}
                  required
                >
                  <option value="">Select company</option>
                  {companies.map((company) => (
                    <option key={company.id} value={company.id}>
                      {company.name} ({company.abbreviation || "N/A"})
                    </option>
                  ))}
                </select>
              </div>
              <div>
                <label className={labelClass}>Reference year</label>
                <input
                  type="number"
                  className={inputClass}
                  value={formState.year}
                  onChange={handleChange("year")}
                  min={2000}
                  max={9999}
                />
              </div>
              <div>
                <label className={labelClass}>Letter date</label>
                <input
                  type="date"
                  className={inputClass}
                  value={formState.letterDate}
                  onChange={handleChange("letterDate")}
                />
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className={labelClass}>Recipient company</label>
                <input
                  type="text"
                  className={inputClass}
                  value={formState.recipientCompany}
                  onChange={handleChange("recipientCompany")}
                />
              </div>
              <div>
                <label className={labelClass}>Prepared by</label>
                <input
                  type="text"
                  className={inputClass}
                  value={formState.preparedBy}
                  onChange={handleChange("preparedBy")}
                />
              </div>
            </div>

            <div>
              <label className={labelClass}>Subject</label>
              <input
                type="text"
                className={inputClass}
                value={formState.subject}
                onChange={handleChange("subject")}
                required
              />
        </div>

        <div>
          <label className={labelClass}>Notes / Description</label>
          <textarea
            className={textareaClass}
            rows={3}
            value={formState.notes}
            onChange={handleChange("notes")}
          />
        </div>

        <div>
          <label className={`${labelClass} flex items-center gap-2`}>
            Attach letter files
            <span className="text-xs text-slate-500">
              (stored in SharePoint under the Letters library)
            </span>
          </label>
          <input
            type="file"
            onChange={handleFileChange}
            className={inputClass}
            accept=".pdf,.doc,.docx,.txt"
            multiple
          />
          {attachments.length > 0 && (
            <div className="mt-2 rounded-md border border-slate-200 divide-y divide-slate-200">
              {attachments.map((file, index) => (
                <div
                  key={index}
                  className="flex items-center justify-between px-3 py-2 text-sm"
                >
                  <div className="flex items-center gap-2 overflow-hidden">
                    <Paperclip size={14} className="text-slate-500" />
                    <span className="truncate">{file.name}</span>
                  </div>
                  <button
                    type="button"
                    className="text-xs text-rose-600 hover:text-rose-700"
                    onClick={() => removeAttachment(index)}
                  >
                    Remove
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>

        <div className="rounded-md border border-blue-200 bg-blue-50 px-4 py-3 text-sm text-blue-900">
          <p className="font-medium">Next reference preview</p>
                <p className="font-mono text-lg">
                  {referencePreview || "Select a company to preview"}
                </p>
              </div>

            {formError && (
              <div className="rounded-md bg-rose-50 border border-rose-200 px-3 py-2 text-sm text-rose-700">
                {formError}
              </div>
            )}

            {successRef && (
              <div className="rounded-md bg-green-50 border border-green-200 px-3 py-2 text-sm text-green-700">
                Letter saved with reference{" "}
                <strong className="font-mono">{successRef}</strong>
              </div>
            )}

            <div className="flex items-center justify-end gap-3">
              <button
                type="submit"
                disabled={
                  isSubmitting || loading.creating || !canEditLetters
                }
                className={`${primaryButtonClass} disabled:opacity-50`}
              >
                {isSubmitting || loading.creating
                  ? "Saving..."
                  : canEditLetters
                  ? "Generate Reference"
                  : "Read-only access"}
              </button>
            </div>
          </form>
        </div>

      <div className="card p-6">
        <div className="mb-4">
          <h3 className="text-lg font-semibold text-slate-900">
            Letter Registry
          </h3>
          <p className="text-sm text-slate-500">
            Showing letters for{" "}
            <span className="font-semibold">
              {formState.companyId
                ? companies.find((c) => c.id === formState.companyId)?.name ||
                  "selected company"
                : "all companies"}
            </span>
            .
          </p>
        </div>

        {letterTableLoading ? (
          <LoadingSpinner text="Loading letters..." />
        ) : filteredLetters.length === 0 ? (
          <p className="text-sm text-slate-500 mt-4">
            No letters found for this selection.
          </p>
        ) : (
          <div className="mt-4 overflow-x-auto">
            <table className="min-w-full divide-y divide-slate-200 text-sm">
              <thead className="bg-slate-50 text-xs uppercase text-slate-500">
                <tr>
                  <th className="px-3 py-2 text-left font-medium">Date</th>
                  <th className="px-3 py-2 text-left font-medium">
                    Reference
                  </th>
                  <th className="px-3 py-2 text-left font-medium">
                    Recipient
                  </th>
                  <th className="px-3 py-2 text-left font-medium">Subject</th>
                  <th className="px-3 py-2 text-left font-medium">
                    Prepared By
                  </th>
                  <th className="px-3 py-2 text-left font-medium text-center">
                    Attachments
                  </th>
                  <th className="px-3 py-2 text-right font-medium">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 bg-white">
                {filteredLetters.map((letter) => (
                  <React.Fragment key={letter.id}>
                    <tr className="hover:bg-slate-50">
                      <td className="px-3 py-2 whitespace-nowrap">
                        {formatDisplayDate(letter.letterDate)}
                      </td>
                      <td className="px-3 py-2 font-mono text-slate-900">
                        {letter.referenceNumber}
                      </td>
                      <td className="px-3 py-2">
                        {letter.recipientCompany || "—"}
                      </td>
                      <td className="px-3 py-2 max-w-xs truncate">
                        {letter.subject || "—"}
                      </td>
                      <td className="px-3 py-2">
                        {letter.preparedBy || "—"}
                      </td>
                      <td className="px-3 py-2 text-center">
                    {!letter.attachmentsLoaded ||
                    (letter.attachments || []).length > 0 ? (
                      <button
                        className="inline-flex items-center gap-2 text-sm text-blue-600 hover:text-blue-700"
                        onClick={() => toggleAttachments(letter)}
                      >
                        <Paperclip size={14} />
                        {expandedLetterId === letter.id ? "Hide" : "View"}
                      </button>
                    ) : (
                      <span className="text-xs text-slate-400">None</span>
                    )}
                  </td>
                  <td className="px-3 py-2 text-right">
                    <button
                      className="inline-flex items-center gap-2 rounded-md border border-slate-200 px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-50"
                      onClick={() => openEditModal(letter)}
                      disabled={!canEditLetters}
                    >
                      <Edit3 size={14} />
                      Edit
                    </button>
                  </td>
                </tr>
                {expandedLetterId === letter.id && (
                  <tr>
                    <td colSpan={7} className="bg-slate-50 px-4 py-3">
                          {!letter.attachmentsLoaded ? (
                            <LoadingSpinner text="Loading documents..." />
                          ) : letter.attachments &&
                            letter.attachments.length > 0 ? (
                            <div className="flex flex-col gap-2">
                              {letter.attachments.map((attachment) => {
                                const attachmentKey =
                                  attachment.path ||
                                  attachment.id ||
                                  attachment.name;
                                return (
                                  <div
                                    key={attachmentKey}
                                    className="flex items-center justify-between rounded border border-slate-200 bg-white px-3 py-2 text-xs"
                                  >
                                    <div className="flex items-center gap-2">
                                      <Paperclip size={14} />
                                      <span>{attachment.name}</span>
                                    </div>
                                    <button
                                      className="inline-flex items-center gap-1 text-blue-600 hover:text-blue-700"
                                      onClick={() =>
                                        handleDownloadAttachment(
                                          letter.id,
                                          attachment
                                        )
                                      }
                                    >
                                      <Download size={14} />
                                      {downloadingId ===
                                      `${letter.id}:${attachmentKey}`
                                        ? "Opening..."
                                        : "Open"}
                                    </button>
                                  </div>
                                );
                              })}
                            </div>
                          ) : (
                            <p className="text-xs text-slate-500">
                              No documents uploaded for this reference.
                            </p>
                          )}
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                ))}
              </tbody>
            </table>
          </div>
        )}
      {editModalOpen && editingLetter && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 px-4">
          <div className="w-full max-w-3xl overflow-hidden rounded-xl bg-white shadow-2xl">
            <div className="flex items-center justify-between border-b border-slate-200 px-6 py-4">
              <div>
                <h3 className="text-lg font-semibold text-slate-900">
                  Edit Letter
                </h3>
                <p className="text-sm text-slate-500">
                  {editingLetter.referenceNumber}
                </p>
              </div>
              <button
                className="rounded-md border border-slate-200 p-2 text-slate-500 hover:text-slate-700"
                onClick={closeEditModal}
                type="button"
              >
                <X size={16} />
              </button>
            </div>
            <form
              onSubmit={handleEditSubmit}
              className="max-h-[80vh] overflow-y-auto px-6 py-4 space-y-4"
            >
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className={labelClass}>Letter date</label>
                  <input
                    type="date"
                    className={inputClass}
                    value={editForm.letterDate}
                    onChange={handleEditFieldChange("letterDate")}
                  />
                </div>
                <div>
                  <label className={labelClass}>Recipient company</label>
                  <input
                    type="text"
                    className={inputClass}
                    value={editForm.recipientCompany}
                    onChange={handleEditFieldChange("recipientCompany")}
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className={labelClass}>Subject</label>
                  <input
                    type="text"
                    className={inputClass}
                    value={editForm.subject}
                    onChange={handleEditFieldChange("subject")}
                    required
                  />
                </div>
                <div>
                  <label className={labelClass}>Prepared by</label>
                  <input
                    type="text"
                    className={inputClass}
                    value={editForm.preparedBy}
                    onChange={handleEditFieldChange("preparedBy")}
                  />
                </div>
              </div>

              <div>
                <label className={labelClass}>Notes</label>
                <textarea
                  className={textareaClass}
                  rows={3}
                  value={editForm.notes}
                  onChange={handleEditFieldChange("notes")}
                />
              </div>

              <div>
                <label className={labelClass}>Existing files</label>
                {isEditLoadingAttachments ? (
                  <LoadingSpinner size="small" text="Loading attachments..." />
                ) : editExistingAttachments.length === 0 ? (
                  <p className="text-sm text-slate-500">
                    No files have been uploaded yet.
                  </p>
                ) : (
                  <div className="space-y-2">
                    {editExistingAttachments.map((attachment) => {
                      const key =
                        attachment.path || attachment.id || attachment.name;
                      const marked = editAttachmentsToRemove.some(
                        (item) =>
                          (item.path || item.id || item.name) === key
                      );
                      return (
                        <div
                          key={key}
                          className={`flex items-center justify-between rounded border px-3 py-2 text-sm ${
                            marked
                              ? "border-amber-300 bg-amber-50"
                              : "border-slate-200 bg-white"
                          }`}
                        >
                          <div className="flex items-center gap-2">
                            <Paperclip
                              size={14}
                              className={marked ? "text-amber-700" : ""}
                            />
                            <span
                              className={
                                marked
                                  ? "text-amber-800 line-through"
                                  : "text-slate-800"
                              }
                            >
                              {attachment.name}
                            </span>
                          </div>
                          <button
                            type="button"
                            onClick={() => toggleAttachmentRemoval(attachment)}
                            className="inline-flex items-center gap-1 rounded-md border border-slate-200 px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50"
                          >
                            <Trash2 size={14} />
                            {marked ? "Keep file" : "Remove"}
                          </button>
                        </div>
                      );
                    })}
                    {editAttachmentsToRemove.length > 0 && (
                      <p className="text-xs text-amber-700">
                        Files marked as removed will be deleted once you save.
                      </p>
                    )}
                </div>
                )}
              </div>

              <div>
                <label className={labelClass}>Add files</label>
                <input
                  type="file"
                  multiple
                  className={inputClass}
                  onChange={handleEditFileChange}
                />
                {editNewAttachments.length > 0 && (
                  <div className="mt-2 space-y-2">
                    {editNewAttachments.map((file, index) => (
                      <div
                        key={`${file.name}-${index}`}
                        className="flex items-center justify-between rounded border border-slate-200 bg-white px-3 py-2 text-sm"
                      >
                        <span className="truncate">{file.name}</span>
                        <button
                          type="button"
                          className="inline-flex items-center gap-1 text-rose-600 hover:text-rose-700"
                          onClick={() => removeNewEditAttachment(index)}
                        >
                          <Trash2 size={14} />
                          Remove
                        </button>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              {editError && (
                <div className="rounded-md border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">
                  {editError}
                </div>
              )}

              <div className="flex items-center justify-end gap-3 border-t border-slate-200 pt-4">
                <button
                  type="button"
                  className="text-sm text-slate-600 hover:text-slate-800"
                  onClick={closeEditModal}
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className={`${primaryButtonClass} disabled:opacity-50`}
                  disabled={isEditSaving}
                >
                  {isEditSaving ? "Saving..." : "Save changes"}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
      </div>
    </div>
  );
};

export default LetterNumbering;
