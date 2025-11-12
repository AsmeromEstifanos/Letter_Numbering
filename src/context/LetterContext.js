import React, {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useReducer,
} from "react";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import {
  createSharePointList,
  createSharePointListItem,
  updateSharePointListItem,
  listSharePointListItems,
  ensureSharePointListColumns,
  ensureDriveFolderPath,
  uploadBinaryFileToDrive,
  listDriveFiles,
  getDriveItemViewUrl,
  deleteSharePointListItem,
  searchAzureUsers,
  deleteDriveFile,
} from "../services/sharepoint";

const LetterContext = createContext(null);

const LETTER_LIST_IDENTIFIER =
  process.env.REACT_APP_LETTER_LIST_ID ||
  process.env.REACT_APP_LETTER_LIST_NAME ||
  "LetterNumbers";
const COMPANY_LIST_IDENTIFIER =
  process.env.REACT_APP_COMPANY_LIST_ID ||
  process.env.REACT_APP_COMPANY_LIST_NAME ||
  "LetterCompanies";

const SHAREPOINT_SITE_URL =
  process.env.REACT_APP_MEETING_SITE_URL ||
  process.env.REACT_APP_SHAREPOINT_SITE_URL ||
  "";

const LETTER_LIBRARY_NAME =
  process.env.REACT_APP_LETTER_LIBRARY_NAME ||
  process.env.REACT_APP_MEETING_LIBRARY_NAME ||
  "Documents";

const LETTER_LIBRARY_ROOT = (
  process.env.REACT_APP_LETTER_LIBRARY_ROOT || ""
).trim();

const COMPANY_COLUMNS = [
  { name: "Abbreviation", text: {} },
  { name: "StartingNumber", number: {} },
  { name: "Color", text: {} },
];

const LETTER_COLUMNS = [
  { name: "ReferenceNumber", text: {} },
  { name: "CompanyItemId", text: {} },
  { name: "CompanyName", text: {} },
  { name: "CompanyAbbreviation", text: {} },
  { name: "SequenceNumber", number: {} },
  { name: "Year", number: {} },
  { name: "LetterDate", dateTime: { format: "dateOnly" } },
  { name: "RecipientCompany", text: {} },
  { name: "Subject", text: {} },
  { name: "PreparedBy", text: {} },
  { name: "Notes", multilineText: {} },
];

const USER_ACCESS_LIST_IDENTIFIER =
  process.env.REACT_APP_USER_ACCESS_LIST_ID ||
  process.env.REACT_APP_USER_ACCESS_LIST_NAME ||
  "LetterUserAccess";

const USER_ACCESS_COLUMNS = [
  { name: "UserPrincipalName", text: {} },
  {
    name: "Role",
    choice: { choices: ["Admin", "Editor", "Viewer"], allowTextEntry: true },
  },
  { name: "CompanyIds", text: {} },
  { name: "CompanyNames", text: {} },
];

const initialState = {
  loading: {
    companies: false,
    letters: false,
    creating: false,
    creatingCompany: false,
    updatingCompany: false,
    deletingCompany: false,
    userAccess: false,
    creatingUserAccess: false,
    updatingUserAccess: false,
  },
  error: null,
  companies: [],
  letters: [],
  userAccess: [],
  userAccessLoaded: false,
  initialized: false,
};

const letterReducer = (state, action) => {
  switch (action.type) {
    case "SET_LOADING":
      return {
        ...state,
        loading: {
          ...state.loading,
          [action.payload.key]: action.payload.value,
        },
      };
    case "SET_ERROR":
      return { ...state, error: action.payload };
    case "SET_COMPANIES":
      return { ...state, companies: action.payload };
    case "ADD_COMPANY":
      return { ...state, companies: [...state.companies, action.payload] };
    case "UPDATE_COMPANY":
      return {
        ...state,
        companies: state.companies.map((company) =>
          company.id === action.payload.id ? action.payload.company : company
        ),
      };
    case "REMOVE_COMPANY":
      return {
        ...state,
        companies: state.companies.filter(
          (company) => company.id !== action.payload
        ),
      };
    case "SET_LETTERS":
      return { ...state, letters: action.payload };
    case "ADD_LETTER":
      return { ...state, letters: [action.payload, ...state.letters] };
    case "UPDATE_LETTER":
      return {
        ...state,
        letters: state.letters.map((letter) =>
          letter.id === action.payload.id ? action.payload.letter : letter
        ),
      };
    case "SET_USER_ACCESS":
      return {
        ...state,
        userAccess: action.payload,
        userAccessLoaded: true,
      };
    case "ADD_USER_ACCESS":
      return { ...state, userAccess: [action.payload, ...state.userAccess] };
    case "REMOVE_USER_ACCESS":
      return {
        ...state,
        userAccess: state.userAccess.filter(
          (entry) => entry.id !== action.payload
        ),
      };
    case "UPDATE_USER_ACCESS":
      return {
        ...state,
        userAccess: state.userAccess.map((entry) =>
          entry.id === action.payload.id ? action.payload.entry : entry
        ),
      };
    case "SET_INITIALIZED":
      return { ...state, initialized: action.payload };
    case "SET_USER_ACCESS_READY":
      return { ...state, userAccessLoaded: true };
    case "SET_LETTER_ATTACHMENTS":
      return {
        ...state,
        letters: state.letters.map((letter) =>
          letter.id === action.payload.id
            ? {
                ...letter,
                attachmentsLoaded: true,
                attachments: action.payload.attachments,
                hasAttachments:
                  (action.payload.attachments || []).length > 0,
              }
            : letter
        ),
      };
    default:
      return state;
  }
};

const twoDigitYear = (year) => String(year).slice(-2);
const isGuid = (value) => /^[0-9a-fA-F-]{20,}$/.test((value || "").trim());

const sanitizeSegment = (value, fallback = "General") => {
  if (!value || typeof value !== "string") return fallback;
  const cleaned = value
    .trim()
    .replace(/[<>:"/\\|?*\r\n]+/g, "")
    .replace(/\s+/g, "-");
  return cleaned || fallback;
};

const buildLetterFolderPath = (company) => {
  const companySegment = sanitizeSegment(
    company?.abbreviation || company?.name,
    "Company"
  );
  if (LETTER_LIBRARY_ROOT) {
    const rootSegment = sanitizeSegment(LETTER_LIBRARY_ROOT, "Letters");
    return `${rootSegment}/${companySegment}`;
  }
  return companySegment;
};

const buildStoredFileName = (reference, originalName = "document") => {
  const extensionMatch = (originalName || "").match(/(\.[^.\s]{1,10})$/);
  const extension = extensionMatch ? extensionMatch[1] : "";
  const safeReference = sanitizeSegment(
    (reference || "REF").replace(/[/\\]/g, "-"),
    "REF"
  );
  return `${safeReference}${extension}`.replace(/-+/g, "-");
};

const buildReferenceNumber = (abbr, sequenceNumber, year) => {
  if (!abbr || !sequenceNumber || !year) return "";
  const seq = String(sequenceNumber).padStart(4, "0");
  return `${abbr}/${seq}/${twoDigitYear(year)}`;
};

const normalizeLetterDate = (value) => {
  if (!value) return new Date().toISOString();
  if (typeof value === "string" && value.length <= 10) {
    return `${value}T00:00:00Z`;
  }
  return value;
};

const extractSequenceFromRef = (reference) => {
  if (!reference) return 0;
  const match = reference.match(/^[^/]+\/(\d+)\/\d{2}$/);
  return match ? Number(match[1]) : 0;
};

const extractYearFromRef = (reference) => {
  if (!reference) return 0;
  const match = reference.match(/\/(\d{2})$/);
  if (!match) return 0;
  const currentYear = new Date().getFullYear();
  const century = Math.floor(currentYear / 100) * 100;
  let year = century + Number(match[1]);
  if (year > currentYear + 10) {
    year -= 100;
  }
  return year;
};

const mapCompany = (item) => {
  const fields = item.fields || {};
  return {
    id: item.id,
    name: fields.Title || fields.CompanyName || "Unnamed Company",
    abbreviation:
      fields.Abbreviation || fields.CompanyAbbreviation || fields.Code || "",
    startingNumber: Number(fields.StartingNumber || fields.SequenceSeed || 1),
    color: fields.Color || "#2563eb",
    raw: fields,
  };
};

const mapLetter = (item) => {
  const fields = item.fields || {};
  const reference =
    fields.ReferenceNumber ||
    buildReferenceNumber(
      fields.CompanyAbbreviation,
      fields.SequenceNumber,
      fields.Year
    ) ||
    fields.Title;
  const sequence =
    Number(fields.SequenceNumber) || extractSequenceFromRef(reference);
  const year = Number(fields.Year) || extractYearFromRef(reference);

  return {
    id: item.id,
    referenceNumber: reference,
    companyId: fields.CompanyItemId || "",
    companyName: fields.CompanyName || "",
    companyAbbreviation: fields.CompanyAbbreviation || "",
    sequenceNumber: sequence || 0,
    year: year || new Date().getFullYear(),
    letterDate: fields.LetterDate || item.createdDateTime,
    recipientCompany: fields.RecipientCompany || "",
    subject: fields.Subject || reference,
    preparedBy: fields.PreparedBy || "",
    notes: fields.Notes || "",
    hasAttachments: false,
    attachmentsLoaded: false,
    attachments: [],
    webUrl: item.webUrl,
    createdDateTime: item.createdDateTime,
    raw: fields,
  };
};

export const LetterProvider = ({ children }) => {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const [state, dispatch] = useReducer(letterReducer, initialState);
  const userAccessResolved = state.userAccessLoaded;
  const activeAccount = useMemo(() => {
    const current = instance.getActiveAccount();
    if (current) return current;
    if (accounts && accounts.length > 0) {
      return accounts[0];
    }
    return null;
  }, [instance, accounts]);

  const currentUserPrincipalName = useMemo(() => {
    if (!activeAccount) return "";
    const preferred =
      activeAccount.username ||
      activeAccount.idTokenClaims?.preferred_username ||
      activeAccount.idTokenClaims?.email ||
      "";
    return (preferred || "").toLowerCase();
  }, [activeAccount]);

  const currentUserAccess = useMemo(() => {
    if (!currentUserPrincipalName) return null;
    return (
      state.userAccess.find((entry) => {
        const candidate =
          entry.userPrincipalName || entry.title || "";
        return candidate?.toLowerCase() === currentUserPrincipalName;
      }) || null
    );
  }, [state.userAccess, currentUserPrincipalName]);

  const hasAssignments = state.userAccess.length > 0;
  const baseRole = !userAccessResolved
    ? "Viewer"
    : currentUserAccess
    ? currentUserAccess.role
    : hasAssignments
    ? "Viewer"
    : "Admin";
  const currentRole = (baseRole || "Viewer").toLowerCase();
  const isAdmin = currentRole === "admin";
  const canEditLetters = currentRole === "admin" || currentRole === "editor";
  const canManageCompanies = isAdmin;

  const allowedCompanySet = useMemo(() => {
    if (!userAccessResolved || !currentUserAccess || isAdmin) return null;
    const ids = (currentUserAccess.companyIds || [])
      .map((id) => (id !== undefined && id !== null ? String(id).trim() : ""))
      .filter(Boolean);
    if (ids.length === 0) return null;
    return new Set(ids);
  }, [currentUserAccess, isAdmin, userAccessResolved]);

  const scopedCompanies = useMemo(() => {
    if (!allowedCompanySet) return state.companies;
    return state.companies.filter((company) =>
      allowedCompanySet.has(String(company.id))
    );
  }, [state.companies, allowedCompanySet]);

  const scopedLetters = useMemo(() => {
    if (!allowedCompanySet) return state.letters;
    return state.letters.filter((letter) =>
      allowedCompanySet.has(String(letter.companyId))
    );
  }, [state.letters, allowedCompanySet]);

  const setLoading = (key, value) =>
    dispatch({ type: "SET_LOADING", payload: { key, value } });

  const handleError = useCallback((error) => {
    console.error("[Letters] Graph error", error);
    dispatch({ type: "SET_ERROR", payload: error.message });
  }, []);

  const isListMissingError = (error) =>
    typeof error?.message === "string" &&
    error.message.toLowerCase().includes("list not found");

  const provisionCompanyList = useCallback(async () => {
    if (
      !COMPANY_LIST_IDENTIFIER ||
      isGuid(COMPANY_LIST_IDENTIFIER) ||
      !SHAREPOINT_SITE_URL
    ) {
      return;
    }

    await createSharePointList(instance, SHAREPOINT_SITE_URL, {
      displayName: COMPANY_LIST_IDENTIFIER,
      description: "Auto-generated by Letter Numbering to store companies.",
      columns: COMPANY_COLUMNS,
    });
  }, [instance]);

  const provisionLetterList = useCallback(async () => {
    if (
      !LETTER_LIST_IDENTIFIER ||
      isGuid(LETTER_LIST_IDENTIFIER) ||
      !SHAREPOINT_SITE_URL
    ) {
      return;
    }

    await createSharePointList(instance, SHAREPOINT_SITE_URL, {
      displayName: LETTER_LIST_IDENTIFIER,
      description: "Auto-generated by Letter Numbering to store letters.",
      columns: LETTER_COLUMNS,
    });
  }, [instance]);

  const provisionUserAccessList = useCallback(async () => {
    if (
      !USER_ACCESS_LIST_IDENTIFIER ||
      isGuid(USER_ACCESS_LIST_IDENTIFIER) ||
      !SHAREPOINT_SITE_URL
    ) {
      return;
    }

    await createSharePointList(instance, SHAREPOINT_SITE_URL, {
      displayName: USER_ACCESS_LIST_IDENTIFIER,
      description: "Stores user roles and company access for letter numbering.",
      columns: USER_ACCESS_COLUMNS,
    });
  }, [instance]);

  const ensureCompanyColumns = useCallback(async () => {
    if (!SHAREPOINT_SITE_URL || !COMPANY_LIST_IDENTIFIER) return;
    try {
      await ensureSharePointListColumns(
        instance,
        SHAREPOINT_SITE_URL,
        COMPANY_LIST_IDENTIFIER,
        COMPANY_COLUMNS
      );
    } catch (error) {
      console.warn("[SharePoint] Failed to ensure company columns:", error);
    }
  }, [instance]);

  const ensureLetterColumns = useCallback(async () => {
    if (!SHAREPOINT_SITE_URL || !LETTER_LIST_IDENTIFIER) return;
    try {
      await ensureSharePointListColumns(
        instance,
        SHAREPOINT_SITE_URL,
        LETTER_LIST_IDENTIFIER,
        LETTER_COLUMNS
      );
    } catch (error) {
      console.warn("[SharePoint] Failed to ensure letter columns:", error);
    }
  }, [instance]);

  const ensureUserAccessColumns = useCallback(async () => {
    if (!SHAREPOINT_SITE_URL || !USER_ACCESS_LIST_IDENTIFIER) return;
    try {
      await ensureSharePointListColumns(
        instance,
        SHAREPOINT_SITE_URL,
        USER_ACCESS_LIST_IDENTIFIER,
        USER_ACCESS_COLUMNS
      );
    } catch (error) {
      console.warn("[SharePoint] Failed to ensure user access columns:", error);
    }
  }, [instance]);

  const loadCompanies = useCallback(async () => {
    if (!SHAREPOINT_SITE_URL) {
      const err = new Error(
        "SharePoint site URL is missing. Set REACT_APP_SHAREPOINT_SITE_URL."
      );
      handleError(err);
      return [];
    }
    setLoading("companies", true);
    try {
      await ensureCompanyColumns();
      const fetchCompanies = async () =>
        await listSharePointListItems(
          instance,
          SHAREPOINT_SITE_URL,
          COMPANY_LIST_IDENTIFIER,
          {
            expand: "fields($select=Title,Abbreviation,StartingNumber,Color)",
            orderBy: "fields/Title asc",
            top: 500,
          }
        );

      const items = await fetchCompanies();
      const companies = items
        .map(mapCompany)
        .sort((a, b) => a.name.localeCompare(b.name));
      dispatch({ type: "SET_COMPANIES", payload: companies });
      dispatch({ type: "SET_ERROR", payload: null });
      return companies;
    } catch (error) {
      if (
        isListMissingError(error) &&
        COMPANY_LIST_IDENTIFIER &&
        !isGuid(COMPANY_LIST_IDENTIFIER)
      ) {
        try {
          await provisionCompanyList();
          const items = await listSharePointListItems(
            instance,
            SHAREPOINT_SITE_URL,
            COMPANY_LIST_IDENTIFIER,
            {
              expand:
                "fields($select=Title,Abbreviation,StartingNumber,Color)",
              orderBy: "fields/Title asc",
              top: 500,
            }
          );
          const companies = items
            .map(mapCompany)
            .sort((a, b) => a.name.localeCompare(b.name));
          dispatch({ type: "SET_COMPANIES", payload: companies });
          dispatch({ type: "SET_ERROR", payload: null });
          return companies;
        } catch (innerError) {
          handleError(innerError);
          throw innerError;
        }
      }
      handleError(error);
      throw error;
    } finally {
      setLoading("companies", false);
    }
  }, [instance, handleError, provisionCompanyList, ensureCompanyColumns]);

  const loadLetters = useCallback(
    async (options = {}) => {
      if (!SHAREPOINT_SITE_URL) {
        const err = new Error(
          "SharePoint site URL is missing. Set REACT_APP_SHAREPOINT_SITE_URL."
        );
        handleError(err);
        return [];
      }
      setLoading("letters", true);
    try {
      await ensureLetterColumns();
      const fetchLetters = async () =>
        await listSharePointListItems(
          instance,
          SHAREPOINT_SITE_URL,
          LETTER_LIST_IDENTIFIER,
            {
              expand:
                "fields($select=Title,ReferenceNumber,CompanyItemId,CompanyName,CompanyAbbreviation,SequenceNumber,Year,LetterDate,RecipientCompany,Subject,PreparedBy,Notes,Attachments)",
              orderBy: "fields/LetterDate desc",
              top: options.top || 500,
              fetchAll: options.fetchAll || false,
            }
          );

        const items = await fetchLetters();
        const letters = items
          .map(mapLetter)
          .sort(
            (a, b) =>
              new Date(b.letterDate || b.createdDateTime) -
              new Date(a.letterDate || a.createdDateTime)
          );
        dispatch({ type: "SET_LETTERS", payload: letters });
        dispatch({ type: "SET_ERROR", payload: null });
        return letters;
      } catch (error) {
        if (
          isListMissingError(error) &&
          LETTER_LIST_IDENTIFIER &&
          !isGuid(LETTER_LIST_IDENTIFIER)
        ) {
          try {
            await provisionLetterList();
            const retryItems = await listSharePointListItems(
              instance,
              SHAREPOINT_SITE_URL,
              LETTER_LIST_IDENTIFIER,
              {
                expand:
                  "fields($select=Title,ReferenceNumber,CompanyItemId,CompanyName,CompanyAbbreviation,SequenceNumber,Year,LetterDate,RecipientCompany,Subject,PreparedBy,Notes,Attachments)",
                orderBy: "fields/LetterDate desc",
                top: options.top || 500,
                fetchAll: options.fetchAll || false,
              }
            );
            const letters = retryItems
              .map(mapLetter)
              .sort(
                (a, b) =>
                  new Date(b.letterDate || b.createdDateTime) -
                  new Date(a.letterDate || a.createdDateTime)
              );
            dispatch({ type: "SET_LETTERS", payload: letters });
            dispatch({ type: "SET_ERROR", payload: null });
            return letters;
          } catch (innerError) {
            handleError(innerError);
            throw innerError;
          }
        }
        handleError(error);
        throw error;
      } finally {
        setLoading("letters", false);
      }
    },
    [instance, handleError, provisionLetterList, ensureLetterColumns]
  );

  const loadUserAccess = useCallback(async () => {
    if (!SHAREPOINT_SITE_URL || !USER_ACCESS_LIST_IDENTIFIER) {
      dispatch({ type: "SET_USER_ACCESS_READY" });
      return [];
    }
    setLoading("userAccess", true);
    const fetchEntries = async () => {
      const items = await listSharePointListItems(
        instance,
        SHAREPOINT_SITE_URL,
        USER_ACCESS_LIST_IDENTIFIER,
        {
          expand:
            "fields($select=Title,UserPrincipalName,Role,CompanyIds,CompanyNames)",
          orderBy: "fields/Title asc",
        }
      );
      return items.map((item) => {
        const fields = item.fields || {};
        const companyIds = (fields.CompanyIds || "")
          .split(",")
          .map((id) => id.trim())
          .filter(Boolean);
        const companyNames = (fields.CompanyNames || "")
          .split(",")
          .map((name) => name.trim())
          .filter(Boolean);
        return {
          id: item.id,
          title: fields.Title || fields.UserPrincipalName || "",
          userPrincipalName: fields.UserPrincipalName || "",
          role: fields.Role || "Viewer",
          companyIds,
          companyNames,
          raw: fields,
        };
      });
    };
    try {
      await ensureUserAccessColumns();
      const parsed = await fetchEntries();
      dispatch({ type: "SET_USER_ACCESS", payload: parsed });
      return parsed;
    } catch (error) {
      if (
        isListMissingError(error) &&
        USER_ACCESS_LIST_IDENTIFIER &&
        !isGuid(USER_ACCESS_LIST_IDENTIFIER)
      ) {
        try {
          await provisionUserAccessList();
          await ensureUserAccessColumns();
          const parsed = await fetchEntries();
          dispatch({ type: "SET_USER_ACCESS", payload: parsed });
          return parsed;
        } catch (innerError) {
          handleError(innerError);
          throw innerError;
        }
      }
      handleError(error);
      throw error;
    } finally {
      dispatch({ type: "SET_USER_ACCESS_READY" });
      setLoading("userAccess", false);
    }
  }, [
    instance,
    handleError,
    ensureUserAccessColumns,
    provisionUserAccessList,
  ]);

  const refreshAll = useCallback(async () => {
    await Promise.all([loadCompanies(), loadLetters(), loadUserAccess()]);
  }, [loadCompanies, loadLetters, loadUserAccess]);

  useEffect(() => {
    let isMounted = true;
    if (isAuthenticated) {
      dispatch({ type: "SET_INITIALIZED", payload: false });
      refreshAll()
        .catch(() => {})
        .finally(() => {
          if (isMounted) {
            dispatch({ type: "SET_INITIALIZED", payload: true });
          }
        });
    } else {
      dispatch({ type: "SET_INITIALIZED", payload: false });
    }
    return () => {
      isMounted = false;
    };
  }, [isAuthenticated, refreshAll]);

  const getCompany = useCallback(
    (companyId) => {
      const company =
        state.companies.find((item) => item.id === companyId) || null;
      if (
        company &&
        allowedCompanySet &&
        !allowedCompanySet.has(String(company.id))
      ) {
        return null;
      }
      return company;
    },
    [state.companies, allowedCompanySet]
  );

  const getNextSequence = useCallback(
    (companyId, year) => {
      if (!companyId || !year) return null;
      if (
        allowedCompanySet &&
        !allowedCompanySet.has(String(companyId))
      ) {
        return null;
      }
      const relevant = state.letters.filter(
        (letter) =>
          letter.companyId === companyId &&
          Number(letter.year) === Number(year)
      );
      if (relevant.length === 0) {
        const company = getCompany(companyId);
        return company?.startingNumber || 1;
      }
      const maxExisting = Math.max(
        ...relevant.map((letter) => letter.sequenceNumber || 0),
        0
      );
      return maxExisting + 1;
    },
    [state.letters, getCompany, allowedCompanySet]
  );

  const getReferencePreview = useCallback(
    (companyId, year) => {
      const company = getCompany(companyId);
      if (!company || !year) return "";
      const nextSequence = getNextSequence(companyId, year);
      if (!nextSequence) return "";
      return buildReferenceNumber(company.abbreviation, nextSequence, year);
    },
    [getCompany, getNextSequence]
  );

  const uploadLetterFiles = useCallback(
    async (company, referenceNumber, files = []) => {
      if (!files || files.length === 0) return [];
      const companyInfo = {
        abbreviation:
          company?.abbreviation ||
          company?.companyAbbreviation ||
          company?.companyAbbrev ||
          "",
        name: company?.name || company?.companyName || "Company",
      };
      const folderPath = buildLetterFolderPath(companyInfo);
      await ensureDriveFolderPath(
        instance,
        SHAREPOINT_SITE_URL,
        LETTER_LIBRARY_NAME,
        folderPath
      );
      const uploaded = [];
      for (const file of files) {
        const storedName = buildStoredFileName(referenceNumber, file.name);
        const targetPath = `${folderPath}/${storedName}`;
        const uploadResult = await uploadBinaryFileToDrive(
          instance,
          SHAREPOINT_SITE_URL,
          LETTER_LIBRARY_NAME,
          targetPath,
          file,
          file.type || "application/octet-stream"
        );
        let webUrl = uploadResult?.webUrl;
        if (!webUrl) {
          webUrl = await getDriveItemViewUrl(
            instance,
            SHAREPOINT_SITE_URL,
            LETTER_LIBRARY_NAME,
            targetPath
          );
        }
        uploaded.push({
          id: uploadResult?.id || targetPath,
          name: uploadResult?.name || storedName,
          path: targetPath,
          size: uploadResult?.size || file.size,
          webUrl,
        });
      }
      return uploaded;
    },
    [instance]
  );

  const deleteLetterFiles = useCallback(
    async (attachments = []) => {
      if (!attachments || attachments.length === 0) return;
      for (const attachment of attachments) {
        const path = attachment?.path;
        if (!path) continue;
        try {
          await deleteDriveFile(
            instance,
            SHAREPOINT_SITE_URL,
            LETTER_LIBRARY_NAME,
            path
          );
        } catch (error) {
          console.warn(
            "[SharePoint] Failed to delete attachment:",
            path,
            error
          );
        }
      }
    },
    [instance]
  );

  const addCompany = useCallback(
    async ({ name, abbreviation, startingNumber = 1 }) => {
      if (!name || !abbreviation) {
        throw new Error("Company name and abbreviation are required.");
      }
      if (!canManageCompanies) {
        throw new Error("You do not have permission to add companies.");
      }
      setLoading("creatingCompany", true);
      try {
        await ensureCompanyColumns();
        const fields = {
          Title: name,
          Abbreviation: abbreviation,
          StartingNumber: startingNumber,
        };
        const item = await createSharePointListItem(
          instance,
          SHAREPOINT_SITE_URL,
          COMPANY_LIST_IDENTIFIER,
          fields
        );
        const company = mapCompany(item);
        dispatch({ type: "ADD_COMPANY", payload: company });
        return company;
      } catch (error) {
        handleError(error);
        throw error;
      } finally {
        setLoading("creatingCompany", false);
      }
    },
    [instance, handleError, ensureCompanyColumns, canManageCompanies]
  );

  const updateCompany = useCallback(
    async (companyId, updates = {}) => {
      if (!companyId) {
        throw new Error("Company ID is required.");
      }
      if (!canManageCompanies) {
        throw new Error("You do not have permission to edit companies.");
      }
      const existing =
        state.companies.find((company) => company.id === companyId) || null;
      if (!existing) {
        throw new Error("Company not found.");
      }
      const fields = {};
      if (typeof updates.name !== "undefined") {
        fields.Title = updates.name;
      }
      if (typeof updates.abbreviation !== "undefined") {
        fields.Abbreviation = updates.abbreviation;
      }
      if (typeof updates.startingNumber !== "undefined") {
        fields.StartingNumber = updates.startingNumber;
      }
      if (Object.keys(fields).length === 0) {
        return existing;
      }
      setLoading("updatingCompany", true);
      try {
        await updateSharePointListItem(
          instance,
          SHAREPOINT_SITE_URL,
          COMPANY_LIST_IDENTIFIER,
          companyId,
          fields
        );
        const updatedCompany = {
          ...existing,
          name: fields.Title ?? existing.name,
          abbreviation: fields.Abbreviation ?? existing.abbreviation,
          startingNumber: Number(fields.StartingNumber ?? existing.startingNumber),
        };
        dispatch({
          type: "UPDATE_COMPANY",
          payload: { id: companyId, company: updatedCompany },
        });
        return updatedCompany;
      } catch (error) {
        handleError(error);
        throw error;
      } finally {
        setLoading("updatingCompany", false);
      }
    },
    [state.companies, instance, handleError, canManageCompanies]
  );

  const deleteCompany = useCallback(
    async (companyId) => {
      if (!companyId) return;
      if (!canManageCompanies) {
        throw new Error("You do not have permission to delete companies.");
      }
      setLoading("deletingCompany", true);
      try {
        await deleteSharePointListItem(
          instance,
          SHAREPOINT_SITE_URL,
          COMPANY_LIST_IDENTIFIER,
          companyId
        );
        dispatch({ type: "REMOVE_COMPANY", payload: companyId });
      } catch (error) {
        handleError(error);
        throw error;
      } finally {
        setLoading("deletingCompany", false);
      }
    },
    [instance, handleError, canManageCompanies]
  );

  const createLetter = useCallback(
    async ({
      companyId,
      letterDate,
      recipientCompany,
      subject,
      preparedBy,
      notes,
      year,
      attachments = [],
    }) => {
      if (!companyId) throw new Error("Please choose a company.");
      if (!canEditLetters) {
        throw new Error("You do not have permission to create letters.");
      }
      if (
        allowedCompanySet &&
        !allowedCompanySet.has(String(companyId))
      ) {
        throw new Error("You do not have access to this company.");
      }
      const company = getCompany(companyId);
      if (!company) throw new Error("Selected company was not found.");

      const derivedYear =
        year ||
        new Date(letterDate || new Date().toISOString()).getFullYear();
      const sequenceNumber = getNextSequence(companyId, derivedYear);
      if (!sequenceNumber) {
        throw new Error("Unable to determine the next sequence number.");
      }
      const referenceNumber = buildReferenceNumber(
        company.abbreviation,
        sequenceNumber,
        derivedYear
      );
      const normalizedDate = normalizeLetterDate(letterDate);

      setLoading("creating", true);
      try {
        await ensureLetterColumns();
        const fields = {
          Title: subject || referenceNumber,
          ReferenceNumber: referenceNumber,
          CompanyItemId: company.id,
          CompanyName: company.name,
          CompanyAbbreviation: company.abbreviation,
          SequenceNumber: sequenceNumber,
          Year: derivedYear,
          LetterDate: normalizedDate,
          RecipientCompany: recipientCompany,
          Subject: subject,
          PreparedBy: preparedBy,
          Notes: notes,
        };
        const created = await createSharePointListItem(
          instance,
          SHAREPOINT_SITE_URL,
          LETTER_LIST_IDENTIFIER,
          fields
        );
        let mapped = mapLetter(created);

        if (attachments && attachments.length > 0) {
          const uploaded = await uploadLetterFiles(
            company,
            referenceNumber,
            attachments
          );
          mapped = {
            ...mapped,
            hasAttachments: uploaded.length > 0,
            attachmentsLoaded: true,
            attachments: uploaded,
          };
        } else {
          mapped = {
            ...mapped,
            attachmentsLoaded: true,
          };
        }

        dispatch({ type: "ADD_LETTER", payload: mapped });
        return {
          referenceNumber,
          sequenceNumber,
          year: derivedYear,
        };
      } catch (error) {
        handleError(error);
        throw error;
      } finally {
        setLoading("creating", false);
      }
    },
    [
      instance,
      getCompany,
      getNextSequence,
      handleError,
      ensureLetterColumns,
      canEditLetters,
      allowedCompanySet,
      uploadLetterFiles,
    ]
  );

  const updateLetter = useCallback(
    async ({
      letterId,
      updates = {},
      attachmentsToAdd = [],
      attachmentsToRemove = [],
    }) => {
      if (!letterId) {
        throw new Error("Letter ID is required to update a record.");
      }
      if (!canEditLetters) {
        throw new Error("You do not have permission to edit letters.");
      }
      const existing =
        state.letters.find((letter) => letter.id === letterId) || null;
      if (!existing) {
        throw new Error("Letter not found.");
      }
      if (
        allowedCompanySet &&
        !allowedCompanySet.has(String(existing.companyId))
      ) {
        throw new Error("You do not have access to this letter.");
      }
      const fieldUpdates = {};
      if (typeof updates.subject !== "undefined") {
        fieldUpdates.Subject = updates.subject;
      }
      if (typeof updates.recipientCompany !== "undefined") {
        fieldUpdates.RecipientCompany = updates.recipientCompany;
      }
      if (typeof updates.preparedBy !== "undefined") {
        fieldUpdates.PreparedBy = updates.preparedBy;
      }
      if (typeof updates.notes !== "undefined") {
        fieldUpdates.Notes = updates.notes;
      }
      if (typeof updates.letterDate !== "undefined") {
        fieldUpdates.LetterDate = normalizeLetterDate(
          updates.letterDate || existing.letterDate
        );
      }
      if (typeof updates.year !== "undefined") {
        fieldUpdates.Year = updates.year;
      }
      if (Object.keys(fieldUpdates).length > 0) {
        await updateSharePointListItem(
          instance,
          SHAREPOINT_SITE_URL,
          LETTER_LIST_IDENTIFIER,
          letterId,
          fieldUpdates
        );
      }
      let currentAttachments =
        existing.attachmentsLoaded && existing.attachments
          ? [...existing.attachments]
          : [];
      if (attachmentsToRemove && attachmentsToRemove.length > 0) {
        await deleteLetterFiles(attachmentsToRemove);
        const removeKeys = attachmentsToRemove.map(
          (item) => item.path || item.id || item.name
        );
        currentAttachments = currentAttachments.filter((attachment) => {
          const key = attachment.path || attachment.id || attachment.name;
          return !removeKeys.includes(key);
        });
      }
      if (attachmentsToAdd && attachmentsToAdd.length > 0) {
        const uploaded = await uploadLetterFiles(
          {
            abbreviation: existing.companyAbbreviation,
            name: existing.companyName,
          },
          existing.referenceNumber,
          attachmentsToAdd
        );
        currentAttachments = [...currentAttachments, ...uploaded];
      }
      const updatedLetter = {
        ...existing,
        subject:
          typeof updates.subject !== "undefined"
            ? updates.subject
            : existing.subject,
        recipientCompany:
          typeof updates.recipientCompany !== "undefined"
            ? updates.recipientCompany
            : existing.recipientCompany,
        preparedBy:
          typeof updates.preparedBy !== "undefined"
            ? updates.preparedBy
            : existing.preparedBy,
        notes:
          typeof updates.notes !== "undefined"
            ? updates.notes
            : existing.notes,
        letterDate:
          typeof updates.letterDate !== "undefined"
            ? updates.letterDate
            : existing.letterDate,
        year:
          typeof updates.year !== "undefined" ? updates.year : existing.year,
        attachments: currentAttachments,
        attachmentsLoaded: true,
        hasAttachments: currentAttachments.length > 0,
      };
      dispatch({
        type: "UPDATE_LETTER",
        payload: { id: letterId, letter: updatedLetter },
      });
      return updatedLetter;
    },
    [
      state.letters,
      instance,
      canEditLetters,
      allowedCompanySet,
      uploadLetterFiles,
      deleteLetterFiles,
    ]
  );

  const addUserAccessEntry = useCallback(
    async ({ userPrincipalName, role, companyIds }) => {
      if (!userPrincipalName || !role) {
        throw new Error("User principal name and role are required.");
      }
      if (!isAdmin) {
        throw new Error("Only admins can manage user access entries.");
      }
      const normalizedIds = (companyIds || []).filter(Boolean);
      const companyNames = normalizedIds
        .map((id) => state.companies.find((c) => c.id === id)?.name)
        .filter(Boolean);
      setLoading("creatingUserAccess", true);
      try {
        await ensureUserAccessColumns();
        const fields = {
          Title: userPrincipalName,
          UserPrincipalName: userPrincipalName,
          Role: role,
          CompanyIds: normalizedIds.join(","),
          CompanyNames: companyNames.join(", "),
        };
        const created = await createSharePointListItem(
          instance,
          SHAREPOINT_SITE_URL,
          USER_ACCESS_LIST_IDENTIFIER,
          fields
        );
        const entry = {
          id: created.id,
          title: fields.Title,
          userPrincipalName,
          role,
          companyIds: normalizedIds,
          companyNames,
          raw: fields,
        };
        dispatch({ type: "ADD_USER_ACCESS", payload: entry });
        return entry;
      } catch (error) {
        handleError(error);
        throw error;
      } finally {
        setLoading("creatingUserAccess", false);
      }
    },
    [instance, handleError, ensureUserAccessColumns, state.companies, isAdmin]
  );

  const updateUserAccessEntry = useCallback(
    async (itemId, { userPrincipalName, role, companyIds }) => {
      if (!itemId) {
        throw new Error("User access entry ID is required.");
      }
      if (!isAdmin) {
        throw new Error("Only admins can manage user access entries.");
      }
      const normalizedIds = (companyIds || []).filter(Boolean);
      const companyNames = normalizedIds
        .map((id) => state.companies.find((c) => c.id === id)?.name)
        .filter(Boolean);
      const fields = {};
      if (typeof userPrincipalName !== "undefined") {
        fields.Title = userPrincipalName;
        fields.UserPrincipalName = userPrincipalName;
      }
      if (typeof role !== "undefined") {
        fields.Role = role;
      }
      fields.CompanyIds = normalizedIds.join(",");
      fields.CompanyNames = companyNames.join(", ");
      setLoading("updatingUserAccess", true);
      try {
        await updateSharePointListItem(
          instance,
          SHAREPOINT_SITE_URL,
          USER_ACCESS_LIST_IDENTIFIER,
          itemId,
          fields
        );
        const entry = {
          id: itemId,
          title: fields.Title || userPrincipalName || "",
          userPrincipalName:
            fields.UserPrincipalName || userPrincipalName || "",
          role: fields.Role || role || "Viewer",
          companyIds: normalizedIds,
          companyNames,
          raw: fields,
        };
        dispatch({
          type: "UPDATE_USER_ACCESS",
          payload: { id: itemId, entry },
        });
        return entry;
      } catch (error) {
        handleError(error);
        throw error;
      } finally {
        setLoading("updatingUserAccess", false);
      }
    },
    [instance, handleError, state.companies, isAdmin]
  );

  const deleteUserAccessEntry = useCallback(
    async (itemId) => {
      if (!itemId) return;
      if (!isAdmin) {
        throw new Error("Only admins can remove user access entries.");
      }
      try {
        await deleteSharePointListItem(
          instance,
          SHAREPOINT_SITE_URL,
          USER_ACCESS_LIST_IDENTIFIER,
          itemId
        );
        dispatch({ type: "REMOVE_USER_ACCESS", payload: itemId });
      } catch (error) {
        handleError(error);
        throw error;
      }
    },
    [instance, handleError, isAdmin]
  );

  const loadLetterAttachments = useCallback(
    async (letterId) => {
      if (!letterId) return [];
      const letter = state.letters.find((item) => item.id === letterId);
      if (!letter) return [];
      if (
        allowedCompanySet &&
        !allowedCompanySet.has(String(letter.companyId))
      ) {
        throw new Error("You do not have access to this letter.");
      }
      if (letter.attachmentsLoaded) {
        return letter.attachments || [];
      }
      try {
        const folderPath = buildLetterFolderPath({
          abbreviation: letter.companyAbbreviation,
          name: letter.companyName,
        });
        const filePrefix = sanitizeSegment(
          (letter.referenceNumber || "REF").replace(/[/\\]/g, "-"),
          "REF"
        ).toLowerCase();
        const files = await listDriveFiles(
          instance,
          SHAREPOINT_SITE_URL,
          LETTER_LIBRARY_NAME,
          folderPath
        );
        const matches = [];
        for (const file of files) {
          if (
            !file.name ||
            !file.name.toLowerCase().startsWith(filePrefix)
          ) {
            continue;
          }
          let webUrl = file.webUrl;
          if (!webUrl) {
            webUrl = await getDriveItemViewUrl(
              instance,
              SHAREPOINT_SITE_URL,
              LETTER_LIBRARY_NAME,
              file.path
            );
          }
          matches.push({
            id: file.id || file.path,
            name: file.name,
            path: file.path,
            size: file.size,
            webUrl,
            lastModified: file.lastModified,
          });
        }
        dispatch({
          type: "SET_LETTER_ATTACHMENTS",
          payload: { id: letterId, attachments: matches },
        });
        return matches;
      } catch (error) {
        handleError(error);
        throw error;
      }
    },
    [state.letters, instance, handleError, allowedCompanySet]
  );

  const downloadAttachment = useCallback(
    async (letterId, attachment) => {
      if (!attachment) return null;
      const letter = state.letters.find((item) => item.id === letterId);
      if (
        letter &&
        allowedCompanySet &&
        !allowedCompanySet.has(String(letter.companyId))
      ) {
        throw new Error("You do not have access to this letter.");
      }
      if (attachment.webUrl) return attachment.webUrl;
      if (attachment.path) {
        try {
          return await getDriveItemViewUrl(
            instance,
            SHAREPOINT_SITE_URL,
            LETTER_LIBRARY_NAME,
            attachment.path
          );
        } catch (error) {
          handleError(error);
          throw error;
        }
      }
      return null;
    },
    [instance, handleError, state.letters, allowedCompanySet]
  );

  const searchUsers = useCallback(
    async (term) => {
      try {
        return await searchAzureUsers(instance, term);
      } catch (error) {
        console.warn("[Graph] Failed to search users:", error);
        return [];
      }
    },
    [instance]
  );

  const yearOptions = useMemo(() => {
    const uniqueYears = new Set(
      scopedLetters
        .map((letter) => Number(letter.year))
        .filter((year) => !Number.isNaN(year))
    );
    const sorted = Array.from(uniqueYears).sort((a, b) => b - a);
    const currentYear = new Date().getFullYear();
    if (!sorted.includes(currentYear)) {
      sorted.unshift(currentYear);
    }
    return sorted;
  }, [scopedLetters]);

  const value = {
    isAuthenticated,
    loading: state.loading,
    error: state.error,
    companies: scopedCompanies,
    letters: scopedLetters,
    userAccess: isAdmin ? state.userAccess : [],
    currentUserAccess,
    currentUserPrincipalName,
    userAccessResolved,
    isAdmin,
    canEditLetters,
    canManageCompanies,
    initializing: isAuthenticated && !state.initialized,
    allowedCompanyIds: allowedCompanySet
      ? Array.from(allowedCompanySet)
      : null,
    yearOptions,
    refreshAll,
    loadLetters,
    loadCompanies,
    loadUserAccess,
    addCompany,
    updateCompany,
    deleteCompany,
    createLetter,
    updateLetter,
    getNextSequence,
    getReferencePreview,
    addUserAccessEntry,
    updateUserAccessEntry,
    deleteUserAccessEntry,
    loadLetterAttachments,
    downloadAttachment,
    searchUsers,
  };

  return (
    <LetterContext.Provider value={value}>{children}</LetterContext.Provider>
  );
};

export const useLetters = () => {
  const context = useContext(LetterContext);
  if (!context) {
    throw new Error("useLetters must be used within a LetterProvider");
  }
  return context;
};
