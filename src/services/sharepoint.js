// Microsoft Graph client for SharePoint Drive operations (meeting management)

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

export const defaultScopes = () => {
  const envScopes = (process.env.REACT_APP_GRAPH_SCOPES || "").trim();
  if (envScopes) return envScopes.split(/[\s,]+/).filter(Boolean);
  return ["Sites.ReadWrite.All"];
};

export async function acquireToken(instance, scopes) {
  if (!instance || !instance.getActiveAccount) {
    throw new Error("MSAL instance not initialized");
  }

  const allAccounts = instance.getAllAccounts();
  const activeAccount = instance.getActiveAccount();

  if (allAccounts.length > 0 || activeAccount) {
    const accountToUse = activeAccount || allAccounts[0];

    try {
      const request = {
        scopes: scopes && scopes.length ? scopes : defaultScopes(),
        account: accountToUse,
      };
      const response = await instance.acquireTokenSilent(request);
      return response.accessToken;
    } catch (silentError) {
      console.warn("Silent token acquisition failed:", silentError);
      throw new Error("Authentication required. Please sign in.");
    }
  } else {
    throw new Error("No authenticated user found. Please sign in.");
  }
}

// Generic Graph API GET helper
async function graphGet(url, token) {
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(
      `Graph API error: ${response.status} ${response.statusText} - ${errorText}`
    );
  }

  return response.json();
}

// Get site ID from site URL
async function getSiteId(instance, siteUrl) {
  const token = await acquireToken(instance, ["Sites.Read.All"]);

  // Extract hostname and site path from the URL
  const urlObj = new URL(siteUrl);
  const hostname = urlObj.hostname;
  const pathname = urlObj.pathname;

  // For Microsoft Graph, we need to use the format: sites/{hostname}:{pathname}
  const siteIdentifier = `${hostname}:${pathname}`;
  const graphUrl = `${GRAPH_BASE}/sites/${siteIdentifier}`;

  const data = await graphGet(graphUrl, token);
  return data.id;
}

const looksLikeGuid = (value) =>
  typeof value === "string" && /^[0-9a-fA-F-]{20,}$/.test(value);

// Resolve a drive by display name within a SharePoint site
export async function getDriveIdByName(instance, siteUrl, driveName) {
  const token = await acquireToken(instance, ["Sites.ReadWrite.All"]);
  const siteId = await getSiteId(instance, siteUrl);
  const url = `${GRAPH_BASE}/sites/${siteId}/drives?$select=id,name`;
  const data = await graphGet(url, token);
  const drives = data.value || [];

  const drive = drives.find((d) => d.name === driveName);
  if (!drive) {
    throw new Error(`Drive not found: ${driveName}`);
  }
  return drive.id;
}

// Resolve drive ID (allows direct ID usage)
async function resolveDriveId(instance, siteUrl, driveNameOrId) {
  if (!driveNameOrId) {
    throw new Error("Drive name/ID is required");
  }
  if (looksLikeGuid(driveNameOrId)) {
    return driveNameOrId;
  }
  return getDriveIdByName(instance, siteUrl, driveNameOrId);
}

async function getListIdByName(instance, siteUrl, listName) {
  const token = await acquireToken(instance, ["Sites.ReadWrite.All"]);
  const siteId = await getSiteId(instance, siteUrl);
  const url = `${GRAPH_BASE}/sites/${siteId}/lists?$select=id,displayName`;
  const data = await graphGet(url, token);
  const list =
    (data.value || []).find((item) => item.displayName === listName) || null;
  if (!list) {
    throw new Error(`SharePoint list not found: ${listName}`);
  }
  return list.id;
}

async function resolveListId(instance, siteUrl, listNameOrId) {
  if (!listNameOrId) {
    throw new Error("SharePoint list name/ID is required");
  }
  if (looksLikeGuid(listNameOrId)) {
    return listNameOrId;
  }
  return getListIdByName(instance, siteUrl, listNameOrId);
}

const encodeParams = (paramsObj = {}) => {
  const params = new URLSearchParams();
  Object.entries(paramsObj).forEach(([key, value]) => {
    if (value || value === 0) {
      params.append(key, value);
    }
  });
  const query = params.toString();
  return query ? `?${query}` : "";
};

async function fetchPagedResults(
  url,
  token,
  fetchAll = false,
  extraHeaders = {}
) {
  let nextUrl = url;
  const aggregated = [];

  while (nextUrl) {
    const response = await fetch(nextUrl, {
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ...extraHeaders,
      },
    });
    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(
        `Graph API error: ${response.status} ${response.statusText} - ${errorText}`
      );
    }
    const data = await response.json();
    aggregated.push(...(data.value || []));
    nextUrl = fetchAll ? data["@odata.nextLink"] : null;
  }

  return aggregated;
}


// ---- SharePoint List helpers (letter numbering) ----

export async function listSharePointListItems(
  instance,
  siteUrl,
  listNameOrId,
  options = {}
) {
  if (!siteUrl) {
    throw new Error("SharePoint site URL is not configured");
  }
  const token = await acquireToken(instance, ["Sites.ReadWrite.All"]);
  const siteId = await getSiteId(instance, siteUrl);
  const listId = await resolveListId(instance, siteUrl, listNameOrId);

  const query = encodeParams({
    $expand: options.expand || "fields",
    $select: options.select,
    $orderby: options.orderBy,
    $filter: options.filter,
    $search: options.search,
    $top: options.top,
  });

  const url = `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/items${query}`;
  const needsNonIndexedHeader =
    options.preferHonorNonIndexedQueries ?? true;
  const headers = needsNonIndexedHeader
    ? { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" }
    : {};
  return fetchPagedResults(url, token, options.fetchAll, headers);
}

export async function createSharePointListItem(
  instance,
  siteUrl,
  listNameOrId,
  fields
) {
  if (!siteUrl) {
    throw new Error("SharePoint site URL is not configured");
  }
  const token = await acquireToken(instance, ["Sites.ReadWrite.All"]);
  const siteId = await getSiteId(instance, siteUrl);
  const listId = await resolveListId(instance, siteUrl, listNameOrId);
  const url = `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/items`;

  const response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ fields: fields || {} }),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(
      `Failed to create list item: ${response.status} ${response.statusText} - ${errorText}`
    );
  }

  return response.json();
}

export async function updateSharePointListItem(
  instance,
  siteUrl,
  listNameOrId,
  itemId,
  fields
) {
  const token = await acquireToken(instance, ["Sites.ReadWrite.All"]);
  const siteId = await getSiteId(instance, siteUrl);
  const listId = await resolveListId(instance, siteUrl, listNameOrId);
  const url = `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/items/${itemId}`;

  const response = await fetch(url, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ fields: fields || {} }),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(
      `Failed to update list item: ${response.status} ${response.statusText} - ${errorText}`
    );
  }

  return response.json();
}

export async function deleteSharePointListItem(
  instance,
  siteUrl,
  listNameOrId,
  itemId
) {
  const token = await acquireToken(instance, ["Sites.ReadWrite.All"]);
  const siteId = await getSiteId(instance, siteUrl);
  const listId = await resolveListId(instance, siteUrl, listNameOrId);
  const url = `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/items/${itemId}`;

  const response = await fetch(url, {
    method: "DELETE",
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(
      `Failed to delete list item: ${response.status} ${response.statusText} - ${errorText}`
    );
  }

  return true;
}

export async function ensureDriveFolderPath(
  instance,
  siteUrl,
  driveName,
  fullPath
) {
  const token = await acquireToken(instance, ["Sites.ReadWrite.All"]);
  const driveId = await resolveDriveId(instance, siteUrl, driveName);
  const segments = fullPath.split("/").filter(Boolean);
  let currentPath = "";

  for (const segment of segments) {
    currentPath = currentPath ? `${currentPath}/${segment}` : segment;
    const encoded = encodeURIComponent(currentPath).replace(/%2F/g, "/");
    const url = `${GRAPH_BASE}/drives/${driveId}/root:/${encoded}`;
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (response.ok) continue;
    if (response.status === 404) {
      const parentPath = currentPath.split("/").slice(0, -1).join("/");
      const parentId = parentPath
        ? await getDriveItemId(instance, driveId, parentPath, token)
        : "root";
      await createDriveFolder(instance, driveId, parentId, segment, token);
    } else {
      const errorText = await response.text();
      throw new Error(
        `Failed to inspect folder "${currentPath}": ${response.status} ${response.statusText} - ${errorText}`
      );
    }
  }
}

async function getDriveItemId(instance, driveId, path, token) {
  if (!path || path === "root") return "root";
  const encoded = encodeURIComponent(path).replace(/%2F/g, "/");
  const url = `${GRAPH_BASE}/drives/${driveId}/root:/${encoded}`;
  const data = await graphGet(url, token);
  return data.id;
}

async function createDriveFolder(instance, driveId, parentId, name, token) {
  const url = `${GRAPH_BASE}/drives/${driveId}/items/${parentId}/children`;
  const body = {
    name,
    folder: {},
    "@microsoft.graph.conflictBehavior": "replace",
  };
  const response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(
      `Failed to create folder "${name}": ${response.status} ${response.statusText} - ${errorText}`
    );
  }
  return response.json();
}

export async function uploadBinaryFileToDrive(
  instance,
  siteUrl,
  driveName,
  filePath,
  file,
  contentType
) {
  const token = await acquireToken(instance, ["Sites.ReadWrite.All"]);
  const driveId = await resolveDriveId(instance, siteUrl, driveName);
  const encoded = encodeURIComponent(filePath).replace(/%2F/g, "/");
  const url = `${GRAPH_BASE}/drives/${driveId}/root:/${encoded}:/content`;
  const response = await fetch(url, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": contentType || "application/octet-stream",
    },
    body: file,
  });
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(
      `Failed to upload document "${filePath}": ${response.status} ${response.statusText} - ${errorText}`
    );
  }
  return response.json();
}

export async function listDriveFiles(
  instance,
  siteUrl,
  driveName,
  folderPath
) {
  const token = await acquireToken(instance, ["Sites.ReadWrite.All"]);
  const driveId = await resolveDriveId(instance, siteUrl, driveName);
  const encoded = encodeURIComponent(folderPath).replace(/%2F/g, "/");
  const url = `${GRAPH_BASE}/drives/${driveId}/root:/${encoded}:/children?$select=id,name,size,lastModifiedDateTime,webUrl,file`;
  const data = await graphGet(url, token);
  return (data.value || [])
    .filter((item) => !!item.file)
    .map((item) => ({
      id: item.id,
      name: item.name,
      size: item.size,
      lastModified: item.lastModifiedDateTime,
      path: `${folderPath}/${item.name}`.replace(/\/+/g, "/"),
      webUrl: item.webUrl,
    }));
}

export async function getDriveItemViewUrl(
  instance,
  siteUrl,
  driveName,
  filePath
) {
  const token = await acquireToken(instance, ["Sites.ReadWrite.All"]);
  const driveId = await resolveDriveId(instance, siteUrl, driveName);
  const encoded = encodeURIComponent(filePath).replace(/%2F/g, "/");
  const itemUrl = `${GRAPH_BASE}/drives/${driveId}/root:/${encoded}`;
  const item = await graphGet(itemUrl, token);
  if (item.webUrl) return item.webUrl;

  const createLinkUrl = `${GRAPH_BASE}/drives/${driveId}/items/${item.id}/createLink`;
  const response = await fetch(createLinkUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ type: "view", scope: "organization" }),
  });
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(
      `Failed to create view link: ${response.status} ${response.statusText} - ${errorText}`
    );
  }
  const data = await response.json();
  return data?.link?.webUrl || itemUrl;
}

export async function deleteDriveFile(
  instance,
  siteUrl,
  driveName,
  filePath
) {
  if (!filePath) return;
  const token = await acquireToken(instance, ["Sites.ReadWrite.All"]);
  const driveId = await resolveDriveId(instance, siteUrl, driveName);
  const encoded = encodeURIComponent(filePath).replace(/%2F/g, "/");
  const url = `${GRAPH_BASE}/drives/${driveId}/root:/${encoded}`;
  const response = await fetch(url, {
    method: "DELETE",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
  });
  if (!response.ok && response.status !== 404) {
    const errorText = await response.text();
    throw new Error(
      `Failed to delete file "${filePath}": ${response.status} ${response.statusText} - ${errorText}`
    );
  }
}

export async function searchAzureUsers(
  instance,
  searchTerm,
  options = { top: 10 }
) {
  if (!searchTerm || searchTerm.trim().length < 2) return [];
  const sanitized = searchTerm.trim().replace(/'/g, "''");
  const token = await acquireToken(instance, ["User.Read.All"]);
  const filter = encodeURIComponent(
    `startsWith(mail,'${sanitized}') or startsWith(userPrincipalName,'${sanitized}')`
  );
  const top = Math.max(1, Math.min(options.top || 10, 25));
  const url = `${GRAPH_BASE}/users?$top=${top}&$select=id,displayName,mail,userPrincipalName&$filter=${filter}`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
  });
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(
      `Failed to search Azure AD users: ${response.status} ${response.statusText} - ${errorText}`
    );
  }
  const data = await response.json();
  return (data.value || []).map((user) => ({
    id: user.id,
    displayName: user.displayName || user.userPrincipalName || user.mail || "",
    userPrincipalName: user.userPrincipalName || user.mail || "",
    email: user.mail || user.userPrincipalName || "",
  }));
}

export async function createSharePointList(
  instance,
  siteUrl,
  { displayName, description, columns = [] }
) {
  if (!displayName) {
    throw new Error("List displayName is required for provisioning");
  }
  const token = await acquireToken(instance, ["Sites.ReadWrite.All"]);
  const siteId = await getSiteId(instance, siteUrl);
  const createUrl = `${GRAPH_BASE}/sites/${siteId}/lists`;

  const response = await fetch(createUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      displayName,
      description: description || "",
      list: {
        template: "genericList",
      },
    }),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(
      `Failed to create SharePoint list "${displayName}": ${response.status} ${response.statusText} - ${errorText}`
    );
  }

  const list = await response.json();

  if (columns && columns.length > 0) {
    for (const column of columns) {
      try {
        await fetch(
          `${GRAPH_BASE}/sites/${siteId}/lists/${list.id}/columns`,
          {
            method: "POST",
            headers: {
              Authorization: `Bearer ${token}`,
              "Content-Type": "application/json",
            },
            body: JSON.stringify(column),
          }
        );
      } catch (error) {
        console.warn(
          `[SharePoint] Failed to add column "${column.name}" to list "${displayName}":`,
          error
        );
      }
    }
  }

  return list;
}

export async function ensureSharePointListColumns(
  instance,
  siteUrl,
  listNameOrId,
  columns = []
) {
  if (!columns || columns.length === 0) return;
  const token = await acquireToken(instance, ["Sites.ReadWrite.All"]);
  const siteId = await getSiteId(instance, siteUrl);
  const listId = await resolveListId(instance, siteUrl, listNameOrId);
  const columnsUrl = `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/columns?$select=name`;

  const existingResponse = await fetch(columnsUrl, {
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
  });
  if (!existingResponse.ok) {
    const errorText = await existingResponse.text();
    throw new Error(
      `Failed to read columns for list "${listNameOrId}": ${existingResponse.status} ${existingResponse.statusText} - ${errorText}`
    );
  }
  const existingData = await existingResponse.json();
  const existingNames = new Set(
    (existingData.value || []).map((column) => column.name)
  );

  for (const column of columns) {
    if (!column?.name || existingNames.has(column.name)) continue;
    try {
      await fetch(`${GRAPH_BASE}/sites/${siteId}/lists/${listId}/columns`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(column),
      });
    } catch (error) {
      console.warn(
        `[SharePoint] Failed to ensure column "${column.name}" on list "${listNameOrId}":`,
        error
      );
    }
  }
}
