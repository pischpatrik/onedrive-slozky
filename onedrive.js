const TOKEN_STORAGE_KEY = "odf_auth_token_v1";
const PKCE_STORAGE_KEY = "odf_auth_pkce_v1";
const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";

function normalizePath(path) {
  if (!path || path === "/") {
    return "/";
  }

  return `/${path.split("/").filter(Boolean).join("/")}`;
}

function authorityBase(tenant) {
  return `https://login.microsoftonline.com/${tenant || "common"}/oauth2/v2.0`;
}

function createRedirectUri() {
  const { origin, pathname } = window.location;
  return `${origin}${pathname}`;
}

function readToken() {
  const raw = localStorage.getItem(TOKEN_STORAGE_KEY);
  if (!raw) {
    return null;
  }

  try {
    return JSON.parse(raw);
  } catch (error) {
    localStorage.removeItem(TOKEN_STORAGE_KEY);
    return null;
  }
}

function writeToken(token) {
  localStorage.setItem(TOKEN_STORAGE_KEY, JSON.stringify(token));
}

function clearToken() {
  localStorage.removeItem(TOKEN_STORAGE_KEY);
}

function readPkce() {
  const raw = sessionStorage.getItem(PKCE_STORAGE_KEY);
  if (!raw) {
    return null;
  }

  try {
    return JSON.parse(raw);
  } catch (error) {
    sessionStorage.removeItem(PKCE_STORAGE_KEY);
    return null;
  }
}

function writePkce(value) {
  sessionStorage.setItem(PKCE_STORAGE_KEY, JSON.stringify(value));
}

function clearPkce() {
  sessionStorage.removeItem(PKCE_STORAGE_KEY);
}

function humanizeGraphError(message) {
  const text = String(message || "").trim();

  if (text.includes("Tenant does not have a SPO license")) {
    return "Jsi prihlaseny uctem, ktery nema OneDrive nebo SharePoint licenci. Pro osobni OneDrive se prihlas osobnim Microsoft uctem (Outlook/Hotmail/Live). Pro firemni ucet musi mit tenant aktivni OneDrive/SharePoint.";
  }

  if (text.includes("Access denied")) {
    return "Microsoft zamitl pristup. Zkus se prihlasit znovu a potvrdit opravneni k OneDrive.";
  }

  if (text.includes("itemNotFound")) {
    return "Tahle slozka v OneDrive nebyla nalezena.";
  }

  return text || "OneDrive vratil chybu.";
}

function randomString(length = 64) {
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~";
  const values = crypto.getRandomValues(new Uint8Array(length));
  return Array.from(values, (value) => alphabet[value % alphabet.length]).join("");
}

function bytesToBase64Url(bytes) {
  let text = "";
  for (const byte of bytes) {
    text += String.fromCharCode(byte);
  }

  return btoa(text).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

async function createCodeChallenge(verifier) {
  const buffer = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier));
  return bytesToBase64Url(new Uint8Array(buffer));
}

async function requestToken(config, params) {
  const response = await fetch(`${authorityBase(config.microsoftTenant)}/token`, {
    body: params,
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    method: "POST"
  });

  const payload = await response.json();
  if (!response.ok) {
    const message = humanizeGraphError(payload.error_description || payload.error || "Prihlaseni se nezdarilo.");
    throw new Error(message);
  }

  const token = {
    accessToken: payload.access_token,
    expiresAt: Date.now() + (Number(payload.expires_in) * 1000),
    refreshToken: payload.refresh_token || null
  };

  writeToken(token);
  return token;
}

async function exchangeCodeForToken(config, code, codeVerifier) {
  const params = new URLSearchParams({
    client_id: config.microsoftClientId,
    code,
    code_verifier: codeVerifier,
    grant_type: "authorization_code",
    redirect_uri: createRedirectUri(),
    scope: config.scopes.join(" ")
  });

  return requestToken(config, params);
}

async function refreshToken(config, currentToken) {
  if (!currentToken?.refreshToken) {
    throw new Error("Relace vyprsela. Prihlas se znovu.");
  }

  const params = new URLSearchParams({
    client_id: config.microsoftClientId,
    grant_type: "refresh_token",
    redirect_uri: createRedirectUri(),
    refresh_token: currentToken.refreshToken,
    scope: config.scopes.join(" ")
  });

  return requestToken(config, params);
}

async function ensureValidToken(config) {
  const token = readToken();
  if (!token?.accessToken) {
    throw new Error("Nejprve se prihlas k OneDrive.");
  }

  if (token.expiresAt && token.expiresAt > Date.now() + 60000) {
    return token.accessToken;
  }

  const nextToken = await refreshToken(config, token);
  return nextToken.accessToken;
}

async function graphFetch(config, pathOrUrl) {
  const accessToken = await ensureValidToken(config);
  const url = pathOrUrl.startsWith("http") ? pathOrUrl : `${GRAPH_ROOT}${pathOrUrl}`;

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });

  if (response.status === 401) {
    clearToken();
    throw new Error("Prihlaseni vyprselo. Prihlas se znovu.");
  }

  const payload = await response.json();
  if (!response.ok) {
    const message = humanizeGraphError(payload.error?.message || "OneDrive vratil chybu.");
    throw new Error(message);
  }

  return payload;
}

function graphPath(path) {
  const safePath = normalizePath(path);
  if (safePath === "/") {
    return "/me/drive/root/children?$top=200";
  }

  const encoded = safePath
    .split("/")
    .filter(Boolean)
    .map((segment) => encodeURIComponent(segment))
    .join("/");

  return `/me/drive/root:/${encoded}:/children?$top=200`;
}

async function fetchAllChildren(config, path) {
  let url = graphPath(path);
  const items = [];

  while (url) {
    const payload = await graphFetch(config, url);
    items.push(...(payload.value || []));
    url = payload["@odata.nextLink"] || null;
  }

  return items;
}

function mapDriveItem(parentPath, item) {
  const base = {
    id: item.id,
    modifiedAt: item.lastModifiedDateTime || null,
    name: item.name,
    path: parentPath === "/" ? `/${item.name}` : `${parentPath}/${item.name}`,
    size: item.size ?? 0
  };

  if (item.folder) {
    return {
      ...base,
      childCount: item.folder.childCount ?? 0,
      type: "folder"
    };
  }

  const extension = item.file?.mimeType
    ? item.name.split(".").pop()?.toLowerCase() || ""
    : item.name.split(".").pop()?.toLowerCase() || "";

  return {
    ...base,
    downloadUrl: item["@microsoft.graph.downloadUrl"] || null,
    extension,
    type: "file",
    webUrl: item.webUrl || null
  };
}

async function fetchProfile(config) {
  const profile = await graphFetch(config, "/me?$select=displayName,userPrincipalName");
  return profile.displayName || profile.userPrincipalName || "Microsoft ucet";
}

async function handleAuthRedirect(config) {
  const url = new URL(window.location.href);
  const code = url.searchParams.get("code");
  const error = url.searchParams.get("error");

  if (error) {
    const description = url.searchParams.get("error_description") || error;
    history.replaceState({}, document.title, createRedirectUri());
    throw new Error(description);
  }

  if (!code) {
    return null;
  }

  const state = url.searchParams.get("state");
  const stored = readPkce();
  if (!stored || stored.state !== state) {
    history.replaceState({}, document.title, createRedirectUri());
    clearPkce();
    throw new Error("Prihlaseni se nepodarilo overit. Zkus to znovu.");
  }

  await exchangeCodeForToken(config, code, stored.codeVerifier);
  clearPkce();
  history.replaceState({}, document.title, createRedirectUri());
  return true;
}

export function createOneDriveProvider(config) {
  return {
    async initialize() {
      await handleAuthRedirect(config);
    },

    async getSession() {
      const configured = Boolean(config.microsoftClientId);
      if (!configured) {
        return {
          configured: false,
          displayMode: "OneDrive",
          isAuthenticated: false,
          userName: "Chybi clientId"
        };
      }

      const token = readToken();
      if (!token?.accessToken) {
        return {
          configured: true,
          displayMode: "OneDrive",
          isAuthenticated: false,
          userName: "Neprihlasen"
        };
      }

      try {
        const userName = await fetchProfile(config);
        return {
          configured: true,
          displayMode: "OneDrive",
          isAuthenticated: true,
          userName
        };
      } catch (error) {
        clearToken();
        return {
          configured: true,
          displayMode: "OneDrive",
          isAuthenticated: false,
          userName: "Relace vyprsela"
        };
      }
    },

    async listFolder(path) {
      const items = await fetchAllChildren(config, path);
      return items
        .map((item) => mapDriveItem(normalizePath(path), item))
        .sort((left, right) => {
          if (left.type !== right.type) {
            return left.type === "folder" ? -1 : 1;
          }
          return left.name.localeCompare(right.name, "cs");
        });
    },

    async openFile(item) {
      const url = item.webUrl || item.downloadUrl;
      if (!url) {
        return false;
      }

      window.open(url, "_blank", "noopener");
      return true;
    },

    async signIn() {
      if (!config.microsoftClientId) {
        throw new Error("Nejdriv dopln clientId do config.js.");
      }

      const verifier = randomString(96);
      const state = randomString(24);
      const challenge = await createCodeChallenge(verifier);
      const redirectUri = createRedirectUri();

      writePkce({
        codeVerifier: verifier,
        state
      });

      const params = new URLSearchParams({
        client_id: config.microsoftClientId,
        code_challenge: challenge,
        code_challenge_method: "S256",
        prompt: "select_account",
        redirect_uri: redirectUri,
        response_mode: "query",
        response_type: "code",
        scope: config.scopes.join(" "),
        state
      });

      window.location.assign(`${authorityBase(config.microsoftTenant)}/authorize?${params.toString()}`);
      return true;
    },

    async signOut() {
      clearToken();
      clearPkce();
      window.location.assign(
        `${authorityBase(config.microsoftTenant)}/logout?post_logout_redirect_uri=${encodeURIComponent(createRedirectUri())}`
      );
      return true;
    },

    async clearSession() {
      clearToken();
      clearPkce();
      return true;
    }
  };
}
