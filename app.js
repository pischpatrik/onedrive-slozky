import { createOneDriveProvider } from "./onedrive.js";
import { createDemoProvider } from "./providers.js";

const FAVORITES_STORAGE_KEY = "odf_favorites_v1";
const SETTINGS_STORAGE_KEY = "odf_app_settings_v1";
const DEFAULT_CONFIG = {
  allowRuntimeSettings: false,
  initialFavorites: [],
  microsoftClientId: "",
  microsoftTenant: "common",
  mode: "onedrive",
  scopes: ["openid", "profile", "offline_access", "Files.Read", "User.Read"]
};

const FILE_APP_HINTS = {
  docx: "Microsoft Word",
  dwg: "AutoCAD viewer",
  numbers: "Apple Numbers",
  pages: "Apple Pages",
  pdf: "Adobe Acrobat",
  pptx: "Microsoft PowerPoint",
  psd: "Photoshop viewer",
  xlsx: "Microsoft Excel",
  zip: "ZIP opener"
};

const PREVIEWABLE_EXTENSIONS = new Set([
  "csv",
  "doc",
  "docx",
  "gif",
  "jpeg",
  "jpg",
  "pdf",
  "png",
  "ppt",
  "pptx",
  "svg",
  "txt",
  "webp",
  "xls",
  "xlsx"
]);

const baseConfig = {
  ...DEFAULT_CONFIG,
  ...(window.APP_CONFIG || {})
};

const favoriteDialog = document.querySelector("#favorite-dialog");
const settingsDialog = document.querySelector("#settings-dialog");
const elements = {
  accountBadge: document.querySelector("#account-badge"),
  addFavoriteButton: document.querySelector("#add-favorite-button"),
  authButton: document.querySelector("#auth-button"),
  breadcrumbs: document.querySelector("#breadcrumbs"),
  cancelDialogButton: document.querySelector("#cancel-dialog-button"),
  closeDialogButton: document.querySelector("#close-dialog-button"),
  closeSettingsButton: document.querySelector("#close-settings-button"),
  closeSettingsTopButton: document.querySelector("#close-settings-top-button"),
  copyRedirectButton: document.querySelector("#copy-redirect-button"),
  currentFolderName: document.querySelector("#current-folder-name"),
  favoriteForm: document.querySelector("#favorite-form"),
  favoriteLabelInput: document.querySelector("#favorite-label-input"),
  favoriteNote: document.querySelector("#favorite-note"),
  favoritePathInput: document.querySelector("#favorite-path-input"),
  favoritesGrid: document.querySelector("#favorites-grid"),
  fileList: document.querySelector("#file-list"),
  goUpButton: document.querySelector("#go-up-button"),
  installButton: document.querySelector("#install-button"),
  installHint: document.querySelector("#install-hint"),
  modeBadge: document.querySelector("#mode-badge"),
  redirectUriOutput: document.querySelector("#redirect-uri-output"),
  resetSettingsButton: document.querySelector("#reset-settings-button"),
  settingsButton: document.querySelector("#settings-button"),
  settingsClientIdInput: document.querySelector("#settings-client-id-input"),
  settingsForm: document.querySelector("#settings-form"),
  settingsModeSelect: document.querySelector("#settings-mode-select"),
  settingsTenantInput: document.querySelector("#settings-tenant-input"),
  statusCard: document.querySelector("#status-card"),
  statusText: document.querySelector("#status-text")
};

const state = {
  currentPath: "/",
  favorites: [],
  installPrompt: null,
  items: [],
  localSettings: {},
  provider: null,
  runtimeConfig: null,
  session: {
    configured: false,
    displayMode: "OneDrive",
    isAuthenticated: false,
    userName: "Chybi nastaveni"
  }
};

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function normalizePath(path) {
  if (!path || path === "/") {
    return "/";
  }

  return `/${path.split("/").filter(Boolean).join("/")}`;
}

function parentPath(path) {
  const safePath = normalizePath(path);
  if (safePath === "/") {
    return "/";
  }

  const parts = safePath.split("/").filter(Boolean);
  parts.pop();
  return parts.length ? `/${parts.join("/")}` : "/";
}

function pathLabel(path) {
  const safePath = normalizePath(path);
  if (safePath === "/") {
    return "Root";
  }

  return safePath.split("/").filter(Boolean).at(-1) || "Root";
}

function formatSize(size) {
  if (!Number.isFinite(size) || size <= 0) {
    return "0 B";
  }

  const units = ["B", "KB", "MB", "GB"];
  let value = size;
  let unitIndex = 0;

  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }

  const digits = value >= 10 || unitIndex === 0 ? 0 : 1;
  return `${value.toFixed(digits)} ${units[unitIndex]}`;
}

function formatDate(dateString) {
  if (!dateString) {
    return "";
  }

  try {
    return new Intl.DateTimeFormat("cs-CZ", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric"
    }).format(new Date(dateString));
  } catch (error) {
    return "";
  }
}

function getPlatform() {
  const userAgent = navigator.userAgent.toLowerCase();
  if (/iphone|ipad|ipod/.test(userAgent)) {
    return "ios";
  }
  if (/android/.test(userAgent)) {
    return "android";
  }
  return "other";
}

function installHintText() {
  const platform = getPlatform();
  if (platform === "ios") {
    return "V Safari zvol Sdilet -> Na plochu";
  }
  if (platform === "android") {
    return "V Chrome zvol Instalovat aplikaci";
  }
  return "Lze nainstalovat z prohlizece";
}

function currentRedirectUri() {
  return `${window.location.origin}${window.location.pathname}`;
}

function readLocalSettings() {
  const raw = localStorage.getItem(SETTINGS_STORAGE_KEY);
  if (!raw) {
    return {};
  }

  try {
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") {
      return {};
    }

    return {
      microsoftClientId: String(parsed.microsoftClientId || "").trim(),
      microsoftTenant: String(parsed.microsoftTenant || "").trim() || "common",
      mode: parsed.mode === "demo" ? "demo" : "onedrive"
    };
  } catch (error) {
    return {};
  }
}

function saveLocalSettings() {
  localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(state.localSettings));
}

function clearLocalSettings() {
  localStorage.removeItem(SETTINGS_STORAGE_KEY);
}

function buildRuntimeConfig() {
  return {
    ...baseConfig,
    ...state.localSettings,
    scopes: [...baseConfig.scopes]
  };
}

function normalizeFavoriteEntry(entry) {
  if (!entry) {
    return null;
  }

  if (typeof entry === "string") {
    const path = normalizePath(entry);
    return {
      label: pathLabel(path),
      path
    };
  }

  if (!entry.path) {
    return null;
  }

  const path = normalizePath(String(entry.path).trim());
  return {
    label: String(entry.label || pathLabel(path)).trim() || pathLabel(path),
    path
  };
}

function readFavorites() {
  const fallbackFavorites = baseConfig.initialFavorites
    .map((entry) => normalizeFavoriteEntry(entry))
    .filter(Boolean);

  const raw = localStorage.getItem(FAVORITES_STORAGE_KEY);
  if (!raw) {
    return fallbackFavorites;
  }

  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return fallbackFavorites;
    }

    return parsed
      .map((entry) => normalizeFavoriteEntry(entry))
      .filter(Boolean);
  } catch (error) {
    return fallbackFavorites;
  }
}

function saveFavorites() {
  localStorage.setItem(FAVORITES_STORAGE_KEY, JSON.stringify(state.favorites));
}

function setStatus(message, options = {}) {
  const { isError = false, showList = false } = options;
  elements.statusCard.hidden = false;
  elements.fileList.hidden = !showList;
  elements.statusText.textContent = message;
  elements.statusCard.style.borderStyle = isError ? "solid" : "dashed";
  elements.statusCard.style.borderColor = isError
    ? "rgba(174, 61, 61, 0.32)"
    : "rgba(24, 95, 88, 0.22)";
}

function chooseProvider() {
  if (state.runtimeConfig.mode === "onedrive") {
    state.provider = createOneDriveProvider(state.runtimeConfig);
    return;
  }

  state.provider = createDemoProvider();
}

function renderFavorites() {
  if (!state.favorites.length) {
    elements.favoritesGrid.innerHTML = `
      <div class="status-card">
        <p>Nemame zadnou oblibenou slozku. Klepni na plus a jednu pridej.</p>
      </div>
    `;
    return;
  }

  elements.favoritesGrid.innerHTML = state.favorites
    .map((favorite) => {
      const favoritePath = normalizePath(favorite.path);
      const currentPath = normalizePath(state.currentPath);
      const isActive =
        currentPath === favoritePath ||
        (favoritePath !== "/" && currentPath.startsWith(`${favoritePath}/`));
      return `
        <article class="favorite-card ${isActive ? "favorite-card-active" : ""}">
          <button class="favorite-remove" type="button" data-action="remove-favorite" data-path="${escapeHtml(favorite.path)}">
            -
          </button>
          <button class="favorite-open" type="button" data-action="open-favorite" data-path="${escapeHtml(favorite.path)}">
            <span class="favorite-badge">Slozka</span>
            <span class="favorite-name">${escapeHtml(favorite.label)}</span>
            <span class="favorite-path">${escapeHtml(favorite.path)}</span>
          </button>
        </article>
      `;
    })
    .join("");
}

function renderBreadcrumbs() {
  const parts = normalizePath(state.currentPath).split("/").filter(Boolean);
  const crumbs = [{ label: "Root", path: "/" }];
  let running = "";

  for (const part of parts) {
    running += `/${part}`;
    crumbs.push({ label: part, path: running });
  }

  elements.breadcrumbs.innerHTML = crumbs
    .map(
      (crumb, index) => `
        <button
          class="crumb-button"
          type="button"
          data-action="open-folder"
          data-path="${escapeHtml(crumb.path)}"
        >
          ${escapeHtml(crumb.label)}
        </button>
        ${index < crumbs.length - 1 ? '<span aria-hidden="true">/</span>' : ""}
      `
    )
    .join("");
}

function fileIconLabel(item) {
  if (item.type === "folder") {
    return "DIR";
  }

  return (item.extension || "soubor").slice(0, 4);
}

function renderItems() {
  if (!state.items.length) {
    elements.fileList.innerHTML = "";
    setStatus("Slozka je prazdna.", { showList: false });
    return;
  }

  elements.fileList.innerHTML = state.items
    .map((item) => {
      const metaParts = [];
      if (item.type === "folder") {
        metaParts.push(`${item.childCount ?? 0} polozek`);
      } else {
        metaParts.push(formatSize(item.size));
      }

      const formattedDate = formatDate(item.modifiedAt);
      if (formattedDate) {
        metaParts.push(formattedDate);
      }

      const primaryLabel = item.type === "folder" ? "Otevrit" : "Otevrit soubor";
      const showAppFinder = item.type === "file" && !PREVIEWABLE_EXTENSIONS.has((item.extension || "").toLowerCase());

      return `
        <article class="file-card">
          <div class="file-main">
            <span class="file-icon">${escapeHtml(fileIconLabel(item))}</span>
            <div>
              <div class="file-name">${escapeHtml(item.name)}</div>
              <div class="file-meta">
                ${metaParts.map((part) => `<span>${escapeHtml(part)}</span>`).join("")}
              </div>
            </div>
          </div>
          <div class="file-actions">
            <button
              class="ghost-button"
              type="button"
              data-action="${item.type === "folder" ? "open-folder" : "open-file"}"
              data-path="${escapeHtml(item.path)}"
            >
              ${primaryLabel}
            </button>
            ${
              showAppFinder
                ? `
                  <button
                    class="secondary-button"
                    type="button"
                    data-action="find-app"
                    data-name="${escapeHtml(item.name)}"
                    data-extension="${escapeHtml(item.extension || "")}"
                  >
                    Najit appku
                  </button>
                `
                : ""
            }
          </div>
        </article>
      `;
    })
    .join("");

  setStatus("Obsah nacten.", { showList: true });
}

function syncSettingsDialog() {
  elements.settingsModeSelect.value = state.runtimeConfig.mode;
  elements.settingsClientIdInput.value = state.runtimeConfig.microsoftClientId || "";
  elements.settingsTenantInput.value = state.runtimeConfig.microsoftTenant || "common";
  elements.redirectUriOutput.textContent = currentRedirectUri();
}

function shouldShowSettingsButton() {
  if (state.runtimeConfig.allowRuntimeSettings) {
    return true;
  }

  if (!state.runtimeConfig.microsoftClientId) {
    return true;
  }

  return false;
}

function renderSession() {
  elements.modeBadge.textContent = state.session.displayMode;
  elements.accountBadge.textContent = state.session.userName;
  elements.settingsButton.hidden = !shouldShowSettingsButton();

  if (state.session.displayMode === "Demo") {
    elements.authButton.textContent = "Nastavit OneDrive";
    elements.favoriteNote.textContent =
      "Ted jedes v demo rezimu. V nastaveni prepnes aplikaci na skutecny OneDrive a vlozis client ID.";
    return;
  }

  if (!state.session.configured) {
    elements.authButton.textContent = "Doplnit OneDrive";
    elements.favoriteNote.textContent =
      "OneDrive rezim je pripraveny, ale chybi Microsoft client ID. Otevri Nastaveni a vloz ho.";
    return;
  }

  elements.authButton.textContent = state.session.isAuthenticated ? "Odhlasit" : "Prihlasit k OneDrive";
  elements.favoriteNote.textContent =
    "Oblibene slozky jsou ulozene lokalne v telefonu. Obsah slozek se cte z tveho OneDrive.";
}

async function refreshSession() {
  state.session = await state.provider.getSession();
  renderSession();
}

function setupPromptText() {
  if (state.runtimeConfig.mode === "demo") {
    return "Jses v demo rezimu. Otevri Nastaveni a prepni aplikaci na OneDrive.";
  }

  if (!state.session.configured) {
    return "Nejdriv otevri Nastaveni, vloz Microsoft client ID a uloz ho.";
  }

  return "Klikni na Prihlasit k OneDrive a pak se vrat sem.";
}

async function browse(path) {
  const safePath = normalizePath(path);
  state.currentPath = safePath;
  elements.currentFolderName.textContent = pathLabel(safePath);
  elements.goUpButton.hidden = safePath === "/";
  renderFavorites();
  renderBreadcrumbs();

  if (state.runtimeConfig.mode === "onedrive" && !state.session.isAuthenticated) {
    state.items = [];
    elements.fileList.innerHTML = "";
    setStatus(setupPromptText(), { showList: false });
    return;
  }

  setStatus("Nacitam obsah slozky...");

  try {
    const items = await state.provider.listFolder(safePath);
    state.items = items;
    elements.currentFolderName.textContent = pathLabel(state.currentPath);
    renderItems();
  } catch (error) {
    state.items = [];
    elements.fileList.innerHTML = "";
    setStatus(error.message || "Obsah slozky se nepodarilo nacist.", {
      isError: true,
      showList: false
    });
  }
}

function openFavoriteDialog() {
  elements.favoriteLabelInput.value = "";
  elements.favoritePathInput.value = state.currentPath || "/";
  favoriteDialog.showModal();
  elements.favoriteLabelInput.focus();
}

function closeFavoriteDialog() {
  favoriteDialog.close();
}

function openSettingsDialog() {
  syncSettingsDialog();
  settingsDialog.showModal();
}

function closeSettingsDialog() {
  settingsDialog.close();
}

function findItemByPath(path) {
  return state.items.find((item) => normalizePath(item.path) === normalizePath(path));
}

function storeSearchUrl(query) {
  const encoded = encodeURIComponent(query);
  const platform = getPlatform();
  if (platform === "ios") {
    return `https://apps.apple.com/us/search?term=${encoded}`;
  }
  if (platform === "android") {
    return `https://play.google.com/store/search?q=${encoded}&c=apps`;
  }
  return `https://www.google.com/search?q=${encoded}`;
}

function openAppSearch(itemName, extension) {
  const cleanExtension = (extension || itemName.split(".").pop() || "").toLowerCase();
  const hint = FILE_APP_HINTS[cleanExtension] || `${cleanExtension} file opener`;
  const question =
    cleanExtension
      ? `Soubor .${cleanExtension} nema jistou podporu v prohlizeci. Chces vyhledat appku "${hint}"?`
      : "Soubor nema jistou podporu v prohlizeci. Chces vyhledat appku pro otevreni?";

  if (!window.confirm(question)) {
    return;
  }

  window.open(storeSearchUrl(hint), "_blank", "noopener");
}

async function handleFileOpen(path) {
  const item = findItemByPath(path);
  if (!item) {
    setStatus("Soubor uz v seznamu neni. Zkus slozku znovu nacist.", { isError: true });
    return;
  }

  const isPreviewable = PREVIEWABLE_EXTENSIONS.has((item.extension || "").toLowerCase());
  if (!isPreviewable) {
    openAppSearch(item.name, item.extension || "");
    return;
  }

  try {
    const opened = await state.provider.openFile(item);
    if (!opened) {
      openAppSearch(item.name, item.extension || "");
    }
  } catch (error) {
    setStatus(error.message || "Soubor se nepodarilo otevrit.", { isError: true });
  }
}

async function applyRuntimeConfig(options = {}) {
  const { preservePath = true } = options;
  state.runtimeConfig = buildRuntimeConfig();
  chooseProvider();
  syncSettingsDialog();

  try {
    await state.provider.initialize();
  } catch (error) {
    setStatus(error.message || "Aplikace narazila na problem pri startu.", { isError: true });
  }

  await refreshSession();
  renderFavorites();
  const targetPath = preservePath && state.currentPath !== "/" ? state.currentPath : state.favorites[0]?.path || "/";
  await browse(targetPath);
}

async function handleAuth() {
  if (state.session.displayMode === "Demo") {
    openSettingsDialog();
    setStatus("V nastaveni prepni aplikaci na OneDrive a uloz zmeny.", {
      showList: state.items.length > 0
    });
    return;
  }

  if (!state.session.configured) {
    openSettingsDialog();
    setStatus("Dopln Microsoft client ID a tenant, pak znovu klikni na Prihlasit.", {
      isError: true,
      showList: state.items.length > 0
    });
    return;
  }

  try {
    if (state.session.isAuthenticated) {
      await state.provider.signOut();
      return;
    }

    await state.provider.signIn();
  } catch (error) {
    setStatus(error.message || "Prihlaseni se nepodarilo spustit.", {
      isError: true,
      showList: state.items.length > 0
    });
  }
}

async function saveSettingsFromDialog() {
  const nextSettings = {
    microsoftClientId: elements.settingsClientIdInput.value.trim(),
    microsoftTenant: elements.settingsTenantInput.value.trim() || "common",
    mode: elements.settingsModeSelect.value === "demo" ? "demo" : "onedrive"
  };

  const changed =
    nextSettings.microsoftClientId !== state.runtimeConfig.microsoftClientId ||
    nextSettings.microsoftTenant !== state.runtimeConfig.microsoftTenant ||
    nextSettings.mode !== state.runtimeConfig.mode;

  state.localSettings = nextSettings;
  saveLocalSettings();

  if (changed && state.provider?.clearSession) {
    await state.provider.clearSession();
  }

  closeSettingsDialog();
  await applyRuntimeConfig({ preservePath: false });

  if (nextSettings.mode === "demo") {
    setStatus("Demo rezim je ulozeny. Kdykoli ho muzes v nastaveni prepnout na OneDrive.", {
      showList: state.items.length > 0
    });
    return;
  }

  if (!nextSettings.microsoftClientId) {
    setStatus("Nastaveni je ulozene. Ted jeste vloz Microsoft client ID.", {
      isError: true,
      showList: state.items.length > 0
    });
    return;
  }

  setStatus("Nastaveni je ulozene. Ted klikni na Prihlasit k OneDrive.", {
    showList: state.items.length > 0
  });
}

async function resetSettings() {
  if (!window.confirm("Opravdu chces smazat lokalni nastaveni OneDrive v tomhle prohlizeci?")) {
    return;
  }

  clearLocalSettings();
  state.localSettings = {};
  if (state.provider?.clearSession) {
    await state.provider.clearSession();
  }

  closeSettingsDialog();
  await applyRuntimeConfig({ preservePath: false });
  setStatus("Lokalni nastaveni bylo smazano.", {
    showList: state.items.length > 0
  });
}

async function copyRedirectUri() {
  const value = currentRedirectUri();
  try {
    await navigator.clipboard.writeText(value);
    setStatus("Redirect URI jsem zkopiroval do schranky.", {
      showList: state.items.length > 0
    });
  } catch (error) {
    setStatus(`Redirect URI: ${value}`, {
      showList: state.items.length > 0
    });
  }
}

function registerInstallPrompt() {
  window.addEventListener("beforeinstallprompt", (event) => {
    event.preventDefault();
    state.installPrompt = event;
    elements.installButton.hidden = false;
  });

  window.addEventListener("appinstalled", () => {
    state.installPrompt = null;
    elements.installButton.hidden = true;
    elements.installHint.textContent = "Aplikace je nainstalovana";
  });
}

function bindEvents() {
  elements.installHint.textContent = installHintText();

  elements.installButton.addEventListener("click", async () => {
    if (!state.installPrompt) {
      setStatus("Instalace je dostupna pres menu prohlizece.", { showList: state.items.length > 0 });
      return;
    }

    await state.installPrompt.prompt();
    state.installPrompt = null;
    elements.installButton.hidden = true;
  });

  elements.authButton.addEventListener("click", () => {
    handleAuth();
  });

  elements.settingsButton.addEventListener("click", openSettingsDialog);
  elements.copyRedirectButton.addEventListener("click", copyRedirectUri);
  elements.closeSettingsButton.addEventListener("click", closeSettingsDialog);
  elements.closeSettingsTopButton.addEventListener("click", closeSettingsDialog);
  elements.resetSettingsButton.addEventListener("click", () => {
    resetSettings();
  });

  elements.addFavoriteButton.addEventListener("click", openFavoriteDialog);
  elements.cancelDialogButton.addEventListener("click", closeFavoriteDialog);
  elements.closeDialogButton.addEventListener("click", closeFavoriteDialog);
  elements.goUpButton.addEventListener("click", () => browse(parentPath(state.currentPath)));

  elements.settingsForm.addEventListener("submit", (event) => {
    event.preventDefault();
    saveSettingsFromDialog();
  });

  elements.favoriteForm.addEventListener("submit", async (event) => {
    event.preventDefault();

    const label = elements.favoriteLabelInput.value.trim();
    const path = normalizePath(elements.favoritePathInput.value.trim());
    if (!label || !path) {
      return;
    }

    const shouldVerify = state.runtimeConfig.mode === "demo" || state.session.isAuthenticated;
    if (shouldVerify) {
      try {
        await state.provider.listFolder(path);
      } catch (error) {
        setStatus(error.message || "Slozku se nepodarilo overit.", {
          isError: true,
          showList: state.items.length > 0
        });
        return;
      }
    }

    const exists = state.favorites.some((favorite) => normalizePath(favorite.path) === path);
    if (!exists) {
      state.favorites.push({ label, path });
      saveFavorites();
      renderFavorites();
    }

    closeFavoriteDialog();
    await browse(path);
  });

  elements.favoritesGrid.addEventListener("click", async (event) => {
    const button = event.target.closest("button[data-action]");
    if (!button) {
      return;
    }

    const action = button.dataset.action;
    const path = button.dataset.path || "/";

    if (action === "open-favorite") {
      await browse(path);
      return;
    }

    if (action === "remove-favorite") {
      state.favorites = state.favorites.filter((favorite) => normalizePath(favorite.path) !== normalizePath(path));
      saveFavorites();
      renderFavorites();
      if (normalizePath(state.currentPath) === normalizePath(path)) {
        await browse("/");
      }
    }
  });

  elements.breadcrumbs.addEventListener("click", async (event) => {
    const button = event.target.closest("button[data-action='open-folder']");
    if (!button) {
      return;
    }

    await browse(button.dataset.path || "/");
  });

  elements.fileList.addEventListener("click", async (event) => {
    const button = event.target.closest("button[data-action]");
    if (!button) {
      return;
    }

    const { action } = button.dataset;
    if (action === "open-folder") {
      await browse(button.dataset.path || "/");
      return;
    }

    if (action === "open-file") {
      await handleFileOpen(button.dataset.path || "/");
      return;
    }

    if (action === "find-app") {
      openAppSearch(button.dataset.name || "", button.dataset.extension || "");
    }
  });
}

async function registerServiceWorker() {
  if (!("serviceWorker" in navigator)) {
    return;
  }

  try {
    await navigator.serviceWorker.register("./sw.js");
  } catch (error) {
    console.error("Service worker registration failed", error);
  }
}

async function init() {
  state.localSettings = readLocalSettings();
  state.favorites = readFavorites();
  state.currentPath = state.favorites[0]?.path || "/";
  bindEvents();
  registerInstallPrompt();
  await registerServiceWorker();
  await applyRuntimeConfig({ preservePath: false });
}

init();
