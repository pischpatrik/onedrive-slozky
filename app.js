import { createOneDriveProvider } from "./onedrive.js";

const FAVORITES_STORAGE_KEY = "odf_favorites_v1";
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

const FILE_APP_HINTS = {
  csv: "Microsoft Excel",
  doc: "Microsoft Word",
  docx: "Microsoft Word",
  dwg: "AutoCAD viewer",
  numbers: "Apple Numbers",
  pages: "Apple Pages",
  pdf: "Adobe Acrobat",
  ppt: "Microsoft PowerPoint",
  pptx: "Microsoft PowerPoint",
  psd: "Photoshop viewer",
  txt: "Text editor",
  xls: "Microsoft Excel",
  xlsx: "Microsoft Excel",
  zip: "ZIP opener"
};

const IMAGE_EXTENSIONS = new Set(["gif", "jpeg", "jpg", "png", "svg", "webp"]);
const FILE_ICON_ASSETS = {
  archive: assetUrl("file-archive.svg"),
  excel: assetUrl("file-excel.svg"),
  file: assetUrl("file-generic.svg"),
  folder: assetUrl("folder.svg"),
  image: assetUrl("file-image.svg"),
  pdf: assetUrl("file-pdf.svg"),
  powerpoint: assetUrl("file-powerpoint.svg"),
  text: assetUrl("file-text.svg"),
  word: assetUrl("file-word.svg")
};

const config = {
  ...(window.APP_CONFIG || {})
};

const favoriteDialog = document.querySelector("#favorite-dialog");
const elements = {
  accountBadge: document.querySelector("#account-badge"),
  addFavoriteButton: document.querySelector("#add-favorite-button"),
  authButton: document.querySelector("#auth-button"),
  breadcrumbs: document.querySelector("#breadcrumbs"),
  cancelDialogButton: document.querySelector("#cancel-dialog-button"),
  closeDialogButton: document.querySelector("#close-dialog-button"),
  currentFolderName: document.querySelector("#current-folder-name"),
  favoriteError: document.querySelector("#favorite-error"),
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
  statusCard: document.querySelector("#status-card"),
  statusText: document.querySelector("#status-text")
};

const state = {
  currentPath: "/",
  favorites: [],
  installPrompt: null,
  items: [],
  provider: createOneDriveProvider(config),
  session: {
    configured: true,
    displayMode: "OneDrive",
    isAuthenticated: false,
    userName: "Neprihlasen"
  }
};

function assetUrl(name) {
  return new URL(`./assets/${name}`, import.meta.url).href;
}

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

function readFavorites() {
  const raw = localStorage.getItem(FAVORITES_STORAGE_KEY);
  if (!raw) {
    return (config.initialFavorites || []).map((entry) => ({
      label: entry.label,
      path: normalizePath(entry.path)
    }));
  }

  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return [];
    }

    return parsed
      .filter((entry) => entry?.label && entry?.path)
      .map((entry) => ({
        label: String(entry.label).trim(),
        path: normalizePath(String(entry.path).trim())
      }));
  } catch (error) {
    return [];
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
            <span class="favorite-head">
              <img class="favorite-icon" src="${escapeHtml(FILE_ICON_ASSETS.folder)}" alt="" aria-hidden="true">
              <span class="favorite-badge">Slozka</span>
            </span>
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

function fileIconAsset(item) {
  if (item.type === "folder") {
    return FILE_ICON_ASSETS.folder;
  }

  const extension = String(item.extension || "").toLowerCase();
  if (extension === "pdf") {
    return FILE_ICON_ASSETS.pdf;
  }
  if (extension === "doc" || extension === "docx") {
    return FILE_ICON_ASSETS.word;
  }
  if (extension === "xls" || extension === "xlsx" || extension === "csv") {
    return FILE_ICON_ASSETS.excel;
  }
  if (extension === "ppt" || extension === "pptx") {
    return FILE_ICON_ASSETS.powerpoint;
  }
  if (IMAGE_EXTENSIONS.has(extension)) {
    return FILE_ICON_ASSETS.image;
  }
  if (extension === "txt") {
    return FILE_ICON_ASSETS.text;
  }
  if (extension === "zip") {
    return FILE_ICON_ASSETS.archive;
  }

  return FILE_ICON_ASSETS.file;
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

      const fileExtension = (item.extension || "").toLowerCase();
      const showAppFinder =
        item.type === "file" &&
        (Boolean(FILE_APP_HINTS[fileExtension]) || !PREVIEWABLE_EXTENSIONS.has(fileExtension));

      return `
        <article
          class="file-card ${showAppFinder ? "file-card-has-aux" : ""}"
          data-action="${item.type === "folder" ? "open-folder" : "open-file"}"
          data-path="${escapeHtml(item.path)}"
          role="button"
          tabindex="0"
        >
          <div class="file-main">
            <span class="file-icon">
              <img class="file-icon-image" src="${escapeHtml(fileIconAsset(item))}" alt="" aria-hidden="true">
            </span>
            <div>
              <div class="file-name">${escapeHtml(item.name)}</div>
              <div class="file-meta">
                ${metaParts.map((part) => `<span>${escapeHtml(part)}</span>`).join("")}
              </div>
            </div>
          </div>
          <div class="file-actions">
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

function renderSession() {
  elements.modeBadge.textContent = "OneDrive";
  elements.accountBadge.textContent = state.session.userName;
  elements.authButton.textContent = state.session.isAuthenticated ? "Odhlasit" : "Prihlasit k OneDrive";
  elements.favoriteNote.textContent =
    "Oblibene slozky jsou ulozene lokalne v telefonu. Obsah slozek se cte z tveho OneDrive.";
}

async function refreshSession() {
  state.session = await state.provider.getSession();
  renderSession();
}

async function browse(path) {
  const safePath = normalizePath(path);
  state.currentPath = safePath;
  elements.currentFolderName.textContent = pathLabel(safePath);
  elements.goUpButton.hidden = safePath === "/";
  renderFavorites();
  renderBreadcrumbs();

  if (!state.session.isAuthenticated) {
    state.items = [];
    elements.fileList.innerHTML = "";
    setStatus("Klikni na Prihlasit k OneDrive a pak se vrat sem.", { showList: false });
    return;
  }

  setStatus("Nacitam obsah slozky...");

  try {
    const items = await state.provider.listFolder(safePath);
    state.items = items;
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
  elements.favoriteError.hidden = true;
  elements.favoriteError.textContent = "";
  favoriteDialog.showModal();
  elements.favoriteLabelInput.focus();
}

function closeFavoriteDialog() {
  favoriteDialog.close();
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

  const fileExtension = (item.extension || "").toLowerCase();
  const isPreviewable = PREVIEWABLE_EXTENSIONS.has(fileExtension);
  if (!isPreviewable) {
    openAppSearch(item.name, fileExtension);
    return;
  }

  try {
    const opened = await state.provider.openFile(item);
    if (!opened) {
      openAppSearch(item.name, fileExtension);
    }
  } catch (error) {
    setStatus(error.message || "Soubor se nepodarilo otevrit.", { isError: true });
  }
}

async function handleAuth() {
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

  elements.addFavoriteButton.addEventListener("click", openFavoriteDialog);
  elements.cancelDialogButton.addEventListener("click", closeFavoriteDialog);
  elements.closeDialogButton.addEventListener("click", closeFavoriteDialog);
  elements.goUpButton.addEventListener("click", () => browse(parentPath(state.currentPath)));

  elements.favoriteForm.addEventListener("submit", async (event) => {
    event.preventDefault();

    const label = elements.favoriteLabelInput.value.trim();
    const path = normalizePath(elements.favoritePathInput.value.trim());
    elements.favoriteError.hidden = true;
    elements.favoriteError.textContent = "";

    if (!label || !path) {
      return;
    }

    try {
      await state.provider.listFolder(path);
    } catch (error) {
      elements.favoriteError.hidden = false;
      elements.favoriteError.textContent = error.message || "Slozku se nepodarilo overit.";
      return;
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
    const target = event.target.closest("button[data-action], article[data-action]");
    if (!target) {
      return;
    }

    const { action } = target.dataset;
    if (action === "open-folder") {
      await browse(target.dataset.path || "/");
      return;
    }

    if (action === "open-file") {
      await handleFileOpen(target.dataset.path || "/");
      return;
    }

    if (action === "find-app") {
      openAppSearch(target.dataset.name || "", target.dataset.extension || "");
    }
  });

  elements.fileList.addEventListener("keydown", async (event) => {
    const card = event.target.closest("article[data-action]");
    if (!card || (event.key !== "Enter" && event.key !== " ")) {
      return;
    }

    event.preventDefault();
    const { action, path } = card.dataset;
    if (action === "open-folder") {
      await browse(path || "/");
      return;
    }

    if (action === "open-file") {
      await handleFileOpen(path || "/");
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
  state.favorites = readFavorites();
  state.currentPath = state.favorites[0]?.path || "/";
  bindEvents();
  registerInstallPrompt();
  await registerServiceWorker();

  try {
    await state.provider.initialize();
  } catch (error) {
    setStatus(error.message || "Aplikace narazila na problem pri startu.", { isError: true });
  }

  await refreshSession();
  renderFavorites();
  await browse(state.currentPath);
}

init();
