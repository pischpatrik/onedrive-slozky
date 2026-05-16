const DEMO_TREE = {
  type: "folder",
  name: "root",
  children: {
    Zaloha: {
      type: "folder",
      children: {
        "rodina-2026": {
          type: "folder",
          children: {
            "fotka-1.jpg": { type: "file", extension: "jpg", size: 2140000 },
            "fotka-2.jpg": { type: "file", extension: "jpg", size: 1980000 }
          }
        },
        "scan-obcanka.pdf": { type: "file", extension: "pdf", size: 340000 }
      }
    },
    Faktura: {
      type: "folder",
      children: {
        "2026": {
          type: "folder",
          children: {
            "leden.xlsx": { type: "file", extension: "xlsx", size: 98000 },
            "unor.xlsx": { type: "file", extension: "xlsx", size: 103000 }
          }
        },
        "sablona.docx": { type: "file", extension: "docx", size: 48000 }
      }
    },
    Plocha: {
      type: "folder",
      children: {
        "dulezite.txt": { type: "file", extension: "txt", size: 1200 },
        "navody": {
          type: "folder",
          children: {
            "iphone-navod.pdf": { type: "file", extension: "pdf", size: 520000 }
          }
        }
      }
    },
    Dokumenty: {
      type: "folder",
      children: {
        "Zakaznici": {
          type: "folder",
          children: {
            "nabidka-acer.pdf": { type: "file", extension: "pdf", size: 287000 },
            "smlouva-template.docx": { type: "file", extension: "docx", size: 52100 }
          }
        },
        "Archiv": {
          type: "folder",
          children: {
            "vykaz.zip": { type: "file", extension: "zip", size: 917000 }
          }
        }
      }
    }
  }
};

function normalizePath(path) {
  if (!path || path === "/") {
    return "/";
  }

  const parts = path.split("/").filter(Boolean);
  return `/${parts.join("/")}`;
}

function joinPath(parentPath, name) {
  const safeParent = normalizePath(parentPath);
  if (safeParent === "/") {
    return `/${name}`;
  }
  return `${safeParent}/${name}`;
}

function getNode(path) {
  const safePath = normalizePath(path);
  if (safePath === "/") {
    return DEMO_TREE;
  }

  const segments = safePath.split("/").filter(Boolean);
  let node = DEMO_TREE;

  for (const segment of segments) {
    const child = node.children?.[segment];
    if (!child) {
      throw new Error(`Slozka ${safePath} v demo rezimu neexistuje.`);
    }
    node = child;
  }

  return node;
}

function formatDemoFile(item) {
  const extension = item.extension || "bin";
  return {
    extension,
    name: item.name,
    path: item.path,
    size: item.size ?? 0,
    type: "file"
  };
}

function buildDemoItems(path) {
  const folder = getNode(path);

  if (folder.type !== "folder") {
    throw new Error("Vybrana cesta neni slozka.");
  }

  return Object.entries(folder.children ?? {})
    .map(([name, child]) => {
      const itemPath = joinPath(path, name);
      if (child.type === "folder") {
        return {
          childCount: Object.keys(child.children ?? {}).length,
          name,
          path: itemPath,
          type: "folder"
        };
      }

      return formatDemoFile({
        ...child,
        name,
        path: itemPath
      });
    })
    .sort((left, right) => {
      if (left.type !== right.type) {
        return left.type === "folder" ? -1 : 1;
      }
      return left.name.localeCompare(right.name, "cs");
    });
}

function openBlobDocument(title, bodyHtml) {
  const html = `<!doctype html>
  <html lang="cs">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>${title}</title>
      <style>
        body { font-family: Arial, sans-serif; margin: 0; padding: 2rem; background: #f7f8fb; color: #183235; }
        article { background: white; max-width: 760px; margin: 0 auto; padding: 2rem; border-radius: 24px; box-shadow: 0 12px 32px rgba(0,0,0,.08); }
        h1 { margin-top: 0; }
      </style>
    </head>
    <body>
      <article>${bodyHtml}</article>
    </body>
  </html>`;
  const blob = new Blob([html], { type: "text/html" });
  const url = URL.createObjectURL(blob);
  window.open(url, "_blank", "noopener");
  window.setTimeout(() => URL.revokeObjectURL(url), 60000);
}

export function createDemoProvider() {
  return {
    async initialize() {
      return null;
    },

    async getSession() {
      return {
        configured: true,
        displayMode: "Demo",
        isAuthenticated: false,
        userName: "Ukazkova data"
      };
    },

    async listFolder(path) {
      return buildDemoItems(path);
    },

    async openFile(item) {
      const ext = (item.extension || "").toLowerCase();
      if (ext === "txt") {
        openBlobDocument(item.name, `<h1>${item.name}</h1><p>Toto je ukazkovy textovy soubor v demo rezimu.</p>`);
        return true;
      }

      if (ext === "pdf") {
        openBlobDocument(
          item.name,
          `<h1>${item.name}</h1><p>V ostrem rezimu se zde otevre skutecny PDF soubor z OneDrive.</p>`
        );
        return true;
      }

      if (ext === "jpg" || ext === "jpeg" || ext === "png") {
        openBlobDocument(
          item.name,
          `<h1>${item.name}</h1><p>Ukazka obrazkoveho souboru. V ostrem rezimu by se otevrel soubor z OneDrive.</p>`
        );
        return true;
      }

      if (ext === "docx" || ext === "xlsx") {
        openBlobDocument(
          item.name,
          `<h1>${item.name}</h1><p>Tady uvidis, jak bude reagovat soubor Office po napojeni OneDrive.</p>`
        );
        return true;
      }

      return false;
    },

    async signIn() {
      return false;
    },

    async signOut() {
      return false;
    }
  };
}
