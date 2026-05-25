process.env.ELECTRON_DISABLE_GPU = "1";
process.env.LIBVA_DRIVER_NAME = "dummy";

const { app, BrowserWindow, ipcMain, shell } = require("electron");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const { spawn } = require("node:child_process");

let mainWindow;
let backupBeforeClose = false;
let closeBackupTimer = null;

app.disableHardwareAcceleration();
app.commandLine.appendSwitch("disable-gpu");
app.commandLine.appendSwitch("disable-gpu-compositing");
app.commandLine.appendSwitch("disable-accelerated-video-decode");
app.commandLine.appendSwitch("disable-accelerated-video-encode");
app.commandLine.appendSwitch("disable-features", "VaapiVideoDecoder,VaapiVideoEncoder,UseChromeOSDirectVideoDecoder");

function createWindow() {
  backupBeforeClose = false;
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 820,
    minWidth: 920,
    minHeight: 620,
    title: "Aulas com Slides",
    backgroundColor: "#f6f7f9",
    webPreferences: {
      preload: path.join(__dirname, "preload.cjs"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true
    }
  });

  mainWindow.removeMenu();
  mainWindow.loadFile(path.join(__dirname, "../dist/index.html"));

  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    const parsedUrl = new URL(url);
    const isInternalEditor = parsedUrl.searchParams.has("editor") && (parsedUrl.protocol === "file:" || parsedUrl.origin === "http://localhost:5173");
    if (isInternalEditor) {
      return {
        action: "allow",
        overrideBrowserWindowOptions: {
          width: 1100,
          height: 760,
          minWidth: 900,
          minHeight: 620,
          title: "Editor de Comentarios",
          backgroundColor: "#f6f7f9",
          webPreferences: {
            preload: path.join(__dirname, "preload.cjs"),
            contextIsolation: true,
            nodeIntegration: false,
            sandbox: true
          }
        }
      };
    }
    shell.openExternal(url);
    return { action: "deny" };
  });

  mainWindow.on("close", (event) => {
    if (backupBeforeClose || mainWindow.webContents.isDestroyed()) return;
    event.preventDefault();
    mainWindow.webContents.send("backup:request");
    closeBackupTimer = setTimeout(() => finishClose(), 8000);
  });
}

function finishClose() {
  if (!mainWindow || mainWindow.isDestroyed()) return;
  if (closeBackupTimer) {
    clearTimeout(closeBackupTimer);
    closeBackupTimer = null;
  }
  backupBeforeClose = true;
  mainWindow.close();
}

async function saveBackup(payload) {
  if (!payload || !Array.isArray(payload.lessons)) {
    throw new Error("Backup invalido");
  }

  const backupDir = path.join(app.getPath("userData"), "backups");
  await fs.mkdir(backupDir, { recursive: true });
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const filename = `aulas-com-slides-${stamp}.backup.json`;
  const filePath = path.join(backupDir, filename);
  await fs.writeFile(filePath, JSON.stringify(payload, null, 2));
  await pruneBackups(backupDir, 20);
  return { filePath };
}

async function pruneBackups(backupDir, keepCount) {
  const entries = await fs.readdir(backupDir, { withFileTypes: true });
  const backups = entries
    .filter((entry) => entry.isFile() && entry.name.endsWith(".backup.json"))
    .map((entry) => entry.name)
    .sort()
    .reverse();

  await Promise.all(backups.slice(keepCount).map((name) => fs.unlink(path.join(backupDir, name))));
}

ipcMain.handle("backup:save", async (_event, payload) => saveBackup(payload));

ipcMain.handle("slides:convertToPdf", async (_event, payload) => convertSlidesToPdf(payload));
ipcMain.handle("audio:merge", async (_event, payload) => mergeAudioFiles(payload));
ipcMain.handle("tools:check", async () => checkExternalTools());

function runCommand(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { ...options });
    let stdout = "";
    let stderr = "";
    child.stdout?.on("data", (chunk) => {
      stdout += chunk.toString();
    });
    child.stderr?.on("data", (chunk) => {
      stderr += chunk.toString();
    });
    child.on("error", reject);
    child.on("close", (code) => {
      if (code === 0) resolve({ stdout, stderr });
      else reject(new Error(`${command} encerrou com codigo ${code}: ${stderr || stdout}`));
    });
  });
}

async function makeTempDir(prefix) {
  return fs.mkdtemp(path.join(os.tmpdir(), prefix));
}

async function removeDir(dir) {
  try {
    await fs.rm(dir, { recursive: true, force: true });
  } catch (error) {
    console.warn("Falha ao limpar diretorio temporario", dir, error);
  }
}

function sanitizeBaseName(name) {
  const base = path.basename(String(name || ""), path.extname(String(name || ""))) || "arquivo";
  return base.replace(/[^a-zA-Z0-9._-]+/g, "_").slice(0, 80) || "arquivo";
}

async function convertSlidesToPdf({ buffer, originalName } = {}) {
  if (!buffer || !buffer.byteLength) {
    throw new Error("Arquivo de slides vazio");
  }
  const ext = (path.extname(String(originalName || "")).toLowerCase().replace(".", "")) || "pptx";
  if (!["ppt", "pptx", "odp", "ppsx", "pps", "key"].includes(ext)) {
    throw new Error(`Formato de slide nao suportado: .${ext}`);
  }
  const tmp = await makeTempDir("aula-slides-");
  try {
    const base = sanitizeBaseName(originalName);
    const inputPath = path.join(tmp, `${base}.${ext}`);
    await fs.writeFile(inputPath, Buffer.from(buffer));
    await runCommand("soffice", [
      "--headless",
      "--norestore",
      "--nologo",
      "--nodefault",
      "--nolockcheck",
      "--convert-to",
      "pdf",
      "--outdir",
      tmp,
      inputPath
    ]);
    const outputPath = path.join(tmp, `${base}.pdf`);
    const pdf = await fs.readFile(outputPath);
    return pdf.buffer.slice(pdf.byteOffset, pdf.byteOffset + pdf.byteLength);
  } finally {
    removeDir(tmp);
  }
}

async function mergeAudioFiles({ files } = {}) {
  if (!Array.isArray(files) || !files.length) {
    throw new Error("Nenhum audio para mesclar");
  }
  if (files.length === 1) {
    const single = files[0];
    return {
      buffer: single.buffer,
      mimeType: single.mimeType || "audio/mp4",
      extension: path.extname(single.name || "").replace(".", "") || "m4a"
    };
  }

  const tmp = await makeTempDir("aula-audio-");
  try {
    const inputArgs = [];
    for (let index = 0; index < files.length; index += 1) {
      const file = files[index];
      const ext = path.extname(file.name || "").replace(".", "") || "audio";
      const safe = path.join(tmp, `input-${String(index).padStart(3, "0")}.${ext}`);
      await fs.writeFile(safe, Buffer.from(file.buffer));
      inputArgs.push("-i", safe);
    }
    const filter = `${files.map((_, i) => `[${i}:a]`).join("")}concat=n=${files.length}:v=0:a=1[out]`;
    const outputPath = path.join(tmp, "merged.m4a");
    await runCommand("ffmpeg", [
      "-y",
      "-hide_banner",
      "-loglevel",
      "error",
      ...inputArgs,
      "-filter_complex",
      filter,
      "-map",
      "[out]",
      "-c:a",
      "aac",
      "-b:a",
      "192k",
      "-movflags",
      "+faststart",
      outputPath
    ]);
    const merged = await fs.readFile(outputPath);
    return {
      buffer: merged.buffer.slice(merged.byteOffset, merged.byteOffset + merged.byteLength),
      mimeType: "audio/mp4",
      extension: "m4a"
    };
  } finally {
    removeDir(tmp);
  }
}

async function checkExternalTools() {
  const check = async (cmd, args) => {
    try {
      await runCommand(cmd, args);
      return true;
    } catch {
      return false;
    }
  };
  const [soffice, ffmpeg] = await Promise.all([
    check("soffice", ["--version"]),
    check("ffmpeg", ["-version"])
  ]);
  return { soffice, ffmpeg };
}

ipcMain.on("backup:complete", (_event, result) => {
  if (!result?.ok) {
    console.error("Backup automatico falhou:", result?.error || "erro desconhecido");
  }
  finishClose();
});

app.whenReady().then(() => {
  createWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});
