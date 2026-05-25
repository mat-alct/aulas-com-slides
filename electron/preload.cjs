const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("aulaBackup", {
  onRequest(callback) {
    const handler = async () => {
      try {
        const payload = await callback();
        await ipcRenderer.invoke("backup:save", payload);
        ipcRenderer.send("backup:complete", { ok: true });
      } catch (error) {
        ipcRenderer.send("backup:complete", {
          ok: false,
          error: error?.message || String(error)
        });
      }
    };
    ipcRenderer.on("backup:request", handler);
    return () => ipcRenderer.removeListener("backup:request", handler);
  }
});

contextBridge.exposeInMainWorld("aulaTools", {
  convertSlidesToPdf(buffer, originalName) {
    return ipcRenderer.invoke("slides:convertToPdf", { buffer, originalName });
  },
  mergeAudio(files) {
    return ipcRenderer.invoke("audio:merge", { files });
  },
  check() {
    return ipcRenderer.invoke("tools:check");
  }
});
