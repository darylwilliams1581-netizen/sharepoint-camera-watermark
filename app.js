/* FieldCam: camera + baked watermark + Microsoft Graph upload to SharePoint. Full file lives with the project zip if this commit is truncated. */
(() => {
  const $ = (id) => document.getElementById(id);
  const els = {
    video: $("video"), overlay: $("overlay"), preview: $("preview"), previewPane: $("previewPane"),
    viewfinder: $("viewfinder"), liveWatermark: $("liveWatermark"), caption: $("caption"), tag: $("tag"),
    statusChip: $("statusChip"), gpsChip: $("gpsChip"), message: $("message"),
    btnStart: $("btnStart"), btnCapture: $("btnCapture"), btnRetake: $("btnRetake"), btnFlip: $("btnFlip"),
    btnDownload: $("btnDownload"), btnUpload: $("btnUpload"), btnAuth: $("btnAuth"),
    btnSettings: $("btnSettings"), btnCloseSettings: $("btnCloseSettings"), drawer: $("drawer"), settingsForm: $("settingsForm")
  };
  const DEFAULTS = { siteUrl: "", libraryName: "Documents", folderPath: "FieldPhotos", clientId: "", tenantId: "common", orgName: "", wmText: "CONFIDENTIAL — internal use", wmPosition: "bottom-left", wmOpacity: "88", wmTimestamp: true, wmGps: true, wmUser: true, wmDiagonal: false };
  const state = { stream: null, facingMode: "environment", photoBlob: null, photoUrl: null, coords: null, account: null, msal: null };
  const loadSettings = () => { try { return { ...DEFAULTS, ...JSON.parse(localStorage.getItem("fieldcam.settings") || "{}") }; } catch { return { ...DEFAULTS }; } };
  const saveSettings = (s) => localStorage.setItem("fieldcam.settings", JSON.stringify(s));
  const settings = () => loadSettings();
  function fillSettingsForm() {
    const s = settings();
    ["siteUrl","libraryName","folderPath","clientId","tenantId","orgName","wmText","wmPosition","wmOpacity"].forEach((k) => $(k).value = s[k]);
    ["wmTimestamp","wmGps","wmUser","wmDiagonal"].forEach((k) => $(k).checked = s[k]);
  }
  const msg = (text, kind = "") => { els.message.textContent = text; els.message.className = `message ${kind}`; };
  const pad = (n) => String(n).padStart(2, "0");
  const stampNow = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`; };
  const fileStamp = () => stampNow().replace(/[: ]/g, "-");
  function watermarkLines() {
    const s = settings(); const lines = [];
    if (s.orgName) lines.push(s.orgName);
    if (els.tag.value.trim()) lines.push(els.tag.value.trim().toUpperCase());
    if (els.caption.value.trim()) lines.push(els.caption.value.trim());
    if (s.wmText) lines.push(s.wmText);
    if (s.wmTimestamp) lines.push(stampNow());
    if (s.wmUser && state.account) lines.push(state.account.username || state.account.name);
    if (s.wmGps && state.coords) lines.push(`${state.coords.latitude.toFixed(6)}, ${state.coords.longitude.toFixed(6)}`);
    return lines;
  }
  const refreshLiveWatermark = () => { els.liveWatermark.textContent = watermarkLines().join("\n"); };
  function initMsal() {
    const s = settings();
    if (!s.clientId || typeof msal === "undefined") { state.msal = null; return; }
    state.msal = new msal.PublicClientApplication({ auth: { clientId: s.clientId, authority: `https://login.microsoftonline.com/${s.tenantId || "common"}`, redirectUri: window.location.origin + window.location.pathname }, cache: { cacheLocation: "localStorage" } });
  }
  async function getToken() {
    if (!state.msal) throw new Error("Add an Azure AD client ID in Settings.");
    await state.msal.initialize();
    const scopes = ["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All"];
    const accounts = state.msal.getAllAccounts();
    if (accounts.length) {
      state.account = accounts[0];
      try { return (await state.msal.acquireTokenSilent({ scopes, account: accounts[0] })).accessToken; }
      catch { return (await state.msal.acquireTokenPopup({ scopes, account: accounts[0] })).accessToken; }
    }
    const result = await state.msal.loginPopup({ scopes });
    state.account = result.account; els.btnAuth.textContent = state.account.username; refreshLiveWatermark(); return result.accessToken;
  }
  async function signInOrOut() {
    try {
      initMsal();
      if (!state.msal) { msg("Open Settings and add your Entra ID application (client) ID.", "err"); return; }
      await state.msal.initialize();
      if (state.account || state.msal.getAllAccounts().length) {
        await state.msal.logoutPopup(); state.account = null; els.btnAuth.textContent = "Sign in"; refreshLiveWatermark(); msg("Signed out."); return;
      }
      await getToken(); msg(`Signed in as ${state.account.username}`, "ok");
    } catch (err) { msg(err.message || "Sign-in failed.", "err"); }
  }
  function stopCamera() { if (state.stream) { state.stream.getTracks().forEach((t) => t.stop()); state.stream = null; } }
  async function startCamera() {
    stopCamera();
    try {
      state.stream = await navigator.mediaDevices.getUserMedia({ audio: false, video: { facingMode: { ideal: state.facingMode }, width: { ideal: 1920 }, height: { ideal: 1080 } } });
      els.video.srcObject = state.stream; await els.video.play();
      els.statusChip.textContent = state.facingMode === "environment" ? "Rear camera" : "Front camera";
      els.btnCapture.disabled = false; els.btnStart.textContent = "Restart camera"; showViewfinder(); msg("Camera ready. Tap Capture.");
    } catch (err) { els.statusChip.textContent = "Camera blocked"; msg("Could not open camera. Allow camera permission and use HTTPS.", "err"); }
  }
  function showViewfinder() { els.previewPane.classList.add("hidden"); els.viewfinder.classList.remove("hidden"); els.btnCapture.classList.remove("hidden"); els.btnRetake.classList.add("hidden"); }
  function showPreview() { els.viewfinder.classList.add("hidden"); els.previewPane.classList.remove("hidden"); els.btnCapture.classList.add("hidden"); els.btnRetake.classList.remove("hidden"); }
  function wrapText(ctx, text, maxWidth) {
    const words = text.split(" "); const lines = []; let current = "";
    for (const word of words) { const test = current ? `${current} ${word}` : word; if (ctx.measureText(test).width > maxWidth && current) { lines.push(current); current = word; } else current = test; }
    if (current) lines.push(current); return lines;
  }
  function drawWatermark(ctx, width, height) {
    const s = settings(); const opacity = Number(s.wmOpacity || 88) / 100; const lines = watermarkLines(); if (!lines.length) return;
    const fontSize = Math.max(18, Math.round(width * 0.028)); ctx.save(); ctx.font = `600 ${fontSize}px Segoe UI, sans-serif`; ctx.textBaseline = "top";
    if (s.wmDiagonal) { ctx.globalAlpha = opacity * 0.28; ctx.fillStyle = "#ffffff"; ctx.translate(width/2, height/2); ctx.rotate(-Math.PI/6); ctx.textAlign = "center"; ctx.font = `700 ${Math.round(width*0.06)}px Segoe UI, sans-serif`; ctx.fillText(lines[0], 0, 0); ctx.restore(); ctx.save(); ctx.font = `600 ${fontSize}px Segoe UI, sans-serif`; }
    const padded = lines.flatMap((line) => wrapText(ctx, line, width * 0.62)); const lineH = fontSize * 1.28;
    const padX = Math.round(width * 0.03); const padY = Math.round(height * 0.03);
    const blockH = padded.length * lineH + fontSize * 0.6; const blockW = Math.min(width * 0.7, Math.max(...padded.map((l) => ctx.measureText(l).width)) + fontSize);
    let x = padX, y = height - padY - blockH;
    if (s.wmPosition === "bottom-right") x = width - padX - blockW;
    if (s.wmPosition === "top-left") y = padY;
    if (s.wmPosition === "top-right") { x = width - padX - blockW; y = padY; }
    if (s.wmPosition === "banner") { x = 0; y = height - blockH - padY * 0.4; }
    ctx.globalAlpha = opacity; ctx.fillStyle = "rgba(0,0,0,0.55)";
    if (s.wmPosition === "banner") { ctx.fillRect(0, y - padY * 0.3, width, blockH + padY); x = padX; } else ctx.fillRect(x - fontSize * 0.35, y - fontSize * 0.25, blockW, blockH);
    ctx.fillStyle = "#fff"; padded.forEach((line, i) => ctx.fillText(line, x, y + i * lineH)); ctx.restore();
  }
  async function capture() {
    const video = els.video; if (!video.videoWidth) { msg("Camera is not ready yet.", "err"); return; }
    const canvas = document.createElement("canvas"); canvas.width = video.videoWidth; canvas.height = video.videoHeight;
    const ctx = canvas.getContext("2d"); ctx.drawImage(video, 0, 0); drawWatermark(ctx, canvas.width, canvas.height);
    canvas.toBlob((blob) => {
      if (!blob) { msg("Could not encode JPEG.", "err"); return; }
      if (state.photoUrl) URL.revokeObjectURL(state.photoUrl);
      state.photoBlob = blob; state.photoUrl = URL.createObjectURL(blob); els.preview.src = state.photoUrl;
      els.btnDownload.disabled = false; els.btnUpload.disabled = false; showPreview();
      msg(`Captured ${(blob.size/1024).toFixed(0)} KB JPEG with watermark baked in.`, "ok");
    }, "image/jpeg", 0.92);
  }
  function downloadPhoto() { if (!state.photoUrl) return; const a = document.createElement("a"); a.href = state.photoUrl; a.download = `fieldcam-${fileStamp()}.jpg`; a.click(); }
  async function graph(path, token, options = {}) {
    const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, { ...options, headers: { Authorization: `Bearer ${token}`, ...(options.headers || {}) } });
    if (!res.ok) throw new Error(`Graph ${res.status}: ${(await res.text()).slice(0,280)}`);
    const ct = res.headers.get("content-type") || ""; return ct.includes("json") ? res.json() : res;
  }
  async function uploadToSharePoint() {
    if (!state.photoBlob) { msg("Capture a photo first.", "err"); return; }
    const s = settings(); if (!s.siteUrl) { msg("Add your SharePoint site URL in Settings.", "err"); return; }
    els.btnUpload.disabled = true; msg("Signing in and uploading…");
    try {
      initMsal(); const token = await getToken();
      const u = new URL(s.siteUrl); const site = await graph(`/sites/${u.hostname}:${u.pathname.replace(/\/$/,"") || "/"}`, token);
      const drive = await graph(`/sites/${site.id}/lists/${encodeURIComponent(s.libraryName)}/drive`, token).catch(() => graph(`/sites/${site.id}/drive`, token));
      const folder = (s.folderPath || "").replace(/^\/+|\/+$/g, ""); const fileName = `FieldCam_${fileStamp()}.jpg`;
      const itemPath = folder ? `${folder}/${fileName}` : fileName;
      const res = await fetch(`https://graph.microsoft.com/v1.0/drives/${drive.id}/root:/${itemPath}:/content`, { method: "PUT", headers: { Authorization: `Bearer ${token}`, "Content-Type": "image/jpeg" }, body: state.photoBlob });
      if (!res.ok) throw new Error(`Upload ${res.status}: ${(await res.text()).slice(0,280)}`);
      const item = await res.json(); msg(`Uploaded to ${s.libraryName}/${itemPath}`, "ok");
      if (item.webUrl) { els.message.innerHTML = `Uploaded to <a href="${item.webUrl}" target="_blank" rel="noopener">${s.libraryName}/${itemPath}</a>`; els.message.className = "message ok"; }
    } catch (err) { msg(err.message || "Upload failed.", "err"); } finally { els.btnUpload.disabled = false; }
  }
  function startGps() {
    if (!("geolocation" in navigator)) { els.gpsChip.textContent = "No GPS"; return; }
    navigator.geolocation.getCurrentPosition((pos) => { state.coords = pos.coords; els.gpsChip.textContent = `${pos.coords.latitude.toFixed(4)}, ${pos.coords.longitude.toFixed(4)}`; refreshLiveWatermark(); }, () => { els.gpsChip.textContent = "GPS denied"; }, { enableHighAccuracy: true, timeout: 12000 });
  }
  function bind() {
    els.btnStart.addEventListener("click", startCamera);
    els.btnCapture.addEventListener("click", capture);
    els.btnRetake.addEventListener("click", () => { state.photoBlob = null; els.btnDownload.disabled = true; els.btnUpload.disabled = true; showViewfinder(); msg("Ready to recapture."); });
    els.btnFlip.addEventListener("click", async () => { state.facingMode = state.facingMode === "environment" ? "user" : "environment"; await startCamera(); });
    els.btnDownload.addEventListener("click", downloadPhoto);
    els.btnUpload.addEventListener("click", uploadToSharePoint);
    els.btnAuth.addEventListener("click", signInOrOut);
    els.btnSettings.addEventListener("click", () => { fillSettingsForm(); els.drawer.hidden = false; });
    els.btnCloseSettings.addEventListener("click", () => { els.drawer.hidden = true; });
    els.settingsForm.addEventListener("submit", (e) => {
      e.preventDefault();
      saveSettings({ siteUrl: $("siteUrl").value.trim(), libraryName: $("libraryName").value.trim() || "Documents", folderPath: $("folderPath").value.trim(), clientId: $("clientId").value.trim(), tenantId: $("tenantId").value.trim() || "common", orgName: $("orgName").value.trim(), wmText: $("wmText").value.trim(), wmPosition: $("wmPosition").value, wmOpacity: $("wmOpacity").value, wmTimestamp: $("wmTimestamp").checked, wmGps: $("wmGps").checked, wmUser: $("wmUser").checked, wmDiagonal: $("wmDiagonal").checked });
      initMsal(); refreshLiveWatermark(); els.drawer.hidden = true; msg("Settings saved.", "ok");
    });
    ["caption","tag"].forEach((id) => $(id).addEventListener("input", refreshLiveWatermark));
  }
  bind(); fillSettingsForm(); initMsal(); startGps(); refreshLiveWatermark(); msg("Start the camera, capture, then upload to SharePoint.");
})();
