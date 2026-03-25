/**
 * 仪表盘：拖拽上传、fetch 提交（携带 cookie 以使用 Session）
 */
(function () {
  let formatFile = null;
  let publishFile = null;
  let lastFileId = null;

  const formatDrop = document.getElementById("format-drop");
  const formatFileInput = document.getElementById("format-file");
  const formatPick = document.getElementById("format-pick");
  const formatName = document.getElementById("format-name");
  const formatBtn = document.getElementById("format-btn");
  const formatStatus = document.getElementById("format-status");
  const formatDownloadWrap = document.getElementById("format-download-wrap");
  const formatDownload = document.getElementById("format-download");

  const publishDrop = document.getElementById("publish-drop");
  const publishFileInput = document.getElementById("publish-file");
  const publishPick = document.getElementById("publish-pick");
  const publishName = document.getElementById("publish-name");
  const publishBtn = document.getElementById("publish-btn");
  const publishStatus = document.getElementById("publish-status");
  const publishUrlWrap = document.getElementById("publish-url-wrap");
  const publishUrl = document.getElementById("publish-url");
  const publishBlobHint = document.getElementById("publish-blob-hint");
  const formatIdInput = document.getElementById("format-id");

  function isDocx(f) {
    if (!f) return false;
    const n = f.name.toLowerCase();
    return n.endsWith(".docx");
  }

  function bindDrop(zone, input, setFile, nameEl, btn) {
    zone.addEventListener("click", () => input.click());
    zone.addEventListener("dragover", (e) => {
      e.preventDefault();
      zone.classList.add("dragover");
    });
    zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));
    zone.addEventListener("drop", (e) => {
      e.preventDefault();
      zone.classList.remove("dragover");
      const f = e.dataTransfer.files[0];
      if (isDocx(f)) {
        setFile(f);
        nameEl.textContent = f.name;
        btn.disabled = false;
      } else {
        nameEl.textContent = "请选择 .docx 文件";
      }
    });
    input.addEventListener("change", () => {
      const f = input.files[0];
      if (isDocx(f)) {
        setFile(f);
        nameEl.textContent = f.name;
        btn.disabled = false;
      }
    });
  }

  bindDrop(
    formatDrop,
    formatFileInput,
    (f) => {
      formatFile = f;
    },
    formatName,
    formatBtn
  );
  formatPick.addEventListener("click", (e) => {
    e.stopPropagation();
    formatFileInput.click();
  });

  bindDrop(
    publishDrop,
    publishFileInput,
    (f) => {
      publishFile = f;
    },
    publishName,
    publishBtn
  );
  publishPick.addEventListener("click", (e) => {
    e.stopPropagation();
    publishFileInput.click();
  });

  formatBtn.addEventListener("click", async () => {
    if (!formatFile) return;
    formatStatus.textContent = "正在上传并格式化…";
    formatDownloadWrap.classList.add("d-none");
    const fd = new FormData();
    fd.append("file", formatFile);
    try {
      const res = await fetch("/dashboard/format", {
        method: "POST",
        body: fd,
        credentials: "same-origin",
      });
      const data = await res.json();
      if (!res.ok || !data.ok) {
        formatStatus.textContent = data.error || "格式化失败";
        return;
      }
      lastFileId = data.file_id;
      if (formatIdInput && lastFileId) {
        formatIdInput.value = String(lastFileId);
      }
      formatStatus.textContent = data.message || "完成";
      formatDownload.href = data.download_url;
      formatDownloadWrap.classList.remove("d-none");
    } catch (err) {
      formatStatus.textContent = "网络错误，请稍后重试";
    }
  });

  publishBtn.addEventListener("click", async () => {
    if (!publishFile) return;
    publishStatus.textContent = "正在上传并发布…";
    publishUrlWrap.classList.add("d-none");
    const fd = new FormData();
    fd.append("file", publishFile);
    const fid = formatIdInput && formatIdInput.value.trim();
    if (fid) {
      fd.append("format_id", fid);
    }
    try {
      const res = await fetch("/dashboard/publish", {
        method: "POST",
        body: fd,
        credentials: "same-origin",
      });
      const data = await res.json();
      if (!res.ok || !data.ok) {
        publishStatus.textContent = data.error || "发布失败";
        return;
      }
      publishStatus.textContent = data.message || "已发布";
      const direct =
        data.github_raw_url || data.github_url || "";
      publishUrl.value = direct;
      if (publishBlobHint && data.github_blob_url && data.github_blob_url !== direct) {
        publishBlobHint.textContent =
          "仓库浏览页：" + data.github_blob_url;
        publishBlobHint.classList.remove("d-none");
      } else if (publishBlobHint) {
        publishBlobHint.textContent = "";
        publishBlobHint.classList.add("d-none");
      }
      publishUrlWrap.classList.remove("d-none");
    } catch (err) {
      publishStatus.textContent = "网络错误，请稍后重试";
    }
  });
})();
