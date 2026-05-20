const App = (() => {
  const SCOPES = 'https://www.googleapis.com/auth/drive.readonly';
  const DRIVE_API = 'https://www.googleapis.com/drive/v3';

  let tokenClient = null;
  let accessToken = null;

  const state = {
    user: null,
    currentFolder: 'root',
    breadcrumbs: [{ id: 'root', name: 'My Drive' }],
    files: [],
    filteredFiles: [],
    selectedIds: new Set(),
    printing: false,
    printQueue: [],
    printCancelled: false,
  };

  const FOLDER_ICON = `<svg viewBox="0 0 24 24"><path fill="#5f6368" d="M20 6h-8l-2-2H4c-1.1 0-1.99.9-1.99 2L2 18c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V8c0-1.1-.9-2-2-2zm0 12H4V8h16v10z"/></svg>`;
  const PDF_ICON = `<svg viewBox="0 0 24 24"><path fill="#d93025" d="M20 2H8c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm0 14H8V4h12v12zM4 6H2v14c0 1.1.9 2 2 2h14v-2H4V6zm16-4H8v2h12v12h2V4c0-1.1-.9-2-2-2z"/></svg>`;

  function $(id) { return document.getElementById(id); }

  function formatSize(bytes) {
    if (!bytes) return '';
    const units = ['B', 'KB', 'MB', 'GB'];
    let i = 0;
    let size = parseInt(bytes, 10);
    while (size >= 1024 && i < units.length - 1) { size /= 1024; i++; }
    return size.toFixed(i === 0 ? 0 : 1) + ' ' + units[i];
  }

  function formatDate(iso) {
    if (!iso) return '';
    return new Date(iso).toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
  }

  async function fetchWithAuth(url, options = {}) {
    if (!accessToken) throw new Error('Not authenticated');
    const headers = { ...options.headers, Authorization: `Bearer ${accessToken}` };
    let response = await fetch(url, { ...options, headers });
    if (response.status === 401) {
      await refreshToken();
      headers.Authorization = `Bearer ${accessToken}`;
      response = await fetch(url, { ...options, headers });
    }
    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(err.error?.message || `HTTP ${response.status}`);
    }
    return response;
  }

  function refreshToken() {
    return new Promise((resolve, reject) => {
      tokenClient.callback = (resp) => {
        if (resp.error) return reject(resp);
        accessToken = resp.access_token;
        resolve();
      };
      tokenClient.requestAccessToken({ prompt: '' });
    });
  }

  function initAuth() {
    if (typeof google === 'undefined' || !google.accounts) {
      setTimeout(initAuth, 200);
      return;
    }
    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: GOOGLE_CLIENT_ID,
      scope: SCOPES,
      callback: (resp) => {
        if (resp.error) {
          showToast('Sign-in failed: ' + resp.error, 'error');
          return;
        }
        accessToken = resp.access_token;
        onSignedIn();
      },
    });
    const btn = $('sign-in-btn');
    btn.disabled = false;
    $('auth-loading').classList.add('hidden');
  }

  function signIn() {
    tokenClient.requestAccessToken();
  }

  function signOut() {
    if (accessToken) {
      google.accounts.oauth2.revoke(accessToken);
      accessToken = null;
    }
    state.user = null;
    state.currentFolder = 'root';
    state.breadcrumbs = [{ id: 'root', name: 'My Drive' }];
    state.files = [];
    state.filteredFiles = [];
    state.selectedIds.clear();
    $('auth-screen').classList.remove('hidden');
    $('main-app').classList.add('hidden');
  }

  async function onSignedIn() {
    try {
      const resp = await fetchWithAuth('https://www.googleapis.com/oauth2/v3/userinfo');
      state.user = await resp.json();
    } catch {
      state.user = { name: 'User', picture: '' };
    }
    renderUserInfo();
    $('auth-screen').classList.add('hidden');
    $('main-app').classList.remove('hidden');
    navigateTo('root', 'My Drive');
  }

  function renderUserInfo() {
    if (!state.user) return;
    $('user-name').textContent = state.user.name || '';
    $('user-avatar').src = state.user.picture || '';
    $('user-avatar').alt = state.user.name || '';
  }

  async function listFiles(folderId) {
    $('loading-state').classList.remove('hidden');
    $('file-list').classList.add('hidden');
    $('empty-state').classList.add('hidden');
    try {
      const query = `'${folderId}' in parents and trashed = false`;
      const fields = 'files(id,name,mimeType,size,modifiedTime)';
      const orderBy = 'folder,name';
      const url = `${DRIVE_API}/files?q=${encodeURIComponent(query)}&fields=${encodeURIComponent(fields)}&orderBy=${encodeURIComponent(orderBy)}&pageSize=1000`;
      const resp = await fetchWithAuth(url);
      const data = await resp.json();
      return (data.files || []).map(f => ({
        ...f,
        isFolder: f.mimeType === 'application/vnd.google-apps.folder',
        isPdf: f.mimeType === 'application/pdf',
      }));
    } finally {
      $('loading-state').classList.add('hidden');
    }
  }

  async function downloadFile(fileId) {
    const resp = await fetchWithAuth(`${DRIVE_API}/files/${fileId}?alt=media`);
    return resp.blob();
  }

  async function navigateTo(folderId, folderName) {
    state.currentFolder = folderId;
    state.selectedIds.clear();
    state.filteredFiles = [];
    $('search-input').value = '';
    $('select-all').checked = false;

    if (folderId === 'root') {
      state.breadcrumbs = [{ id: 'root', name: 'My Drive' }];
    } else {
      const idx = state.breadcrumbs.findIndex(b => b.id === folderId);
      if (idx >= 0) {
        state.breadcrumbs = state.breadcrumbs.slice(0, idx + 1);
      } else {
        state.breadcrumbs.push({ id: folderId, name: folderName });
      }
    }

    renderBreadcrumbs();
    state.files = await listFiles(folderId);
    state.filteredFiles = state.files.filter(f => f.isFolder || f.isPdf);
    renderFileList();
    updateSelectionUI();
  }

  function renderBreadcrumbs() {
    const el = $('breadcrumbs');
    el.innerHTML = state.breadcrumbs.map((b, i) => {
      const isLast = i === state.breadcrumbs.length - 1;
      const cls = isLast ? 'breadcrumb-item active' : 'breadcrumb-item';
      const sep = isLast ? '' : '<span class="breadcrumb-sep">/</span>';
      return `<span class="${cls}" onclick="App.navigateBreadcrumb(${i})">${b.name}</span>${sep}`;
    }).join('');
  }

  function navigateBreadcrumb(index) {
    const b = state.breadcrumbs[index];
    if (b) navigateTo(b.id, b.name);
  }

  function renderFileList() {
    const el = $('file-list');
    const files = state.filteredFiles;

    if (files.length === 0) {
      el.classList.add('hidden');
      $('empty-state').classList.remove('hidden');
      return;
    }

    $('empty-state').classList.add('hidden');
    el.classList.remove('hidden');

    el.innerHTML = files.map(f => {
      if (f.isFolder) {
        return `<div class="file-card" onclick="App.navigateTo('${f.id}', '${f.name.replace(/'/g, "\\'")}')">
          <div class="file-icon">${FOLDER_ICON}</div>
          <div class="file-details">
            <div class="file-name" title="${f.name}">${f.name}</div>
            <div class="file-meta">Folder</div>
          </div>
        </div>`;
      }
      if (f.isPdf) {
        const checked = state.selectedIds.has(f.id) ? 'checked' : '';
        return `<div class="file-card ${state.selectedIds.has(f.id) ? 'selected' : ''}" data-id="${f.id}">
          <div class="file-icon">${PDF_ICON}</div>
          <div class="file-details">
            <div class="file-name" title="${f.name}">${f.name}</div>
            <div class="file-meta">${formatSize(f.size)} &middot; ${formatDate(f.modifiedTime)}</div>
          </div>
          <div class="file-checkbox" onclick="event.stopPropagation()">
            <input type="checkbox" ${checked} onchange="App.toggleFile('${f.id}')">
          </div>
        </div>`;
      }
      return '';
    }).join('');
  }

  function toggleFile(fileId) {
    if (state.selectedIds.has(fileId)) {
      state.selectedIds.delete(fileId);
    } else {
      state.selectedIds.add(fileId);
    }
    updateCardSelection(fileId);
    updateSelectionUI();
  }

  function updateCardSelection(fileId) {
    const card = document.querySelector(`.file-card[data-id="${fileId}"]`);
    if (!card) return;
    const isSelected = state.selectedIds.has(fileId);
    card.classList.toggle('selected', isSelected);
    const cb = card.querySelector('input[type="checkbox"]');
    if (cb) cb.checked = isSelected;
  }

  function toggleSelectAll() {
    const pdfs = state.filteredFiles.filter(f => f.isPdf);
    const allSelected = pdfs.length > 0 && pdfs.every(f => state.selectedIds.has(f.id));
    if (allSelected) {
      pdfs.forEach(f => state.selectedIds.delete(f.id));
    } else {
      pdfs.forEach(f => state.selectedIds.add(f.id));
    }
    renderFileList();
    updateSelectionUI();
  }

  function updateSelectionUI() {
    const pdfs = state.filteredFiles.filter(f => f.isPdf);
    const count = state.selectedIds.size;
    const total = pdfs.length;
    $('selection-count').textContent = count > 0 ? `${count} of ${total} selected` : `${total} PDFs`;
    $('print-btn').disabled = count === 0;
    $('print-btn-text').textContent = `Print Selected (${count})`;
    $('select-all').checked = total > 0 && count === total;
  }

  function filterFiles(query) {
    const q = query.toLowerCase().trim();
    if (!q) {
      state.filteredFiles = state.files.filter(f => f.isFolder || f.isPdf);
    } else {
      state.filteredFiles = state.files.filter(f =>
        (f.isFolder || f.isPdf) && f.name.toLowerCase().includes(q)
      );
    }
    renderFileList();
    updateSelectionUI();
  }

  async function startBatchPrint() {
    const pdfs = state.filteredFiles
      .filter(f => f.isPdf && state.selectedIds.has(f.id));

    if (pdfs.length === 0) return;

    state.printing = true;
    state.printCancelled = false;
    state.printQueue = [...pdfs];

    $('progress-modal').classList.remove('hidden');
    $('print-btn').disabled = true;

    await printNext();
  }

  async function printNext() {
    if (state.printCancelled || state.printQueue.length === 0) {
      finishPrint();
      return;
    }

    const file = state.printQueue.shift();
    const current = state.printQueue.length > 0
      ? state.selectedIds.size - state.printQueue.length
      : state.selectedIds.size;
    const total = state.selectedIds.size;

    updateProgress(current, total, file.name);

    try {
      const blob = await downloadFile(file.id);
      if (state.printCancelled) { finishPrint(); return; }
      await printPdf(blob, file.name);
    } catch (err) {
      showToast(`Error downloading ${file.name}: ${err.message}`, 'error');
    }

    if (state.printQueue.length > 0 && !state.printCancelled) {
      await printNext();
    } else {
      finishPrint();
    }
  }

  function printPdf(blob, filename) {
    return new Promise((resolve) => {
      const iframe = $('print-frame');
      const blobUrl = URL.createObjectURL(blob);

      const onLoad = () => {
        iframe.removeEventListener('load', onLoad);
        const afterPrintHandler = () => {
          iframe.contentWindow.removeEventListener('afterprint', afterPrintHandler);
          URL.revokeObjectURL(blobUrl);
          iframe.src = 'about:blank';
          resolve();
        };

        try {
          iframe.contentWindow.addEventListener('afterprint', afterPrintHandler);
          iframe.contentWindow.focus();
          iframe.contentWindow.print();
        } catch {
          URL.revokeObjectURL(blobUrl);
          iframe.src = 'about:blank';
          resolve();
        }

        setTimeout(() => {
          URL.revokeObjectURL(blobUrl);
          try { iframe.contentWindow.removeEventListener('afterprint', afterPrintHandler); } catch {}
          iframe.src = 'about:blank';
          resolve();
        }, 60000);
      };

      iframe.addEventListener('load', onLoad);
      iframe.src = blobUrl;
    });
  }

  function cancelPrint() {
    state.printCancelled = true;
    showToast('Print job cancelled', 'warning');
    finishPrint();
  }

  function finishPrint() {
    state.printing = false;
    state.printQueue = [];
    $('progress-modal').classList.add('hidden');
    $('print-btn').disabled = state.selectedIds.size === 0;

    if (!state.printCancelled) {
      const count = state.selectedIds.size;
      showToast(`Successfully queued ${count} PDF${count !== 1 ? 's' : ''} for printing`, 'success');
    }
  }

  function updateProgress(current, total, filename) {
    $('progress-title').textContent = 'Printing PDFs...';
    $('progress-file').textContent = filename;
    $('progress-count').textContent = `${current} of ${total}`;
    $('progress-bar').style.width = `${total > 0 ? (current / total) * 100 : 0}%`;
  }

  function showToast(message, type = 'info') {
    const container = $('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    container.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transition = 'opacity 0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, 4000);
  }

  document.addEventListener('DOMContentLoaded', initAuth);

  return {
    signIn,
    signOut,
    navigateTo,
    navigateBreadcrumb,
    toggleFile,
    toggleSelectAll,
    filterFiles,
    startBatchPrint,
    cancelPrint,
  };
})();
