const App = (() => {
  const SCOPES = 'https://www.googleapis.com/auth/drive.readonly openid profile email';
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
    lastSelectedId: null,
    viewMode: 'grid',
    printing: false,
    printQueue: [],
    printCancelled: false,
    driveMode: 'my',
    sharedDrives: [],
    recursive: false,
    activeFilter: 'all',
    sortField: 'name',
    sortAsc: true,
  };

  const FOLDER_ICON = `<svg viewBox="0 0 24 24"><path fill="#5f6368" d="M20 6h-8l-2-2H4c-1.1 0-1.99.9-1.99 2L2 18c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V8c0-1.1-.9-2-2-2zm0 12H4V8h16v10z"/></svg>`;
  const PDF_ICON = `<svg viewBox="0 0 24 24"><path fill="#d93025" d="M20 2H8c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm0 14H8V4h12v12zM4 6H2v14c0 1.1.9 2 2 2h14v-2H4V6zm16-4H8v2h12v12h2V4c0-1.1-.9-2-2-2z"/></svg>`;
  const DOC_ICON = `<svg viewBox="0 0 24 24"><path fill="#4285f4" d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/></svg>`;
  const SHEET_ICON = `<svg viewBox="0 0 24 24"><path fill="#0f9d58" d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm-2 14H8v-2h4v2zm0-4H8v-2h4v2zm0-4H8V6h4v2zm4 8h-2v-2h2v2zm0-4h-2v-2h2v2zm-3-5V3.5L18.5 9H13z"/></svg>`;
  const SLIDE_ICON = `<svg viewBox="0 0 24 24"><path fill="#f4b400" d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm-1 12H7v-2h6v2zm0-4H7V8h6v2zm4 8h-2v-2h2v2zm0-4h-2v-2h2v2zm-3-5V3.5L18.5 9H14z"/></svg>`;
  const FILE_ICON = `<svg viewBox="0 0 24 24"><path fill="#5f6368" d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm4 18H6V4h7v5h5v11z"/></svg>`;

  const GOOGLE_DOC = 'application/vnd.google-apps.document';
  const GOOGLE_SHEET = 'application/vnd.google-apps.spreadsheet';
  const GOOGLE_SLIDE = 'application/vnd.google-apps.presentation';
  const GOOGLE_DRAWING = 'application/vnd.google-apps.drawing';
  const PDF = 'application/pdf';
  const FOLDER = 'application/vnd.google-apps.folder';
  const EXPORTABLE = [GOOGLE_DOC, GOOGLE_SHEET, GOOGLE_SLIDE, GOOGLE_DRAWING];
  const OFFICE_TYPES = [
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'application/msword',
    'application/vnd.ms-excel',
    'application/vnd.ms-powerpoint',
  ];

  function isPrintable(mimeType) {
    return mimeType === PDF || EXPORTABLE.includes(mimeType) || OFFICE_TYPES.includes(mimeType);
  }

  function getFileIcon(mimeType) {
    if (mimeType === PDF) return PDF_ICON;
    if (mimeType === GOOGLE_DOC) return DOC_ICON;
    if (mimeType === GOOGLE_SHEET) return SHEET_ICON;
    if (mimeType === GOOGLE_SLIDE) return SLIDE_ICON;
    if (EXPORTABLE.includes(mimeType) || OFFICE_TYPES.includes(mimeType)) return DOC_ICON;
    return FILE_ICON;
  }

  function getFileTypeLabel(mimeType) {
    if (mimeType === PDF) return 'PDF';
    if (mimeType === GOOGLE_DOC) return 'Google Doc';
    if (mimeType === GOOGLE_SHEET) return 'Google Sheet';
    if (mimeType === GOOGLE_SLIDE) return 'Google Slide';
    if (mimeType === GOOGLE_DRAWING) return 'Google Drawing';
    if (mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') return 'Word Document';
    if (mimeType === 'application/msword') return 'Word Document';
    if (mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') return 'Excel Spreadsheet';
    if (mimeType === 'application/vnd.ms-excel') return 'Excel Spreadsheet';
    if (mimeType === 'application/vnd.openxmlformats-officedocument.presentationml.presentation') return 'PowerPoint';
    if (mimeType === 'application/vnd.ms-powerpoint') return 'PowerPoint';
    return 'Document';
  }

  function $(id) { return document.getElementById(id); }

  function esc(s) { const d = document.createElement('div'); d.textContent = s; return d.innerHTML; }

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
    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(err.error?.message || `HTTP ${response.status}`);
    }
    return response;
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
    state.driveMode = 'my';
    state.recursive = false;
    state.activeFilter = 'all';
    state.sortField = 'name';
    state.sortAsc = true;
    $('auth-screen').classList.remove('hidden');
    $('main-app').classList.add('hidden');
  }

  async function onSignedIn() {
    try {
      const resp = await fetchWithAuth('https://www.googleapis.com/oauth2/v3/userinfo');
      state.user = await resp.json();
    } catch (e) {
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
      if (state.driveMode === 'shared' && folderId === 'root') {
        return await listSharedDrives();
      }

      const query = `'${folderId}' in parents and trashed = false`;
      const fields = 'files(id,name,mimeType,size,modifiedTime)';
      const orderBy = 'folder,name';
      let url = `${DRIVE_API}/files?q=${encodeURIComponent(query)}&fields=${encodeURIComponent(fields)}&orderBy=${encodeURIComponent(orderBy)}&pageSize=1000`;
      if (state.driveMode === 'shared') {
        url += '&supportsAllDrives=true&includeItemsFromAllDrives=true';
        const driveId = state.breadcrumbs.length > 0 ? state.breadcrumbs[0].id : '';
        if (driveId && driveId !== 'root') {
          url += `&driveId=${driveId}&corpora=drive`;
        }
      }
      const resp = await fetchWithAuth(url);
      const data = await resp.json();
      return (data.files || []).map(f => ({
        id: f.id,
        name: f.name,
        mimeType: f.mimeType,
        size: f.size,
        modifiedTime: f.modifiedTime,
      }));
    } finally {
      $('loading-state').classList.add('hidden');
    }
  }

  async function listSharedDrives() {
    try {
      const resp = await fetchWithAuth(`${DRIVE_API}/drives?pageSize=100`);
      const data = await resp.json();
      state.sharedDrives = data.drives || [];
      return state.sharedDrives.map(d => ({
        id: d.id,
        name: d.name,
        mimeType: FOLDER,
        size: null,
        modifiedTime: null,
        isSharedDrive: true,
      }));
    } catch {
      showToast('Could not load shared drives', 'error');
      return [];
    }
  }

  async function downloadFile(fileId, mimeType) {
    if (mimeType === PDF) {
      const resp = await fetchWithAuth(`${DRIVE_API}/files/${fileId}?alt=media`);
      return resp.blob();
    }

    if (mimeType === GOOGLE_SHEET || mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || mimeType === 'application/vnd.ms-excel') {
      try {
        const resp = await fetchWithAuth(
          `https://docs.google.com/spreadsheets/d/${fileId}/export?format=pdf&portrait=false&fitw=true&size=A4`
        );
        return resp.blob();
      } catch {
        throw new Error('Spreadsheet cannot be converted to PDF. Save as Google Sheet first.');
      }
    }

    if (mimeType === GOOGLE_DOC || mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || mimeType === 'application/msword') {
      try {
        const resp = await fetchWithAuth(
          `https://docs.google.com/document/d/${fileId}/export?format=pdf`
        );
        return resp.blob();
      } catch {
        throw new Error('Document cannot be converted to PDF. Save as Google Doc first.');
      }
    }

    if (mimeType === GOOGLE_SLIDE || mimeType === 'application/vnd.openxmlformats-officedocument.presentationml.presentation' || mimeType === 'application/vnd.ms-powerpoint') {
      try {
        const resp = await fetchWithAuth(
          `https://docs.google.com/presentation/d/${fileId}/export?format=pdf`
        );
        return resp.blob();
      } catch {
        throw new Error('Presentation cannot be converted to PDF. Save as Google Slides first.');
      }
    }

    if (mimeType === GOOGLE_DRAWING) {
      const resp = await fetchWithAuth(
        `${DRIVE_API}/files/${fileId}/export?mimeType=${encodeURIComponent('application/pdf')}`
      );
      return resp.blob();
    }

    throw new Error('This file type cannot be converted to PDF.');
  }

  async function navigateTo(folderId, folderName) {
    state.currentFolder = folderId;
    state.selectedIds.clear();
    state.lastSelectedId = null;
    state.filteredFiles = [];
    $('search-input').value = '';
    $('select-all').checked = false;

    if (state.driveMode === 'shared' && folderId === 'root') {
      state.breadcrumbs = [{ id: 'root', name: 'Shared Drives' }];
    } else if (folderId === 'root') {
      state.breadcrumbs = [{ id: 'root', name: 'My Drive' }];
    } else {
      const idx = state.breadcrumbs.findIndex(b => b.id === folderId);
      if (idx >= 0) {
        state.breadcrumbs = state.breadcrumbs.slice(0, idx + 1);
      } else {
        const file = state.files.find(f => f.id === folderId);
        if (state.driveMode === 'shared' && (file?.isSharedDrive || state.sharedDrives.some(d => d.id === folderId))) {
          state.breadcrumbs = [{ id: folderId, name: folderName }];
        } else {
          state.breadcrumbs.push({ id: folderId, name: folderName });
        }
      }
    }

    renderBreadcrumbs();
    state.files = await listFiles(folderId);

    if (state.recursive && folderId !== 'root' && !(state.driveMode === 'shared' && folderId === 'root')) {
      state.files = await collectRecursiveFiles(state.files, folderId, 0);
    }

    applyFilters();
    renderFileList();
    updateSelectionUI();
  }

  async function collectRecursiveFiles(files, parentFolderId, depth) {
    if (depth > 10) return files;
    const folders = files.filter(f => f.mimeType === FOLDER);
    const nonFolders = files.filter(f => f.mimeType !== FOLDER);
    for (const folder of folders) {
      try {
        const subFiles = await listFilesSilent(folder.id);
        const collected = await collectRecursiveFiles(subFiles, folder.id, depth + 1);
        nonFolders.push(...collected);
      } catch (e) {
        console.warn(`Could not access subfolder ${folder.name}:`, e);
      }
    }
    return nonFolders;
  }

  async function listFilesSilent(folderId) {
    const query = `'${folderId}' in parents and trashed = false`;
    const fields = 'files(id,name,mimeType,size,modifiedTime)';
    const orderBy = 'folder,name';
    let url = `${DRIVE_API}/files?q=${encodeURIComponent(query)}&fields=${encodeURIComponent(fields)}&orderBy=${encodeURIComponent(orderBy)}&pageSize=1000`;
    if (state.driveMode === 'shared') {
      url += '&supportsAllDrives=true&includeItemsFromAllDrives=true';
      const driveId = state.breadcrumbs.length > 0 ? state.breadcrumbs[0].id : '';
      if (driveId && driveId !== 'root') {
        url += `&driveId=${driveId}&corpora=drive`;
      }
    }
    const resp = await fetchWithAuth(url);
    const data = await resp.json();
    return (data.files || []).map(f => ({
      id: f.id,
      name: f.name,
      mimeType: f.mimeType,
      size: f.size,
      modifiedTime: f.modifiedTime,
    }));
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
    const isListView = state.viewMode === 'list';

    if (files.length === 0) {
      el.classList.add('hidden');
      $('empty-state').classList.remove('hidden');
      return;
    }

    $('empty-state').classList.add('hidden');
    el.classList.remove('hidden');
    el.classList.toggle('list-view', isListView);

    el.innerHTML = files.map(f => {
      const safeName = esc(f.name);
      if (f.mimeType === FOLDER) {
        const folderMeta = isListView ? `<span class="file-meta-inline">Folder</span>` : `<div class="file-meta">Folder</div>`;
        return `<div class="file-card file-folder" onclick="App.navigateTo('${f.id}', '${f.name.replace(/'/g, "\\'")}')">
          <div class="file-icon">${FOLDER_ICON}</div>
          <div class="file-details">
            <div class="file-name" title="${safeName}">${safeName}</div>
            ${folderMeta}
          </div>
        </div>`;
      }
      if (isPrintable(f.mimeType)) {
        const checked = state.selectedIds.has(f.id) ? 'checked' : '';
        const icon = getFileIcon(f.mimeType);
        const typeLabel = getFileTypeLabel(f.mimeType);
        const sizeStr = f.size ? ` &middot; ${formatSize(f.size)}` : '';
        const isLast = state.lastSelectedId === f.id;
        const selectedCls = state.selectedIds.has(f.id) ? ' selected' : '';
        const lastCls = isLast ? ' last-selected' : '';
        if (isListView) {
          return `<div class="file-card file-printable${selectedCls}${lastCls}" data-id="${f.id}" onclick="App.toggleFile('${f.id}')">
            <div class="file-checkbox" onclick="event.stopPropagation()">
              <input type="checkbox" ${checked} onchange="App.toggleFile('${f.id}')">
            </div>
            <div class="file-icon">${icon}</div>
            <div class="file-details">
              <div class="file-name" title="${safeName}">${safeName}</div>
            </div>
            <div class="file-meta-inline">${typeLabel}${sizeStr}</div>
            <div class="file-meta-inline file-meta-date">${formatDate(f.modifiedTime)}</div>
          </div>`;
        }
        return `<div class="file-card file-printable${selectedCls}${lastCls}" data-id="${f.id}" onclick="App.toggleFile('${f.id}')">
          <div class="file-icon">${icon}</div>
          <div class="file-details">
            <div class="file-name" title="${safeName}">${safeName}</div>
            <div class="file-meta">${typeLabel}${sizeStr} &middot; ${formatDate(f.modifiedTime)}</div>
          </div>
          <div class="file-checkbox" onclick="event.stopPropagation()">
            <input type="checkbox" ${checked} onchange="App.toggleFile('${f.id}')">
          </div>
        </div>`;
      }
      return '';
    }).join('');
  }

  function recalcLastSelected() {
    const printables = state.filteredFiles.filter(f => isPrintable(f.mimeType));
    for (let i = printables.length - 1; i >= 0; i--) {
      if (state.selectedIds.has(printables[i].id)) {
        state.lastSelectedId = printables[i].id;
        return;
      }
    }
    state.lastSelectedId = null;
  }

  function toggleFile(fileId) {
    const previousLastId = state.lastSelectedId;
    if (state.selectedIds.has(fileId)) {
      state.selectedIds.delete(fileId);
    } else {
      state.selectedIds.add(fileId);
    }
    recalcLastSelected();
    updateCardSelection(fileId);
    if (previousLastId && previousLastId !== state.lastSelectedId) {
      updateCardSelection(previousLastId);
    }
    if (state.lastSelectedId && state.lastSelectedId !== fileId) {
      updateCardSelection(state.lastSelectedId);
    }
    updateSelectionUI();
  }

  function updateCardSelection(fileId) {
    const card = document.querySelector(`.file-card[data-id="${fileId}"]`);
    if (!card) return;
    const isSelected = state.selectedIds.has(fileId);
    card.classList.toggle('selected', isSelected);
    card.classList.toggle('last-selected', state.lastSelectedId === fileId);
    const cb = card.querySelector('input[type="checkbox"]');
    if (cb) cb.checked = isSelected;
  }

  function toggleSelectAll() {
    const printables = state.filteredFiles.filter(f => isPrintable(f.mimeType));
    const allSelected = printables.length > 0 && printables.every(f => state.selectedIds.has(f.id));
    if (allSelected) {
      printables.forEach(f => state.selectedIds.delete(f.id));
    } else {
      printables.forEach(f => state.selectedIds.add(f.id));
    }
    recalcLastSelected();
    renderFileList();
    updateSelectionUI();
  }

  function updateSelectionUI() {
    const printables = state.filteredFiles.filter(f => isPrintable(f.mimeType));
    const count = state.selectedIds.size;
    const total = printables.length;
    $('selection-count').textContent = count > 0 ? `${count} of ${total} selected` : `${total} documents`;
    $('print-btn').disabled = count === 0;
    $('print-btn-text').textContent = `Print Selected (${count})`;
    $('download-btn').disabled = count === 0;
    $('download-btn-text').textContent = count > 0 ? `Download (${count})` : 'Download';
    $('select-all').checked = total > 0 && count === total;
  }

  function applyFilters() {
    const q = ($('search-input').value || '').toLowerCase().trim();
    let result = state.files.filter(f => f.mimeType === FOLDER || isPrintable(f.mimeType));

    if (q) {
      result = result.filter(f => f.mimeType === FOLDER || f.name.toLowerCase().includes(q));
    }

    if (state.activeFilter !== 'all') {
      result = result.filter(f => {
        if (f.mimeType === FOLDER) return true;
        return getFilterCategory(f.mimeType) === state.activeFilter;
      });
    }

    state.filteredFiles = sortFiles(result);
  }

  function getFilterCategory(mimeType) {
    if (mimeType === PDF) return 'pdf';
    if (mimeType === GOOGLE_DOC || mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || mimeType === 'application/msword') return 'doc';
    if (mimeType === GOOGLE_SHEET || mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || mimeType === 'application/vnd.ms-excel') return 'sheet';
    if (mimeType === GOOGLE_SLIDE || mimeType === 'application/vnd.openxmlformats-officedocument.presentationml.presentation' || mimeType === 'application/vnd.ms-powerpoint') return 'slide';
    return 'other';
  }

  function sortFiles(files) {
    const folders = files.filter(f => f.mimeType === FOLDER);
    const docs = files.filter(f => f.mimeType !== FOLDER);
    const sorted = [...folders, ...docs.sort((a, b) => {
      let cmp = 0;
      if (state.sortField === 'name') {
        cmp = a.name.localeCompare(b.name);
      } else if (state.sortField === 'type') {
        cmp = getFileTypeLabel(a.mimeType).localeCompare(getFileTypeLabel(b.mimeType));
        if (cmp === 0) cmp = a.name.localeCompare(b.name);
      } else if (state.sortField === 'date') {
        const da = a.modifiedTime ? new Date(a.modifiedTime).getTime() : 0;
        const db = b.modifiedTime ? new Date(b.modifiedTime).getTime() : 0;
        cmp = da - db;
        if (cmp === 0) cmp = a.name.localeCompare(b.name);
      }
      return state.sortAsc ? cmp : -cmp;
    })];
    return sorted;
  }

  function filterFiles(query) {
    applyFilters();
    renderFileList();
    updateSelectionUI();
  }

  function setFilter(filter) {
    state.activeFilter = filter;
    document.querySelectorAll('.filter-tab').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.filter === filter);
    });
    applyFilters();
    renderFileList();
    updateSelectionUI();
  }

  function sortBy(field) {
    if (state.sortField === field) {
      state.sortAsc = !state.sortAsc;
    } else {
      state.sortField = field;
      state.sortAsc = true;
    }
    updateSortArrows();
    applyFilters();
    renderFileList();
  }

  function updateSortArrows() {
    ['name', 'type', 'date'].forEach(f => {
      const el = $('sort-' + f);
      if (el) {
        if (f === state.sortField) {
          el.textContent = state.sortAsc ? '\u25B2' : '\u25BC';
        } else {
          el.textContent = '';
        }
      }
    });
  }

  function switchDrive(mode) {
    state.driveMode = mode;
    $('drive-my-btn').classList.toggle('active', mode === 'my');
    $('drive-shared-btn').classList.toggle('active', mode === 'shared');
    navigateTo('root', mode === 'my' ? 'My Drive' : 'Shared Drives');
  }

  function toggleRecursive() {
    state.recursive = $('recursive-check').checked;
    navigateTo(state.currentFolder, state.breadcrumbs[state.breadcrumbs.length - 1]?.name || 'My Drive');
  }

  function setViewMode(mode) {
    state.viewMode = mode;
    $('view-grid-btn').classList.toggle('active', mode === 'grid');
    $('view-list-btn').classList.toggle('active', mode === 'list');
    $('list-header').classList.toggle('hidden', mode !== 'list');
    renderFileList();
  }

  async function startBatchPrint() {
    await _processBatch('print');
  }

  async function startBatchDownload() {
    await _processBatch('download');
  }

  async function _processBatch(mode) {
    const files = state.filteredFiles
      .filter(f => isPrintable(f.mimeType) && state.selectedIds.has(f.id));

    if (files.length === 0) return;

    state.printing = true;
    state.printCancelled = false;

    $('progress-modal').classList.remove('hidden');
    $('print-btn').disabled = true;
    $('download-btn').disabled = true;

    try {
      const blobs = [];

      for (let i = 0; i < files.length; i++) {
        if (state.printCancelled) { finishPrint(); return; }
        updateProgress(i, files.length, `Converting: ${files[i].name}`);
        try {
          const blob = await downloadFile(files[i].id, files[i].mimeType);
          if (blob.type && !blob.type.includes('pdf')) {
            showToast(`Skipped ${files[i].name}: not a convertible document`, 'warning');
            continue;
          }
          blobs.push({ blob, name: files[i].name });
        } catch (e) {
          showToast(`Skipped ${files[i].name}: ${e.message}`, 'warning');
        }
      }

      if (state.printCancelled) { finishPrint(); return; }

      updateProgress(files.length, files.length, 'Merging documents...');
      const mergedBlob = await mergePdfs(blobs);

      if (state.printCancelled) { finishPrint(); return; }

      if (mode === 'download') {
        updateProgress(files.length, files.length, 'Downloading...');
        const url = URL.createObjectURL(mergedBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `merged-${files.length}-documents.pdf`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        showToast(`Downloaded ${files.length} document${files.length !== 1 ? 's' : ''} as merged PDF`, 'success');
      } else {
        updateProgress(files.length, files.length, 'Opening print dialog...');
        await printPdf(mergedBlob, `${files.length} documents merged`);
        showToast(`Sent ${files.length} PDF${files.length !== 1 ? 's' : ''} to printer as one job`, 'success');
      }

      finishPrint();
    } catch (err) {
      console.error('Batch error:', err);
      showToast(`Operation failed: ${err.message}`, 'error');
      finishPrint();
    }
  }

  async function mergePdfs(pdfItems) {
    const { PDFDocument } = PDFLib;
    const mergedPdf = await PDFDocument.create();

    for (const item of pdfItems) {
      const arrayBuffer = await item.blob.arrayBuffer();
      try {
        const donorPdf = await PDFDocument.load(arrayBuffer, { ignoreEncryption: true });
        const pages = await mergedPdf.copyPages(donorPdf, donorPdf.getPageIndices());
        pages.forEach(page => mergedPdf.addPage(page));
      } catch (e) {
        console.warn(`Skipping unmergeable file: ${item.name}`, e);
      }
    }

    const mergedBytes = await mergedPdf.save();
    return new Blob([mergedBytes], { type: 'application/pdf' });
  }

  function printPdf(blob, filename) {
    return new Promise((resolve) => {
      const iframe = $('print-frame');
      const blobUrl = URL.createObjectURL(blob);
      let resolved = false;

      const done = () => {
        if (resolved) return;
        resolved = true;
        window.removeEventListener('focus', onFocus);
        URL.revokeObjectURL(blobUrl);
        iframe.src = 'about:blank';
        resolve();
      };

      const onFocus = () => {
        setTimeout(done, 500);
      };

      const onLoad = () => {
        iframe.removeEventListener('load', onLoad);

        try {
          iframe.contentWindow.addEventListener('afterprint', done);
        } catch {}

        window.addEventListener('focus', onFocus);
        iframe.contentWindow.focus();
        iframe.contentWindow.print();
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
    $('progress-modal').classList.add('hidden');
    $('print-btn').disabled = state.selectedIds.size === 0;
    $('download-btn').disabled = state.selectedIds.size === 0;
  }

  function updateProgress(current, total, label) {
    $('progress-title').textContent = 'Preparing print job...';
    $('progress-file').textContent = label;
    $('progress-count').textContent = total > 0 ? `${current} of ${total}` : '';
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
    setViewMode,
    startBatchPrint,
    startBatchDownload,
    cancelPrint,
    switchDrive,
    toggleRecursive,
    setFilter,
    sortBy,
  };
})();
