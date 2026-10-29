/**
 * Innomick Technologies - Document Management System (DMS) SPA Controller
 * Handles routing, view rendering, search, filtering, upload modals, 
 * approvals workflow actions, category drill-down, interactive context menus,
 * Version History, Manage Access, and Check Out / Check In states.
 */

// Application State
const AppState = {
  currentRoute: 'home',
  activeCategoryDrilldown: null,
  activeApprovalTab: 'inbox',
  activeWorkflowTemplateFilter: 'all',
  activeHistorySubTab: 'inbox',
  activeHomeTab: 'home',
  selectedRecycleBinItems: new Set(),
  openContextMenuDocId: null,
  columnFilters: {
    name: '',
    category: '',
    status: '',
    department: '',
    referenceNumber: '',
    favorite: '',
    stakeHolder: '',
    owner: '',
    subCategory: '',
    uploadDate: '',
    id: '',
    vendor: '',
    expireDate: '',
    description: ''
  },
  globalSearchQuery: '',
  activeBranch: 'Innomick Delhi',
  activeDepartment: 'Information Technology',
  departmentSearchQuery: '',
  hasUnreadInbox: true,
  activeApprovalDocId: null,
  pendingRemarksAction: 'Approve',
  pendingRemarksDocId: '1084/2042',
  homeRecentViewAll: false,
  homeInboxViewAll: false,
  homeAllFilesViewAll: false,
  favoritesPage: 1,
  favoritesPageSize: 10,
  activeDeptFolder: null,
  isDirectUploadMode: false,
  directUploadSelectedFolder: null,
  directUploadTempFolder: null,
  directUploadFilterQuery: ''
};

// Workflow Templates LocalStorage Persistence Key
const WORKFLOW_TEMPLATES_STORAGE_KEY = 'innomick_dms_workflow_templates';

function loadPersistedWorkflowTemplates() {
  try {
    const raw = localStorage.getItem(WORKFLOW_TEMPLATES_STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) {
        InnomickData.workflowTemplates = parsed;
      }
    }
  } catch (err) {
    console.warn('Failed to load persisted workflow templates:', err);
  }
}

function persistWorkflowTemplates() {
  try {
    if (InnomickData && Array.isArray(InnomickData.workflowTemplates)) {
      localStorage.setItem(WORKFLOW_TEMPLATES_STORAGE_KEY, JSON.stringify(InnomickData.workflowTemplates));
    }
  } catch (err) {
    console.warn('Failed to persist workflow templates:', err);
  }
}

// =============================================================================
// APPROVAL ENGINE
// Seed approvals belong to the persona the demo data was captured for; anything
// created at runtime is routed (routed: true) to the approver(s) of its active step.
// One record object is shared by the Inbox and Outbox lists so both stay in sync.
// =============================================================================
const SEED_APPROVAL_OWNER = (typeof InnomickData !== 'undefined' && InnomickData.currentUser && InnomickData.currentUser.name) || 'Mohammed Wasifuddin';
const STEP_APPROVER_POOL = ['Nelluri Venkatarao', 'MD Faisal Ahmed', 'Sayed Asif Ahmed', 'Abdul Jabbar', 'Mohammed Zuberuddin'];

// Group members lookup mapping for authentic SharePoint workflow groups
const GROUP_MEMBERS_MAP = {
  'Test-GM3': ['Abdul Jabbar', 'Sayed Asif Ahmed', 'Mohammed Wasifuddin'],
  'Test-GM5': ['MD Faisal Ahmed', 'Nelluri Venkatarao', 'Imran Hussain', 'Mohammed Zuberuddin', 'Abdul Jabbar'],
  'CRM': ['Sayed Asif Ahmed', 'Neha Joshi'],
  'HSE Committee': ['MD Faisal Ahmed', 'Arjun Mehta', 'Abdul Jabbar'],
  'Hyderabad Staff': ['Mohammed Wasifuddin', 'Mohammed Zuberuddin', 'Imran Hussain'],
  'Finance_CanPrintReceipts': ['Nelluri Venkatarao', 'Sayed Asif Ahmed'],
  'SG Senior Engineers': ['Imran Hussain', 'Mohammed Zuberuddin', 'Neha Joshi'],
  'Compliance Meeting': ['Abdul Jabbar', 'MD Faisal Ahmed'],
  'Test-GMJ': ['Mohammed Wasifuddin'],
  'Enterprise Operations Group': ['Mohammed Wasifuddin', 'Sayed Asif Ahmed'],
  'Central Supply Group': ['Sayed Asif Ahmed', 'Abdul Jabbar'],
  'Bulk Processing Group': ['Mohammed Wasifuddin', 'Nelluri Venkatarao'],
  'Executive Board': ['Anand Krishnan', 'Arjun Mehta']
};

function isCurrentUserAdmin() {
  const u = InnomickData.currentUser;
  if (!u) return false;
  const role = (u.role || '').toLowerCase();
  const id = (u.id || '').toLowerCase();
  const name = (u.name || '').toLowerCase();
  return id === 'imran' || id === 'intranet' || role.includes('admin') || role.includes('owner') || name.includes('imran');
}

function currentUserName() {
  return (InnomickData.currentUser && InnomickData.currentUser.name) || SEED_APPROVAL_OWNER;
}

function formatApprovalTimestamp(d) {
  const dt = d || new Date();
  const pad = n => String(n).padStart(2, '0');
  let h = dt.getHours();
  const ampm = h >= 12 ? 'PM' : 'AM';
  h = h % 12 || 12;
  return `${pad(dt.getDate())}/${pad(dt.getMonth() + 1)}/${dt.getFullYear()} ${h}.${pad(dt.getMinutes())} ${ampm}`;
}

function formatShortDate(d) {
  const dt = d || new Date();
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${dt.getDate()} ${months[dt.getMonth()]} ${dt.getFullYear()}`;
}

function isGroupApproval(item) {
  return String(item.workflowType || '').toUpperCase() === 'GRP';
}

// Bring seed records into one consistent shape: one review per step, each with a status
function normalizeApprovalRecord(item, isSeedInbox) {
  if (!item || item._normalized) return item;
  item._normalized = true;
  if (!item.name && item.documentName) item.name = item.documentName;
  if (!Array.isArray(item.reviews)) item.reviews = [];
  if (!item.history) item.history = [];
  if (isGroupApproval(item)) {
    const groupName = item.group || 'Group';
    const members = (item.groupMembers && item.groupMembers.length)
      ? item.groupMembers
      : (GROUP_MEMBERS_MAP[groupName] || ['Abdul Jabbar', 'Sayed Asif Ahmed', 'Mohammed Wasifuddin']);
    item.groupMembers = members;
    item.progressTotal = members.length;
    if (!item.reviews || !item.reviews.length) {
      item.reviews = members.map(m => ({
        stepName: `Approval by ${m}`,
        approverFullName: m,
        status: 'pending',
        comments: 'No User comments to view'
      }));
    }
    return item;
  }
  const step = Number(item.progressStep) || 0;
  const total = Math.max(Number(item.progressTotal) || 1, item.reviews.length || 1);
  while (item.reviews.length < total) {
    const i = item.reviews.length;
    item.reviews.push({ approverFullName: STEP_APPROVER_POOL[i % STEP_APPROVER_POOL.length], comments: 'No User comments to view' });
  }
  item.reviews.forEach((rev, i) => {
    if (!rev.stepName) {
      rev.stepName = (rev.approverName && /^step\b/i.test(rev.approverName)) ? rev.approverName
        : (total > 1 ? `Step ${i + 1}` : (rev.approverName || `Approval by ${rev.approverFullName || SEED_APPROVAL_OWNER}`));
    }
    if (!['rejected', 'change-requested'].includes(rev.status)) {
      rev.status = i < step ? 'approved' : (i === step ? 'pending' : 'waiting');
    }
  });
  // A seed inbox record is waiting on the seed owner at its active step
  if (isSeedInbox && item.reviews[step] && !isFinalApprovalStatus(item.status)) {
    item.reviews[step].approverFullName = SEED_APPROVAL_OWNER;
    if (total === 1) item.reviews[step].stepName = `Approval by ${SEED_APPROVAL_OWNER}`;
  }
  item.progressStep = step;
  item.progressTotal = total;
  return item;
}

function isFinalApprovalStatus(status) {
  return ['approved', 'rejected', 'request for change', 'change requested'].includes(String(status || '').toLowerCase());
}

function getActiveStepIndex(item) {
  const reviews = item.reviews || [];
  const idx = reviews.findIndex(r => r.status !== 'approved');
  return idx === -1 ? Math.max(0, reviews.length - 1) : idx;
}

function getStepApprovers(item) {
  if (isGroupApproval(item)) {
    const groupName = item.group || 'Group';
    return (item.groupMembers && item.groupMembers.length)
      ? item.groupMembers
      : (GROUP_MEMBERS_MAP[groupName] || ['Abdul Jabbar', 'Sayed Asif Ahmed', 'Mohammed Wasifuddin']);
  }
  if (!item.routed) return [SEED_APPROVAL_OWNER];
  const rev = (item.reviews || [])[getActiveStepIndex(item)];
  return rev ? [rev.approverFullName] : [SEED_APPROVAL_OWNER];
}

function canCurrentUserAct(item) {
  return !isFinalApprovalStatus(item.status) && getStepApprovers(item).includes(currentUserName());
}

function getMyInbox() {
  const me = currentUserName();
  const isAdmin = isCurrentUserAdmin();
  return (InnomickData.approvalsInbox || []).filter(item => isAdmin || getStepApprovers(item).includes(me));
}

function getMyOutbox() {
  const me = currentUserName();
  const isAdmin = isCurrentUserAdmin();
  return (InnomickData.approvalsOutbox || []).filter(item => isAdmin || (item.requestedBy || SEED_APPROVAL_OWNER) === me);
}

function getMyInboxHistory() {
  const me = currentUserName();
  const isAdmin = isCurrentUserAdmin();
  return (InnomickData.approvalsHistory || []).filter(item => isAdmin || (item.routed ? (item.actedBy || []).includes(me) : me === SEED_APPROVAL_OWNER));
}

function getMyOutboxHistory() {
  const me = currentUserName();
  return (InnomickData.approvalsOutboxHistory || []).filter(item => (item.requestedBy || item.submittedBy || SEED_APPROVAL_OWNER) === me);
}

// Exact lookup across every approval list; returns the record and the list it lives in
function findApprovalRecord(docId) {
  if (!docId) return { item: null, source: null };
  const lists = [
    ['inbox', InnomickData.approvalsInbox], ['outbox', InnomickData.approvalsOutbox],
    ['history', InnomickData.approvalsHistory], ['outbox-history', InnomickData.approvalsOutboxHistory]
  ];
  for (const [source, list] of lists) {
    const item = (list || []).find(x => x.id === docId || (x.slug && x.slug === docId));
    if (item) return { item, source };
  }
  return { item: null, source: null };
}

function initApprovalEngine() {
  (InnomickData.approvalsInbox || []).forEach(i => normalizeApprovalRecord(i, true));
  ['approvalsOutbox', 'approvalsHistory', 'approvalsOutboxHistory'].forEach(k =>
    (InnomickData[k] || []).forEach(i => normalizeApprovalRecord(i, false))
  );
}

// Create one approval request for an uploaded document and route it to the first approver
function createApprovalRequest(doc, opts) {
  const isGroup = Boolean(opts.isGroup);
  let reviews;
  if (isGroup) {
    const groupName = opts.groupName || 'Group';
    const members = (opts.groupMembers && opts.groupMembers.length)
      ? opts.groupMembers
      : (GROUP_MEMBERS_MAP[groupName] || ['Abdul Jabbar', 'Sayed Asif Ahmed', 'Mohammed Wasifuddin']);
    reviews = members.map(m => ({
      stepName: `Approval by ${m}`,
      approverFullName: m,
      status: 'pending',
      comments: 'No User comments to view'
    }));
  } else {
    const steps = (opts.steps && opts.steps.length) ? opts.steps : [{ stepName: '', approver: opts.approver }];
    reviews = steps.map((st, idx) => ({
      stepName: st.stepName || (steps.length > 1 ? `Step ${idx + 1}` : `Approval by ${st.approver}`),
      approverFullName: st.approver,
      status: idx === 0 ? 'pending' : 'waiting',
      comments: 'No User comments to view'
    }));
  }

  const record = {
    id: 'app-' + Date.now(),
    routed: true,
    _normalized: true,
    docId: doc.id,
    sourceBranch: AppState.currentUploadSource === 'my-department' ? AppState.activeBranch : null,
    sourceDepartment: AppState.currentUploadSource === 'my-department' ? AppState.activeDepartment : null,
    name: doc.name,
    documentName: doc.name,
    type: doc.type,
    workflowType: isGroup ? 'GRP' : 'IND',
    category: doc.category || 'Accreditation Documents',
    department: doc.department || AppState.activeDepartment || 'Information Technology',
    status: isGroup ? 'Approval Pending' : reviews[0].stepName,
    progressStep: 0,
    progressTotal: reviews.length,
    requestedBy: currentUserName(),
    requestedOn: formatShortDate(),
    group: isGroup ? opts.groupName : '-',
    groupMembers: isGroup ? (opts.groupMembers || (GROUP_MEMBERS_MAP[opts.groupName] || [])) : [],
    owner: doc.owner || currentUserName(),
    tagSubCategory: doc.subCategory || '-',
    referenceNumber: doc.referenceNumber || '-',
    updateDate: formatShortDate(),
    description: doc.description || '-',
    stakeholder: doc.stakeHolder || '-',
    vendor: doc.vendor || '-',
    reviews,
    history: [],
    actedBy: []
  };

  InnomickData.approvalsInbox.unshift(record);
  if (!InnomickData.approvalsOutbox) InnomickData.approvalsOutbox = [];
  InnomickData.approvalsOutbox.unshift(record);
  AppState.hasUnreadInbox = true;
  return record;
}

// Mirror the approval outcome onto the uploaded document everywhere it is listed
function syncApprovalOutcomeToDocument(item, docStatus) {
  if (!item.docId) return;
  ['documents', 'myFiles', 'recentDocuments', 'allFiles'].forEach(k =>
    (InnomickData[k] || []).forEach(d => { if (d.id === item.docId) d.status = docStatus; })
  );
}

// Next free numeric document ID (SharePoint-style list item IDs never collide)
function nextDocumentId() {
  let max = 109;
  const scan = (list) => (list || []).forEach(d => {
    const n = parseInt(d && d.id, 10);
    if (String(n) === String(d && d.id) && n > max) max = n;
  });
  scan(InnomickData.documents);
  scan(InnomickData.myFiles);
  Object.values(InnomickData.departmentSites || {}).forEach(branch =>
    Object.values(branch || {}).forEach(scan)
  );
  if (AppState.lastIssuedDocId && AppState.lastIssuedDocId > max) max = AppState.lastIssuedDocId;
  AppState.lastIssuedDocId = max + 1;
  return String(max + 1);
}

// Document type SVG icons
function getFileIconSvg(type) {
  switch (type) {
    case 'docx':
    case 'doc':
      return `<svg class="file-icon" viewBox="0 0 24 24" fill="#2b579a">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6z"/>
        <path d="M14 2v6h6" fill="#1b3a6b"/>
        <text x="7" y="18" fill="white" font-size="8" font-weight="bold">W</text>
      </svg>`;
    case 'xlsx':
    case 'xls':
      return `<svg class="file-icon" viewBox="0 0 24 24" fill="#217346">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6z"/>
        <path d="M14 2v6h6" fill="#144d2d"/>
        <text x="8" y="18" fill="white" font-size="8" font-weight="bold">X</text>
      </svg>`;
    case 'pptx':
    case 'ppt':
      return `<svg class="file-icon" viewBox="0 0 24 24" fill="#d24726">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6z"/>
        <path d="M14 2v6h6" fill="#993118"/>
        <text x="8" y="18" fill="white" font-size="8" font-weight="bold">P</text>
      </svg>`;
    case 'pdf':
      return `<svg class="file-icon" viewBox="0 0 24 24" fill="#dc2626">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6z"/>
        <path d="M14 2v6h6" fill="#991b1b"/>
        <text x="7" y="17" fill="white" font-size="7" font-weight="bold">PDF</text>
      </svg>`;
    case 'folder':
      return `<svg class="file-icon" viewBox="0 0 24 24" fill="#f59e0b">
        <path d="M4 4h5l2 3h9a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2z"/>
      </svg>`;
    default:
      return `<svg class="file-icon" viewBox="0 0 24 24" fill="#0284c7">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6z"/>
        <path d="M14 2v6h6" fill="#0369a1"/>
        <line x1="8" y1="13" x2="16" y2="13" stroke="white" stroke-width="1.5"/>
        <line x1="8" y1="17" x2="14" y2="17" stroke="white" stroke-width="1.5"/>
      </svg>`;
  }
}

// Router Initialization
function initRouter() {
  if (window.location.pathname === '/' || window.location.pathname === '') {
    const curHash = window.location.hash || '#/';
    try {
      history.replaceState(null, '', '/Home.aspx' + curHash);
    } catch(e) {}
  }
  window.addEventListener('hashchange', handleRouteChange);
  window.addEventListener('click', (e) => {
    // Close context menu if clicked outside
    if (!e.target.closest('.row-actions-group')) {
      closeContextMenu();
    }
    // Close file open popover if clicked outside
    if (!e.target.closest('.file-link') && 
        !e.target.closest('.home-file-name-link') && 
        !e.target.closest('.category-file-link') && 
        !e.target.closest('.sp-file-open-popover')) {
      closeAllFilePopovers();
    }
    // Close modal if backdrop overlay clicked directly
    if (e.target.classList && e.target.classList.contains('modal-overlay')) {
      e.target.classList.remove('active');
    }
    // Close additional members dropdown if clicked outside
    if (!e.target.closest('#additionalMembersRow')) {
      closeAdditionalMembersDropdown();
    }
  });
  window.addEventListener('scroll', () => {
    closeAllFilePopovers();
  }, true);
  window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      document.querySelectorAll('.modal-overlay.active').forEach(m => m.classList.remove('active'));
      closeContextMenu();
      closeAllFilePopovers();
      closeAdditionalMembersDropdown();
    }
  });
  handleRouteChange();
  updateApprovalsLeftBadges();
}

function handleRouteChange() {
  let hash = window.location.hash.replace(/^#\/?/, '');
  
  // Normalize if route includes Home/ or home/
  if (hash.toLowerCase().startsWith('home/')) {
    hash = hash.substring(5);
  }

  // Normalize navigation aliases
  if (hash === 'files') hash = 'my-files';
  if (hash === 'department') hash = 'my-department';
  if (hash === 'org' || hash === 'my-organization') hash = 'organization';
  if (hash === 'recycle') hash = 'recycle-bin';
  if (hash === 'approval' || hash === 'my-approvals' || hash.startsWith('approval?')) hash = 'approvals';

  if (hash.startsWith('view-approval') || hash.startsWith('approve')) {
    const isViewApproval = hash.startsWith('view-approval');
    AppState.currentRoute = isViewApproval ? 'view-approval' : 'approve-detail';
    const docId = hash.includes('/') ? decodeURIComponent(hash.split('/').slice(1).join('/')) : AppState.activeApprovalDocId;
    // Keep the tab the record was opened from; a direct link falls back to where the record lives
    if (docId !== AppState.activeApprovalDocId || !AppState.approvalDocSource) {
      const found = findApprovalRecord(docId);
      AppState.approvalDocSource = found.source || (isViewApproval ? 'outbox' : 'inbox');
    }
    AppState.activeApprovalDocId = docId;
    AppState.activeCategoryDrilldown = null;
  } else if (hash.startsWith('view-document')) {
    const queryStr = hash.includes('?') ? hash.substring(hash.indexOf('?') + 1) : '';
    const params = new URLSearchParams(queryStr);
    AppState.currentRoute = 'view-document';
    AppState.viewDocFileUrl = params.get('fileUrl') || '';
    AppState.viewDocBackUrl = params.get('back') || '';
    AppState.viewDocId = params.get('id') || '';
    AppState.activeCategoryDrilldown = null;
  } else if (hash.startsWith('department') || hash.startsWith('my-department')) {
    AppState.currentRoute = 'department';
    let branch = 'Innomick Delhi';
    let dept = 'Information Technology';
    let folder = null;
    
    if (hash.includes('?')) {
      const queryStr = hash.substring(hash.indexOf('?') + 1);
      const params = new URLSearchParams(queryStr);
      const rawBranch = params.get('branch');
      const rawDept = params.get('dept');
      const rawFolder = params.get('folder');
      if (rawBranch) branch = decodeParam(rawBranch);
      if (rawDept) dept = decodeParam(rawDept);
      if (rawFolder) folder = decodeParam(rawFolder);
    }
    
    AppState.activeBranch = branch;
    AppState.activeDepartment = dept;
    AppState.activeDeptFolder = folder;
    AppState.activeCategoryDrilldown = null;
  } else if (hash.startsWith('category/') || hash.startsWith('categories?cat=')) {
    let rawCat = '';
    if (hash.startsWith('category/')) {
      rawCat = hash.substring('category/'.length);
    } else {
      rawCat = hash.split('cat=')[1].split('&')[0];
    }
    let decodedCat = decodeURIComponent(rawCat);
    try {
      const b64 = atob(decodedCat);
      if (b64 && /^[\x20-\x7E]+$/.test(b64)) {
        decodedCat = b64;
      }
    } catch(e) {}
    AppState.currentRoute = 'category-drilldown';
    AppState.activeCategoryDrilldown = decodedCat;
  } else {
    AppState.currentRoute = hash || 'home';
    AppState.activeCategoryDrilldown = null;
  }

  // A different page (or department/folder) starts with clean filters, search and paging
  const viewKey = [AppState.currentRoute, AppState.activeBranch, AppState.activeDepartment, AppState.activeDeptFolder, AppState.activeCategoryDrilldown].join('|');
  if (viewKey !== AppState.lastViewKey) {
    AppState.lastViewKey = viewKey;
    resetListViewState();
  }

  updateSidebarActiveState();
  renderCurrentView();
}

function resetListViewState() {
  Object.keys(AppState.columnFilters).forEach(k => { AppState.columnFilters[k] = ''; });
  AppState.globalSearchQuery = '';
  AppState.departmentSearchQuery = '';
  AppState.listSortKey = null;
  AppState.listPage = 1;
  AppState.openContextMenuDocId = null;
  window._lastFilterFocus = null;
  window._lastSearchFocus = null;
}

function decodeParam(str) {
  if (!str) return '';
  let dec = decodeURIComponent(str);
  try {
    const b64 = atob(dec);
    if (b64 && /^[\x20-\x7E]+$/.test(b64)) {
      return b64;
    }
  } catch(e) {}
  if (dec.includes('Hyderabad') || dec.startsWith('SW5ub21pY2sgQWwtQWht')) return 'Innomick Hyderabad';
  if (dec.includes('Delhi') || dec.startsWith('SW5ub21pY2sgQWwtQXNz')) return 'Innomick Delhi';
  if (dec.startsWith('SG9zcGl0YWwgRGly')) return 'Executive Director Office';
  if (dec.startsWith('SW5mb3JtYXRpb24g')) return 'Information Technology';
  if (dec.startsWith('RGlyZWN0IFVwbG9h')) return 'Direct Uploads';
  if (dec.startsWith('VGVzdGluZw')) return 'Testing';
  if (dec.startsWith('T2ZmaWNlIEZpbGVz')) return 'Office Files';
  return dec;
}

function navigateTo(route) {
  let targetHash = `#/${route}`;
  if (route === 'approvals' || route === 'approval') {
    targetHash = '#/approval';
  } else if (route === 'home') {
    targetHash = '#/';
  }
  if (!window.location.pathname.toLowerCase().endsWith('home.aspx')) {
    try {
      history.pushState(null, '', '/Home.aspx' + targetHash);
    } catch(e) {}
  }
  window.location.hash = targetHash;
}

function navigateToDepartment(branch, dept, e) {
  if (e) {
    e.preventDefault();
    e.stopPropagation();
  }
  AppState.activeBranch = branch;
  AppState.activeDepartment = dept;
  const b64Branch = btoa(branch);
  const b64Dept = btoa(dept);
  window.location.hash = `#/department?branch=${b64Branch}&dept=${b64Dept}`;
}

function toggleMyDepartmentsTree(e) {
  if (e) {
    e.preventDefault();
    e.stopPropagation();
  }
  const subMenu = document.getElementById('subDepartmentsMenu');
  const chevron = document.getElementById('chevronMyDepartments');
  if (!subMenu) return;
  const isHidden = subMenu.style.display === 'none';
  subMenu.style.display = isHidden ? 'block' : 'none';
  if (chevron) {
    chevron.style.transform = isHidden ? 'rotate(180deg)' : 'rotate(0deg)';
  }
}

function toggleBranchHyderabad(e) {
  if (e) {
    e.preventDefault();
    e.stopPropagation();
  }
  const sub = document.getElementById('subHyderabadDepts');
  const chevron = document.getElementById('chevronHyderabad');
  if (sub) {
    const isHidden = sub.style.display === 'none';
    sub.style.display = isHidden ? 'block' : 'none';
    if (chevron) {
      chevron.style.transform = isHidden ? 'rotate(180deg)' : 'rotate(0deg)';
    }
  }
  navigateToDepartment('Innomick Hyderabad', 'Executive Director Office');
}

function toggleBranchDelhi(e) {
  if (e) {
    e.preventDefault();
    e.stopPropagation();
  }
  const sub = document.getElementById('subDelhiDepts');
  const chevron = document.getElementById('chevronDelhi');
  if (sub) {
    const isHidden = sub.style.display === 'none';
    sub.style.display = isHidden ? 'block' : 'none';
    if (chevron) {
      chevron.style.transform = isHidden ? 'rotate(180deg)' : 'rotate(0deg)';
    }
  }
  navigateToDepartment('Innomick Delhi', 'Information Technology');
}

function updateApprovalsLeftBadges() {
  let count = 85;
  try {
    if (typeof InnomickData !== 'undefined' && InnomickData.approvalsInbox) {
      count = getMyInbox().length;
    } else if (typeof window !== 'undefined' && window.InnomickData && window.InnomickData.approvalsInbox) {
      count = getMyInbox().length;
    }
  } catch(e) {}

  const sidebarBadge = document.getElementById('sidebarApprovalsBadge');
  if (sidebarBadge) {
    sidebarBadge.innerText = count;
    sidebarBadge.style.display = count > 0 ? 'inline-flex' : 'none';
  }
  document.querySelectorAll('.approvals-badge, .tab-badge, .inbox-count-badge').forEach(b => {
    b.innerText = count;
    b.style.display = count > 0 ? 'inline-flex' : 'none';
  });
  document.querySelectorAll('.inbox-red-dot').forEach(dot => {
    dot.style.display = count > 0 ? 'inline-block' : 'none';
  });
}

function updateSidebarActiveState() {
  document.querySelectorAll('.nav-item').forEach(el => {
    el.classList.remove('active');
  });

  if (AppState.currentRoute === 'department' || AppState.currentRoute === 'my-department') {
    const subMenu = document.getElementById('subDepartmentsMenu');
    if (subMenu) subMenu.style.display = 'block';

    const isHyderabad = (AppState.activeBranch && AppState.activeBranch.includes('Hyderabad')) || 
                     (AppState.activeDepartment === 'Executive Director Office');

    if (isHyderabad) {
      const subHyderabad = document.getElementById('subHyderabadDepts');
      if (subHyderabad) subHyderabad.style.display = 'block';
      const hdoItem = document.getElementById('navDeptHyderabadHDO');
      if (hdoItem) hdoItem.classList.add('active');
    } else {
      const subDelhi = document.getElementById('subDelhiDepts');
      if (subDelhi) subDelhi.style.display = 'block';
      const itItem = document.getElementById('navDeptDelhiIT');
      if (itItem) itItem.classList.add('active');
    }
  } else {
    document.querySelectorAll('.nav-item').forEach(el => {
      const target = el.getAttribute('data-route');
      if (target === AppState.currentRoute || 
         (AppState.currentRoute === 'approve-detail' && target === 'approvals') ||
         (AppState.currentRoute === 'view-approval' && target === 'approvals') ||
         (AppState.currentRoute === 'category-drilldown' && target === 'categories') ||
         (AppState.currentRoute === 'view-document' && (AppState.viewDocBackUrl.includes('categor') || target === 'categories'))) {
        el.classList.add('active');
      }
    });
  }
  updateApprovalsLeftBadges();
}

// Main View Dispatcher
function renderCurrentView() {
  const container = document.getElementById('mainContentContainer');
  if (!container) return;

  switch (AppState.currentRoute) {
    case 'home':
      container.innerHTML = renderHomeView();
      break;
    case 'approvals':
      container.innerHTML = renderMyApprovalsView();
      break;
    case 'approve-detail':
    case 'view-approval':
      container.innerHTML = renderApproveDocumentView(AppState.activeApprovalDocId);
      break;
    case 'favorites':
      container.innerHTML = renderMyFavoritesView();
      break;
    case 'my-files':
      container.innerHTML = renderMyFilesView();
      break;
    case 'my-department':
    case 'department':
      container.innerHTML = renderDepartmentView();
      break;
    case 'organization':
      container.innerHTML = renderMyOrganizationView();
      break;
    case 'shared':
      container.innerHTML = renderSharedFilesView();
      break;
    case 'categories':
      container.innerHTML = renderCategoriesView();
      break;
    case 'category-drilldown':
      container.innerHTML = renderCategoryDrilldownView(AppState.activeCategoryDrilldown);
      break;
    case 'view-document':
      container.innerHTML = renderViewDocumentView(AppState.viewDocFileUrl, AppState.viewDocBackUrl, AppState.viewDocId);
      break;
    case 'expiring':
      container.innerHTML = renderExpiringFilesView();
      break;
    case 'recycle-bin':
      container.innerHTML = renderRecycleBinView();
      break;
    default:
      container.innerHTML = renderHomeView();
  }
  updateApprovalsLeftBadges();

  // Restore focus synchronously so no keystroke typed into a filter/search box is lost
  const focusInfo = window._lastFilterFocus || window._lastSearchFocus;
  if (focusInfo) {
    const targetInput = window._lastFilterFocus
      ? container.querySelector(`input.column-search-input[data-filter-col="${focusInfo.col}"]`)
      : document.getElementById(focusInfo.id);
    window._lastFilterFocus = null;
    window._lastSearchFocus = null;
    if (targetInput) {
      targetInput.focus();
      if (typeof focusInfo.cursor === 'number' && targetInput.setSelectionRange) {
        targetInput.setSelectionRange(focusInfo.cursor, focusInfo.cursor);
      }
    }
  }
}

// =============================================================================
// 1. HOME VIEW
// =============================================================================
function renderHomeView() {
  if (AppState.activeHomeTab === 'approvals') {
    return renderHomeApprovalsTab();
  }

  const previewCategories = InnomickData.categories.slice(0, 12);

  let filteredRecent = InnomickData.documents;
  if (AppState.globalSearchQuery) {
    const q = AppState.globalSearchQuery.toLowerCase();
    filteredRecent = filteredRecent.filter(d => 
      d.name.toLowerCase().includes(q) || 
      d.category.toLowerCase().includes(q) ||
      d.department.toLowerCase().includes(q)
    );
  }

  return `
    <div class="content-header-row">
      <h1 class="view-main-title">Welcome to DMS</h1>
      <button class="btn-primary-teal" onclick="triggerDeviceUpload(event, null)">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
          <polyline points="17 8 12 3 7 8"/>
          <line x1="12" y1="3" x2="12" y2="15"/>
        </svg>
        Upload
      </button>
    </div>

    <div class="secondary-tabs-bar">
      <div class="tabs-group">
        <button class="tab-pill ${AppState.activeHomeTab === 'home' ? 'active' : ''}" onclick="switchHomeTab('home')">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>
          </svg>
          Home
        </button>
        <button class="tab-pill ${AppState.activeHomeTab === 'approvals' ? 'active' : ''}" onclick="switchHomeTab('approvals')">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
            <polyline points="14 2 14 8 20 8"/>
            <path d="m9 15 2 2 4-4"/>
          </svg>
          Approvals
          <span class="tab-badge">${getMyInbox().length}</span>
        </button>
      </div>
      <div class="tabs-actions">
        <button class="icon-action-btn" title="Refresh" onclick="refreshData()">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/>
            <path d="M21 3v5h-5"/>
            <path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/>
            <path d="M8 16H3v5"/>
          </svg>
        </button>
      </div>
    </div>

    <!-- Document Categories Section -->
    <div class="section-title-row">
      <h2 class="section-heading">Document Categories</h2>
      <a class="section-link" onclick="navigateTo('categories')">View all</a>
    </div>

    <div class="categories-grid">
      ${previewCategories.map(cat => `
        <div class="category-card" onclick="navigateTo('category/${cat.id}')">
          <svg class="folder-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#008272" stroke-width="1.8">
            <path d="M4 20h16a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.93a2 2 0 0 1-1.66-.9l-.82-1.2A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13c0 1.1.9 2 2 2Z"/>
          </svg>
          <span class="category-card-text" title="${cat.name}">${cat.name}</span>
        </div>
      `).join('')}
    </div>

    <!-- Recent Documents Section (13 Columns: Favorite, Name, Last Accessed, Location, Owner, Category, Sub Category, Department, Reference Number, Stake Holder, Vendor, Upload Date, Description) -->
    <div class="home-section-title-row">
      <h2>Recent Documents</h2>
      <div style="display: flex; align-items: center; gap: 14px;">
        <a class="home-view-all-link" onclick="toggleHomeRecentViewAll(event)" style="cursor: pointer;">
          ${AppState.homeRecentViewAll ? 'Show less (15)' : `View all (${(InnomickData.recentDocuments || []).length})`}
        </a>
        <a class="home-view-all-link" onclick="navigateTo('organization')" title="Open Organization Documents Page" style="font-size: 12px; color: #64748b; font-weight: normal; cursor: pointer;">
          All files &rarr;
        </a>
      </div>
    </div>

    <div class="table-card-container" style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.03); overflow: hidden; margin-bottom: 24px;">
      <div style="display: flex; justify-content: flex-end; padding: 8px 14px 4px 14px;">
        <button class="info-circle-btn" title="View details & properties" style="width: 22px; height: 22px; border-radius: 50%; border: 1px solid #cbd5e1; background: #fff; color: #64748b; font-size: 11px; cursor: pointer; display: flex; align-items: center; justify-content: center;">ⓘ</button>
      </div>
      <div style="overflow-x: auto;">
        ${(InnomickData.recentDocuments || []).length === 0 ? `
          <div class="empty-state-card" style="padding: 48px 16px; text-align: center; background: transparent; border: none; box-shadow: none;">
            <p style="color: #008272; font-weight: 500; font-size: 14px; margin: 0;">No recent documents found</p>
          </div>
        ` : `
        <table class="dms-data-table" style="min-width: 100%; border-top: 1px solid #f1f5f9;">
          <thead>
            <tr class="header-row">
              <th style="width: 48px; text-align: center;">Favorite</th>
              <th style="min-width: 250px;">Name</th>
              <th style="min-width: 130px;">Last Accessed</th>
              <th style="min-width: 140px;">Location</th>
              <th style="min-width: 160px;">Owner</th>
              <th style="min-width: 170px;">Category</th>
              <th style="min-width: 130px;">Sub Category</th>
              <th style="min-width: 160px;">Department</th>
              <th style="min-width: 140px;">Reference Number</th>
              <th style="min-width: 130px;">Stake Holder</th>
              <th style="min-width: 120px;">Vendor</th>
              <th style="min-width: 120px;">Upload Date</th>
              <th style="min-width: 220px;">Description</th>
            </tr>
          </thead>
          <tbody>
            ${(AppState.homeRecentViewAll ? (InnomickData.recentDocuments || []) : (InnomickData.recentDocuments || []).slice(0, 15)).map(doc => {
              let deptBadgeHtml = '<span style="color: #94a3b8;">-</span>';
              if (doc.department === 'Finance & Accounts') {
                deptBadgeHtml = `<span class="dept-pill-tag technical">${doc.department}</span>`;
              } else if (doc.department === 'Information Technology') {
                deptBadgeHtml = `<span class="dept-pill-tag it">${doc.department}</span>`;
              } else if (doc.department && doc.department !== '-') {
                deptBadgeHtml = `<span class="dept-pill-tag it">${doc.department}</span>`;
              }

              return `
                <tr>
                  <td style="text-align: center;">
                    <button class="favorite-star-btn ${doc.favorite ? 'starred' : ''}" onclick="toggleFavorite('${doc.id}')" title="${doc.favorite ? 'Unfavorite' : 'Add to Favorites'}">
                      <svg width="15" height="15" viewBox="0 0 24 24" fill="${doc.favorite ? 'currentColor' : 'none'}" stroke="currentColor" stroke-width="2">
                        <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>
                      </svg>
                    </button>
                  </td>
                  <td>
                    <div class="file-name-cell">
                      ${getFileIconSvg(doc.type)}
                      <span class="file-link" onclick="toggleDocumentMenu(event, '${doc.id}')">${doc.name}</span>
                      ${doc.isCheckedOut ? `<span title="Checked out by ${doc.checkedOutBy}" class="checked-out-lock"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg></span>` : ''}
                    </div>
                  </td>
                  <td style="color: #64748b;">${doc.lastAccessed || doc.uploadDate || '-'}</td>
                  <td style="color: #64748b;">${doc.location || 'My Organization'}</td>
                  <td>${doc.owner || '-'}</td>
                  <td>${doc.category || '-'}</td>
                  <td style="color: ${doc.subCategory === '-' ? '#94a3b8' : '#334155'};">${doc.subCategory || '-'}</td>
                  <td>${deptBadgeHtml}</td>
                  <td style="color: ${doc.referenceNumber === '-' ? '#94a3b8' : '#334155'};">${doc.referenceNumber || '-'}</td>
                  <td style="color: ${doc.stakeHolder === '-' ? '#94a3b8' : '#334155'};">${doc.stakeHolder || '-'}</td>
                  <td style="color: ${doc.vendor === '-' ? '#94a3b8' : '#334155'};">${doc.vendor || '-'}</td>
                  <td style="color: #64748b;">${doc.uploadDate || '-'}</td>
                  <td style="color: ${doc.description === '-' ? '#94a3b8' : '#334155'}; max-width: 240px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${doc.description || ''}">${doc.description || '-'}</td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>
        `}
      </div>
      ${(InnomickData.recentDocuments || []).length > 15 ? `
        <div class="home-table-pagination-bar">
          <span>Showing ${AppState.homeRecentViewAll ? InnomickData.recentDocuments.length : Math.min(15, InnomickData.recentDocuments.length)} of ${InnomickData.recentDocuments.length} documents</span>
          <button type="button" class="home-table-toggle-btn" onclick="toggleHomeRecentViewAll(event)">
            ${AppState.homeRecentViewAll ? 'Show less (15)' : `View all (${InnomickData.recentDocuments.length})`}
          </button>
        </div>
      ` : ''}
    </div>

    <!-- 4. Approval Workflow (Inbox) Section (8 Columns: Name, Category, Department, Status, Progress, Requested By, Requested on, Group) -->
    <div class="home-section-title-row">
      <h2>Approval Workflow (Inbox)</h2>
      <div style="display: flex; align-items: center; gap: 14px;">
        <a class="home-view-all-link" onclick="toggleHomeInboxViewAll(event)" style="cursor: pointer;">
          ${AppState.homeInboxViewAll ? 'Show less (15)' : `View all (${getMyInbox().length})`}
        </a>
        <a class="home-view-all-link" onclick="navigateTo('approvals')" title="Go to My Approvals Full Page" style="font-size: 12px; color: #64748b; font-weight: normal; cursor: pointer;">
          Approvals page &rarr;
        </a>
      </div>
    </div>

    <div class="home-inbox-card" style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.03); overflow: hidden; margin-bottom: 24px; padding: 0;">
      <div style="display: flex; justify-content: flex-end; padding: 8px 14px 4px 14px;">
        <button class="info-circle-btn" title="View details & properties" style="width: 22px; height: 22px; border-radius: 50%; border: 1px solid #cbd5e1; background: #fff; color: #64748b; font-size: 11px; cursor: pointer; display: flex; align-items: center; justify-content: center;">ⓘ</button>
      </div>
      <div style="overflow-x: auto;">
        ${getMyInbox().length === 0 ? `
          <div class="empty-state-card" style="padding: 48px 16px; text-align: center; background: transparent; border: none; box-shadow: none;">
            <p style="color: #008272; font-weight: 500; font-size: 14px; margin: 0;">You don't have any data in inbox</p>
          </div>
        ` : `
        <table class="dms-data-table approvals-table" style="min-width: 100%; border-top: 1px solid #f1f5f9;">
          <thead>
            <tr class="header-row">
              <th style="min-width: 260px;">Name</th>
              <th style="min-width: 170px;">Category</th>
              <th style="min-width: 180px;">Department</th>
              <th style="width: 130px;">Status</th>
              <th style="width: 130px;">Progress</th>
              <th style="width: 170px;">Requested By</th>
              <th style="width: 130px;">Requested on</th>
              <th style="width: 120px;">Group</th>
            </tr>
          </thead>
          <tbody>
            ${(AppState.homeInboxViewAll ? getMyInbox() : getMyInbox().slice(0, 15)).map(item => {
              const step = item.progressStep !== undefined ? item.progressStep : 0;
              const total = item.progressTotal !== undefined ? item.progressTotal : 1;
              const percent = total > 0 ? Math.round((step / total) * 100) : 0;
              const deptClass = getDeptPillClass(item.department);
              const reqBy = item.requestedBy || item.submittedBy || 'Mohammed Wasifuddin';
              const reqOn = item.requestedOn || item.submittedDate || '12 Oct 2026';
              const grpVal = item.group || (item.workflowType === 'GRP' ? 'Test-GM3' : '-');
              const docName = item.name || item.documentName || 'Document';
              const fileExt = (docName.split('.').pop() || 'docx').toLowerCase();

              return `
                <tr style="cursor: pointer;" onclick="openApprovalDocDetail('${item.id}', event, 'home')">
                  <td>
                    <div style="display: flex; align-items: center; gap: 8px;">
                      ${getFileIconSvg(fileExt)}
                      <a href="#/approve/${item.id}" onclick="openApprovalDocDetail('${item.id}', event, 'home')" class="sp-inbox-file-name" title="Open approval review: ${docName}">
                        ${docName}
                      </a>
                    </div>
                  </td>
                  <td><span style="font-size: 13px; color: #334155;">${item.category || 'Accreditation Documents'}</span></td>
                  <td>
                    <span class="sp-dept-pill ${deptClass}">${item.department || 'Information Technology'}</span>
                  </td>
                  <td><span style="font-size: 13px; color: #334155;">${item.status || 'Approval Pending'}</span></td>
                  <td>
                    <div class="sp-inbox-progress-container">
                      <div class="sp-inbox-progress-track">
                        <div class="sp-inbox-progress-fill" style="width: ${percent}%;"></div>
                      </div>
                      <div class="sp-inbox-progress-text">Step ${step} of ${total}</div>
                    </div>
                  </td>
                  <td><span style="font-size: 13px; color: #334155;">${reqBy}</span></td>
                  <td><span style="font-size: 13px; color: #64748b;">${reqOn}</span></td>
                  <td><span style="font-size: 13px; color: #334155;">${grpVal}</span></td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>
        `}
      </div>
      ${getMyInbox().length > 15 ? `
        <div class="home-table-pagination-bar">
          <span>Showing ${AppState.homeInboxViewAll ? getMyInbox().length : Math.min(15, getMyInbox().length)} of ${getMyInbox().length} approval requests</span>
          <button type="button" class="home-table-toggle-btn" onclick="toggleHomeInboxViewAll(event)">
            ${AppState.homeInboxViewAll ? 'Show less (15)' : `View all (${getMyInbox().length})`}
          </button>
        </div>
      ` : ''}
    </div>

    <!-- 5. All Files Section (13 Columns: Favorite, Name, Last Accessed, Location, Owner, Category, Sub Category, Department, Reference Number, Stake Holder, Vendor, Upload Date, Description) -->
    <div class="home-section-title-row">
      <h2>All Files</h2>
      <div style="display: flex; align-items: center; gap: 14px;">
        <a class="home-view-all-link" onclick="toggleHomeAllFilesViewAll(event)" style="cursor: pointer;">
          ${AppState.homeAllFilesViewAll ? 'Show less (15)' : `View all (${(InnomickData.homeAllFiles || []).length})`}
        </a>
        <a class="home-view-all-link" onclick="navigateTo('organization')" title="Open Organization Documents Page" style="font-size: 12px; color: #64748b; font-weight: normal; cursor: pointer;">
          Full list &rarr;
        </a>
      </div>
    </div>

    <div class="home-allfiles-card">
      <div class="home-allfiles-info-bar">
        <button class="info-circle-btn" title="View details & properties">ⓘ</button>
      </div>
      <div class="home-allfiles-table-wrapper">
        <table class="home-allfiles-table">
          <thead>
            <tr>
              <th style="width: 48px; text-align: center;">Favorite</th>
              <th style="min-width: 250px;">Name</th>
              <th style="min-width: 130px;">Last Accessed</th>
              <th style="min-width: 140px;">Location</th>
              <th style="min-width: 160px;">Owner</th>
              <th style="min-width: 170px;">Category</th>
              <th style="min-width: 130px;">Sub Category</th>
              <th style="min-width: 160px;">Department</th>
              <th style="min-width: 140px;">Reference Number</th>
              <th style="min-width: 130px;">Stake Holder</th>
              <th style="min-width: 120px;">Vendor</th>
              <th style="min-width: 120px;">Upload Date</th>
              <th style="min-width: 220px;">Description</th>
            </tr>
          </thead>
          <tbody>
            ${(AppState.homeAllFilesViewAll ? (InnomickData.homeAllFiles || []) : (InnomickData.homeAllFiles || []).slice(0, 15)).map(doc => {
              let deptBadgeHtml = '<span style="color: #94a3b8;">-</span>';
              if (doc.department === 'Finance & Accounts') {
                deptBadgeHtml = `<span class="dept-pill-tag technical">${doc.department}</span>`;
              } else if (doc.department === 'Information Technology') {
                deptBadgeHtml = `<span class="dept-pill-tag it">${doc.department}</span>`;
              } else if (doc.department && doc.department !== '-') {
                deptBadgeHtml = `<span class="dept-pill-tag it">${doc.department}</span>`;
              }

              return `
                <tr>
                  <td style="text-align: center;">
                    <button class="favorite-star-btn ${doc.favorite ? 'starred' : ''}" onclick="toggleHomeFileFavorite('${doc.id}')" title="Favorite">
                      <svg width="15" height="15" viewBox="0 0 24 24" fill="${doc.favorite ? 'currentColor' : 'none'}" stroke="currentColor" stroke-width="2">
                        <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>
                      </svg>
                    </button>
                  </td>
                  <td>
                    <div class="home-file-name-cell">
                      ${getFileIconSvg(doc.type)}
                      <span class="home-file-name-link" onclick="toggleDocumentMenu(event, '${doc.id}')">${doc.name}</span>
                    </div>
                  </td>
                  <td style="color: #64748b;">${doc.lastAccessed || doc.uploadDate || '-'}</td>
                  <td style="color: #64748b;">${doc.location || 'My Organization'}</td>
                  <td>${doc.owner || 'Mohammed Wasifuddin'}</td>
                  <td>${doc.category || '-'}</td>
                  <td style="color: ${doc.subCategory === '-' ? '#94a3b8' : '#334155'};">${doc.subCategory || '-'}</td>
                  <td>${deptBadgeHtml}</td>
                  <td style="color: ${doc.referenceNumber === '-' ? '#94a3b8' : '#334155'};">${doc.referenceNumber || '-'}</td>
                  <td style="color: ${doc.stakeHolder === '-' ? '#94a3b8' : '#334155'};">${doc.stakeHolder || '-'}</td>
                  <td style="color: ${doc.vendor === '-' ? '#94a3b8' : '#334155'};">${doc.vendor || '-'}</td>
                  <td style="color: #64748b;">${doc.uploadDate || '-'}</td>
                  <td style="color: ${doc.description === '-' ? '#94a3b8' : '#334155'}; max-width: 240px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${doc.description || ''}">${doc.description || '-'}</td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>
      </div>
      ${(InnomickData.homeAllFiles || []).length > 15 ? `
        <div class="home-table-pagination-bar">
          <span>Showing ${AppState.homeAllFilesViewAll ? InnomickData.homeAllFiles.length : Math.min(15, InnomickData.homeAllFiles.length)} of ${InnomickData.homeAllFiles.length} files</span>
          <button type="button" class="home-table-toggle-btn" onclick="toggleHomeAllFilesViewAll(event)">
            ${AppState.homeAllFilesViewAll ? 'Show less (15)' : `View all (${InnomickData.homeAllFiles.length})`}
          </button>
        </div>
      ` : ''}
    </div>
  `;
}


function toggleHomeRecentViewAll(e) {
  if (e) {
    e.preventDefault();
    e.stopPropagation();
  }
  AppState.homeRecentViewAll = !AppState.homeRecentViewAll;
  renderCurrentView();
}

function toggleHomeInboxViewAll(e) {
  if (e) {
    e.preventDefault();
    e.stopPropagation();
  }
  AppState.homeInboxViewAll = !AppState.homeInboxViewAll;
  renderCurrentView();
}

function toggleHomeAllFilesViewAll(e) {
  if (e) {
    e.preventDefault();
    e.stopPropagation();
  }
  AppState.homeAllFilesViewAll = !AppState.homeAllFilesViewAll;
  renderCurrentView();
}

function toggleHomeFileFavorite(id) {
  if (!InnomickData.homeAllFiles) return;
  const doc = InnomickData.homeAllFiles.find(d => d.id === id);
  if (doc) {
    doc.favorite = !doc.favorite;
    renderCurrentView();
    showToast(doc.favorite ? `"${doc.name}" added to Favorites` : `Removed from Favorites`, 'info');
  }
}

function switchHomeTab(tab) {
  AppState.activeHomeTab = tab;
  renderCurrentView();
}

function renderHomeApprovalsTab() {
  return `
    <div class="content-header-row">
      <h1 class="view-main-title">Welcome to DMS</h1>
    </div>

    <div class="secondary-tabs-bar" style="margin-bottom: 12px;">
      <div class="tabs-group">
        <button class="tab-pill" onclick="switchHomeTab('home')">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>
          </svg>
          Home
        </button>
        <button class="tab-pill active" onclick="switchHomeTab('approvals')">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
            <polyline points="14 2 14 8 20 8"/>
            <path d="m9 15 2 2 4-4"/>
          </svg>
          Approvals
          <span class="tab-badge">${getMyInbox().length}</span>
        </button>
      </div>
    </div>

    ${renderMyApprovalsInner(AppState.activeApprovalTab || 'inbox')}
  `;
}

// =============================================================================
// 2. MY APPROVALS VIEW
// =============================================================================
function renderMyApprovalsView() {
  return `
    <div class="content-header-row">
      <h1 class="view-main-title">My Approvals</h1>
    </div>

    ${renderMyApprovalsInner(AppState.activeApprovalTab || 'inbox')}
  `;
}

function renderMyApprovalsInner(activeTab) {
  const isTemplates = activeTab === 'templates';

  return `
    <div class="secondary-tabs-bar" style="margin-bottom: 20px;">
      <div class="tabs-group">
        <button class="tab-pill ${activeTab === 'inbox' ? 'active' : ''}" id="approvalTabInbox" onclick="switchApprovalTab('inbox')">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <polyline points="22 12 16 12 14 15 10 15 8 12 2 12"/>
            <path d="M5.45 5.11L2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z"/>
          </svg>
          Inbox
          ${AppState.hasUnreadInbox !== false ? `<span class="inbox-red-dot" title="New approval items"></span>` : ''}
        </button>
        <button class="tab-pill ${activeTab === 'outbox' ? 'active' : ''}" id="approvalTabOutbox" onclick="switchApprovalTab('outbox')">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <polyline points="22 12 16 12 14 9 10 9 8 12 2 12"/>
            <path d="M5.45 5.11L2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z"/>
          </svg>
          Outbox
        </button>
        <button class="tab-pill ${activeTab === 'templates' ? 'active' : ''}" id="approvalTabTemplates" onclick="switchApprovalTab('templates')">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <circle cx="12" cy="5" r="3"/>
            <line x1="12" y1="8" x2="12" y2="14"/>
            <circle cx="6" cy="19" r="3"/>
            <circle cx="18" cy="19" r="3"/>
            <line x1="6" y1="16" x2="12" y2="14"/>
            <line x1="18" y1="16" x2="12" y2="14"/>
          </svg>
          Workflow Templates
        </button>
        <button class="tab-pill ${activeTab === 'history' ? 'active' : ''}" id="approvalTabHistory" onclick="switchApprovalTab('history')">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <circle cx="12" cy="12" r="10"/>
            <polyline points="12 6 12 12 16 14"/>
          </svg>
          History
        </button>
      </div>
      <div class="tabs-actions" style="display: flex; align-items: center; gap: 8px;">
        ${isTemplates ? `
          <button class="icon-action-btn" title="Filter Workflow Templates" style="padding: 6px 10px;">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <line x1="4" y1="21" x2="4" y2="14"></line>
              <line x1="4" y1="10" x2="4" y2="3"></line>
              <line x1="12" y1="21" x2="12" y2="12"></line>
              <line x1="12" y1="8" x2="12" y2="3"></line>
              <line x1="20" y1="21" x2="20" y2="16"></line>
              <line x1="20" y1="12" x2="20" y2="3"></line>
              <line x1="1" y1="14" x2="7" y2="14"></line>
              <line x1="9" y1="8" x2="15" y2="8"></line>
              <line x1="17" y1="16" x2="23" y2="16"></line>
            </svg>
          </button>
        ` : `
          <button class="icon-action-btn" title="Refresh" onclick="refreshData()">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/>
              <path d="M21 3v5h-5"/>
              <path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/>
              <path d="M8 16H3v5"/>
            </svg>
          </button>
        `}
        <button class="btn-primary-teal" onclick="openCreateWorkflowModal()" style="display: flex; align-items: center; gap: 6px; padding: 6px 14px; font-size: 13px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <circle cx="12" cy="5" r="3"/>
            <line x1="12" y1="8" x2="12" y2="14"/>
            <circle cx="6" cy="19" r="3"/>
            <circle cx="18" cy="19" r="3"/>
            <line x1="6" y1="16" x2="12" y2="14"/>
            <line x1="18" y1="16" x2="12" y2="14"/>
          </svg>
          Create Workflow
        </button>
      </div>
    </div>

    ${renderApprovalTabContent(activeTab)}
  `;
}

function switchApprovalTab(tab) {
  AppState.activeApprovalTab = tab;
  renderCurrentView();
}

function switchHistorySubTab(subTab) {
  AppState.activeHistorySubTab = subTab;
  renderCurrentView();
}

function filterWorkflowTemplates(subtab) {
  AppState.activeWorkflowTemplateFilter = subtab;
  renderCurrentView();
}

function getDeptPillClass(dept) {
  if (!dept) return 'dept-pill-default';
  const d = dept.toLowerCase();
  if (d.includes('information technology') || d.includes('it') || d.includes('executive director') || d.includes('director office')) return 'dept-pill-it';
  if (d.includes('finance & accounts')) return 'dept-pill-med';
  if (d.includes('customer relations') || d.includes('sahh')) return 'dept-pill-sahh';
  return 'dept-pill-default';
}

function getHistoryProgressFillColor(status) {
  if (!status) return '#008272';
  const s = status.toLowerCase();
  if (s.includes('approved')) return '#008272';
  if (s.includes('reject')) return '#ef4444';
  if (s.includes('change')) return '#d97706';
  return '#008272';
}

function openApprovalDocDetail(docId, event, source) {
  if (event) {
    event.preventDefault();
    event.stopPropagation();
  }
  const src = source || 'inbox';
  const isOutbox = src === 'outbox' || src === 'outbox-history';
  AppState.activeApprovalDocId = docId;
  AppState.approvalDocSource = src;
  AppState.currentRoute = isOutbox ? 'view-approval' : 'approve-detail';

  const targetHash = isOutbox ? `#/view-approval/${docId}` : `#/approve/${docId}`;
  if (!window.location.pathname.toLowerCase().endsWith('home.aspx')) {
    try {
      history.pushState(null, '', '/Home.aspx' + targetHash);
    } catch(e) {}
  }
  window.location.hash = targetHash;
  updateSidebarActiveState();
  renderCurrentView();
}

function backToApprovalsInbox() {
  const src = AppState.approvalDocSource || 'inbox';
  if (src === 'history' || src === 'outbox-history') {
    AppState.activeHistorySubTab = src === 'history' ? 'inbox' : 'outbox';
    backToApprovalsTab('history');
  } else {
    backToApprovalsTab(src);
  }
}

function backToApprovalsTab(tabName) {
  if (tabName === 'home') {
    navigateTo('home');
    return;
  }
  AppState.currentRoute = 'approvals';
  AppState.activeApprovalTab = tabName || 'inbox';
  const targetHash = '#/approval';
  if (!window.location.pathname.toLowerCase().endsWith('home.aspx')) {
    try {
      history.pushState(null, '', '/Home.aspx' + targetHash);
    } catch(e) {}
  }
  window.location.hash = targetHash;
  updateSidebarActiveState();
  renderCurrentView();
}

function renderWorkflowTypeBadge(item) {
  const wfType = String(item.workflowType || (item.group && item.group !== '-' ? 'GRP' : 'IND')).toUpperCase();
  return `<span class="${wfType === 'GRP' ? 'wf-badge-grp' : 'wf-badge-ind'}">${wfType}</span>`;
}

function renderApprovalProgressCell(item, barColor) {
  const step = Number(item.progressStep) || 0;
  const total = Number(item.progressTotal) || 1;
  const percent = total > 0 ? Math.round((step / total) * 100) : 0;
  return `
    <div class="sp-inbox-progress-container">
      <div class="sp-inbox-progress-track">
        <div class="sp-inbox-progress-fill" style="width: ${percent}%;${barColor ? ` background: ${barColor};` : ''}"></div>
      </div>
      <div class="sp-inbox-progress-text">Step ${step} of ${total}</div>
    </div>`;
}

function renderApprovalTabContent(tab) {
  switch (tab) {
    case 'inbox': {
      const myInbox = getMyInbox();
      if (myInbox.length === 0) {
        return `<div class="empty-state-card" style="padding: 72px 16px; text-align: center; background: transparent; border: none; box-shadow: none;">
          <p style="color: #008272; font-weight: 500; font-size: 15px; margin: 0;">You don't have any data in inbox</p>
        </div>`;
      }
      const inboxPageSize = AppState.approvalsInboxPageSize || 10;
      const totalInboxPages = Math.max(1, Math.ceil(myInbox.length / inboxPageSize));
      const inboxPage = Math.min(AppState.approvalsInboxPage || 1, totalInboxPages);
      const inboxPageItems = myInbox.slice((inboxPage - 1) * inboxPageSize, inboxPage * inboxPageSize);

      return `
        <div class="table-card-container" style="overflow-x: auto;">
          <table class="dms-data-table approvals-table" style="min-width: 100%;">
            <thead>
              <tr class="header-row">
                <th style="min-width: 260px;">Name</th>
                <th style="width: 130px; text-align: center;">Workflow Type</th>
                <th style="min-width: 170px;">Category</th>
                <th style="min-width: 180px;">Department</th>
                <th style="width: 130px;">Status</th>
                <th style="width: 130px;">Progress</th>
                <th style="width: 170px;">Requested By</th>
                <th style="width: 130px;">Requested on</th>
                <th style="width: 120px;">Group</th>
              </tr>
            </thead>
            <tbody>
              ${inboxPageItems.map(item => {
                const docName = item.name || item.documentName || 'Document';
                return `
                  <tr style="cursor: pointer;" onclick="openApprovalDocDetail('${item.id}', event, 'inbox')">
                    <td>
                      <a href="#/approve/${item.id}" onclick="openApprovalDocDetail('${item.id}', event, 'inbox')" class="sp-inbox-file-name" title="Open approval review: ${docName}">${docName}</a>
                    </td>
                    <td style="text-align: center;">${renderWorkflowTypeBadge(item)}</td>
                    <td><span style="font-size: 13px; color: #334155;">${item.category || '-'}</span></td>
                    <td><span class="sp-dept-pill ${getDeptPillClass(item.department)}">${item.department || '-'}</span></td>
                    <td><span style="font-size: 13px; color: #334155;">${item.status || 'Approval Pending'}</span></td>
                    <td>${renderApprovalProgressCell(item)}</td>
                    <td><span style="font-size: 13px; color: #334155;">${item.requestedBy || item.submittedBy || '-'}</span></td>
                    <td><span style="font-size: 13px; color: #64748b;">${item.requestedOn || item.submittedDate || '-'}</span></td>
                    <td><span style="font-size: 13px; color: #334155;">${item.group || '-'}</span></td>
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        </div>
        ${renderApprovalPagination('inbox', inboxPage, totalInboxPages, inboxPageSize)}
      `;
    }
    case 'outbox': {
      const myOutbox = getMyOutbox();
      if (myOutbox.length === 0) {
        return `
          <div class="empty-state-card" style="padding: 72px 16px; text-align: center; background: transparent; border: none; box-shadow: none;">
            <p style="color: #008272; font-weight: 500; font-size: 15px; margin: 0;">You don't have any data in outbox</p>
          </div>
        `;
      }
      const outboxPageSize = AppState.approvalsOutboxPageSize || 10;
      const totalOutboxPages = Math.max(1, Math.ceil(myOutbox.length / outboxPageSize));
      const outboxPage = Math.min(AppState.approvalsOutboxPage || 1, totalOutboxPages);
      const outboxPageItems = myOutbox.slice((outboxPage - 1) * outboxPageSize, outboxPage * outboxPageSize);

      return `
        <div class="table-card-container" style="overflow-x: auto;">
          <table class="dms-data-table approvals-table" style="min-width: 100%;">
            <thead>
              <tr class="header-row">
                <th style="min-width: 320px;">Name</th>
                <th style="width: 44px; text-align: center;"></th>
                <th style="width: 130px; text-align: center;">Workflow Type</th>
                <th style="min-width: 200px;">Category</th>
                <th style="width: 150px;">Status</th>
                <th style="width: 160px;">Progress</th>
                <th style="width: 110px; text-align: center;">Attachment</th>
                <th style="width: 130px;">Requested on</th>
              </tr>
            </thead>
            <tbody>
              ${outboxPageItems.map(item => {
                const docName = item.name || item.documentName || 'Document';
                return `
                  <tr style="cursor: pointer;" onclick="openApprovalDocDetail('${item.id}', event, 'outbox')">
                    <td>
                      <a href="#/view-approval/${item.id}" onclick="openApprovalDocDetail('${item.id}', event, 'outbox')" class="sp-inbox-file-name" style="color: #0f172a; text-decoration: none; font-size: 13.5px; font-weight: 400;" title="View approval: ${docName}">${docName}</a>
                    </td>
                    <td style="text-align: center;">
                      <a href="#/view-approval/${item.id}" onclick="openApprovalDocDetail('${item.id}', event, 'outbox')" class="sp-outbox-open-link" title="Open approval details" style="display: inline-flex; align-items: center; justify-content: center; color: #008272; padding: 4px; border-radius: 3px;">
                        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#008272" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                          <line x1="7" y1="17" x2="17" y2="7"/>
                          <polyline points="7 7 17 7 17 17"/>
                        </svg>
                      </a>
                    </td>
                    <td style="text-align: center;">${renderWorkflowTypeBadge(item)}</td>
                    <td><span style="font-size: 13px; color: #334155;">${item.category || '-'}</span></td>
                    <td><span style="font-size: 13px; color: #334155;">${item.status || '-'}</span></td>
                    <td>${renderApprovalProgressCell(item)}</td>
                    <td style="text-align: center;">
                      <a href="javascript:void(0)" onclick="event.stopPropagation(); previewDocument('${item.docId || item.id}')" title="Open attachment" style="color: #008272; display: inline-flex;">
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                          <path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48"/>
                        </svg>
                      </a>
                    </td>
                    <td><span style="font-size: 13px; color: #334155;">${item.requestedOn || item.submittedDate || '-'}</span></td>
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        </div>
        ${renderApprovalPagination('outbox', outboxPage, totalOutboxPages, outboxPageSize)}
      `;
    }
    case 'templates': {
      const filter = AppState.activeWorkflowTemplateFilter || 'all';
      const filteredTemplates = InnomickData.workflowTemplates.filter(t => {
        if (filter === 'group') return t.type === 'Group Workflow';
        if (filter === 'individual') return t.type === 'Individual Workflow';
        return true;
      });

      return `
        <div class="sp-wf-subtabs-bar">
          <button class="sp-wf-subtab-btn ${filter === 'all' ? 'active' : ''}" onclick="filterWorkflowTemplates('all')">All Workflow Templates</button>
          <button class="sp-wf-subtab-btn ${filter === 'group' ? 'active' : ''}" onclick="filterWorkflowTemplates('group')">Group Workflow Templates</button>
          <button class="sp-wf-subtab-btn ${filter === 'individual' ? 'active' : ''}" onclick="filterWorkflowTemplates('individual')">Individual Workflow Templates</button>
        </div>

        <div class="sp-wf-card-grid">
          ${filteredTemplates.map(t => `
            <div class="sp-wf-template-card" onclick="viewWorkflowTemplate('${t.id}')">
              <div class="sp-wf-card-icon-box">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
                  <circle cx="12" cy="5" r="2.5"/>
                  <line x1="12" y1="7.5" x2="12" y2="13"/>
                  <circle cx="6" cy="18" r="2.5"/>
                  <circle cx="18" cy="18" r="2.5"/>
                  <line x1="6.5" y1="15.5" x2="12" y2="13"/>
                  <line x1="17.5" y1="15.5" x2="12" y2="13"/>
                </svg>
              </div>
              <span class="sp-wf-card-title" title="${t.name}">${t.name}</span>
            </div>
          `).join('')}
        </div>
      `;
    }
    case 'history': {
      const subTab = AppState.activeHistorySubTab || 'inbox';
      const historyList = subTab === 'inbox' ? getMyInboxHistory() : getMyOutboxHistory();
      const historySource = subTab === 'inbox' ? 'history' : 'outbox-history';

      const historyPageSize = AppState.approvalsHistoryPageSize || 10;
      const totalHistoryPages = Math.max(1, Math.ceil(historyList.length / historyPageSize));
      const historyPage = Math.min(AppState.approvalsHistoryPage || 1, totalHistoryPages);
      const historyPageItems = historyList.slice((historyPage - 1) * historyPageSize, historyPage * historyPageSize);

      return `
        <div class="sp-history-subtabs-row" style="display: flex; gap: 24px; margin-bottom: 16px; border-bottom: 1px solid #e2e8f0; padding-bottom: 0;">
          <button class="sp-history-subtab-btn ${subTab === 'inbox' ? 'active' : ''}" onclick="switchHistorySubTab('inbox')" style="background: none; border: none; border-bottom: ${subTab === 'inbox' ? '2.5px solid #008272' : '2.5px solid transparent'}; color: ${subTab === 'inbox' ? '#008272' : '#64748b'}; font-size: 13.5px; font-weight: ${subTab === 'inbox' ? '600' : '500'}; padding: 8px 4px; cursor: pointer;">Inbox History</button>
          <button class="sp-history-subtab-btn ${subTab === 'outbox' ? 'active' : ''}" onclick="switchHistorySubTab('outbox')" style="background: none; border: none; border-bottom: ${subTab === 'outbox' ? '2.5px solid #008272' : '2.5px solid transparent'}; color: ${subTab === 'outbox' ? '#008272' : '#64748b'}; font-size: 13.5px; font-weight: ${subTab === 'outbox' ? '600' : '500'}; padding: 8px 4px; cursor: pointer;">Outbox History</button>
        </div>

        ${historyList.length === 0 ? `
          <div class="empty-state-card" style="padding: 48px 16px; text-align: center;">
            <p style="color: #008272; font-weight: 500; font-size: 15px;">You don't have any data in ${subTab} history.</p>
          </div>
        ` : `
          <div class="table-card-container" style="overflow-x: auto;">
            <table class="dms-data-table approvals-table" style="min-width: 100%;">
              <thead>
                <tr class="header-row">
                  <th style="min-width: 250px;">Name</th>
                  <th style="width: 120px; text-align: center;">Workflow Type</th>
                  <th style="min-width: 170px;">Category</th>
                  <th style="min-width: 180px;">Department</th>
                  <th style="width: 140px;">Status</th>
                  <th style="width: 130px;">Progress</th>
                  <th style="width: 160px;">Requested By</th>
                  <th style="width: 120px;">Requested on</th>
                  <th style="width: 100px;">Group</th>
                </tr>
              </thead>
              <tbody>
                ${historyPageItems.map(item => {
                  const docName = item.name || item.documentName || 'Document';
                  const statusVal = item.status || item.response || 'Approved';
                  return `
                    <tr style="cursor: pointer;" onclick="openApprovalDocDetail('${item.id}', event, '${historySource}')">
                      <td>
                        <a href="#/approve/${item.id}" onclick="openApprovalDocDetail('${item.id}', event, '${historySource}')" class="sp-inbox-file-name" title="Open approval: ${docName}">${docName}</a>
                      </td>
                      <td style="text-align: center;">${renderWorkflowTypeBadge(item)}</td>
                      <td><span style="font-size: 13px; color: #334155;">${item.category || '-'}</span></td>
                      <td><span class="sp-dept-pill ${getDeptPillClass(item.department)}">${item.department || '-'}</span></td>
                      <td><span style="font-size: 13px; color: #334155; font-weight: 500;">${statusVal}</span></td>
                      <td>${renderApprovalProgressCell(item, getHistoryProgressFillColor(statusVal))}</td>
                      <td><span style="font-size: 13px; color: #334155;">${item.requestedBy || item.submittedBy || item.submittedTo || '-'}</span></td>
                      <td><span style="font-size: 13px; color: #64748b;">${item.requestedOn || item.completedDate || '-'}</span></td>
                      <td><span style="font-size: 13px; color: #334155;">${item.group || '-'}</span></td>
                    </tr>
                  `;
                }).join('')}
              </tbody>
            </table>
          </div>
          ${renderApprovalPagination('history', historyPage, totalHistoryPages, historyPageSize)}
        `}
      `;
    }
  }
}

function renderApprovalPagination(tab, currentPage, totalPages, pageSize) {
  const isFirst = currentPage <= 1;
  const isLast = currentPage >= totalPages;

  return `
    <div class="sp-pagination-bar" style="display: flex; align-items: center; justify-content: space-between; padding: 14px 6px; font-size: 13px; color: #475569; border-top: 1px solid #e2e8f0; margin-top: 10px;">
      <div style="display: flex; align-items: center; gap: 14px;">
        <button type="button" class="sp-page-btn" onclick="goToApprovalPage('${tab}', 1)" ${isFirst ? 'disabled style="opacity: 0.35; cursor: not-allowed; background: none; border: none; color: #64748b; font-size: 13px; font-weight: 600;"' : 'style="cursor: pointer; background: none; border: none; color: #008272; font-size: 13px; font-weight: 600;"'}>&lt;&lt;</button>
        <button type="button" class="sp-page-btn" onclick="changeApprovalPage('${tab}', -1)" ${isFirst ? 'disabled style="opacity: 0.35; cursor: not-allowed; background: none; border: none; color: #64748b; font-size: 13px; font-weight: 600;"' : 'style="cursor: pointer; background: none; border: none; color: #008272; font-size: 13px; font-weight: 600;"'}>&lt;</button>
        <span style="font-size: 12.5px; color: #334155; font-weight: 500;">${currentPage} / ${totalPages} pages</span>
        <button type="button" class="sp-page-btn" onclick="changeApprovalPage('${tab}', 1)" ${isLast ? 'disabled style="opacity: 0.35; cursor: not-allowed; background: none; border: none; color: #64748b; font-size: 13px; font-weight: 600;"' : 'style="cursor: pointer; background: none; border: none; color: #008272; font-size: 13px; font-weight: 600;"'}>&gt;</button>
        <button type="button" class="sp-page-btn" onclick="goToApprovalPage('${tab}', ${totalPages})" ${isLast ? 'disabled style="opacity: 0.35; cursor: not-allowed; background: none; border: none; color: #64748b; font-size: 13px; font-weight: 600;"' : 'style="cursor: pointer; background: none; border: none; color: #008272; font-size: 13px; font-weight: 600;"'}>&gt;&gt;</button>
      </div>

      <div style="display: flex; align-items: center; gap: 8px;">
        <select class="sp-items-per-page-select" onchange="setApprovalPageSize('${tab}', this.value)" style="border: 1px solid #008272; border-radius: 4px; padding: 3px 8px; font-size: 12.5px; color: #008272; outline: none; background: #ffffff; cursor: pointer; font-weight: 600;">
          <option value="10" ${pageSize === 10 ? 'selected' : ''}>10</option>
          <option value="20" ${pageSize === 20 ? 'selected' : ''}>20</option>
          <option value="50" ${pageSize === 50 ? 'selected' : ''}>50</option>
          <option value="100" ${pageSize === 100 ? 'selected' : ''}>100</option>
        </select>
        <span style="font-size: 12.5px; color: #64748b;">Items per page</span>
      </div>
    </div>
  `;
}

function changeApprovalPage(tab, delta) {
  if (tab === 'inbox') {
    const totalPages = Math.max(1, Math.ceil(getMyInbox().length / (AppState.approvalsInboxPageSize || 10)));
    AppState.approvalsInboxPage = Math.max(1, Math.min(totalPages, (AppState.approvalsInboxPage || 1) + delta));
  } else if (tab === 'outbox') {
    const totalPages = Math.max(1, Math.ceil((InnomickData.approvalsOutbox || []).length / (AppState.approvalsOutboxPageSize || 10)));
    AppState.approvalsOutboxPage = Math.max(1, Math.min(totalPages, (AppState.approvalsOutboxPage || 1) + delta));
  } else if (tab === 'history') {
    const list = AppState.activeHistorySubTab === 'inbox' ? (InnomickData.approvalsHistory || []) : (InnomickData.approvalsOutboxHistory || []);
    const totalPages = Math.max(1, Math.ceil(list.length / (AppState.approvalsHistoryPageSize || 10)));
    AppState.approvalsHistoryPage = Math.max(1, Math.min(totalPages, (AppState.approvalsHistoryPage || 1) + delta));
  }
  renderCurrentView();
}

function goToApprovalPage(tab, targetPage) {
  if (tab === 'inbox') {
    AppState.approvalsInboxPage = targetPage;
  } else if (tab === 'outbox') {
    AppState.approvalsOutboxPage = targetPage;
  } else if (tab === 'history') {
    AppState.approvalsHistoryPage = targetPage;
  }
  renderCurrentView();
}

function setApprovalPageSize(tab, size) {
  const sz = parseInt(size, 10) || 10;
  if (tab === 'inbox') {
    AppState.approvalsInboxPageSize = sz;
    AppState.approvalsInboxPage = 1;
  } else if (tab === 'outbox') {
    AppState.approvalsOutboxPageSize = sz;
    AppState.approvalsOutboxPage = 1;
  } else if (tab === 'history') {
    AppState.approvalsHistoryPageSize = sz;
    AppState.approvalsHistoryPage = 1;
  }
  renderCurrentView();
}


// =============================================================================
// APPROVAL DOCUMENT DETAIL VIEW & REMARKS (MATCHING SCREENSHOTS 3 & 4)
// =============================================================================

function renderApproveDocumentView(docId) {
  const found = findApprovalRecord(docId);
  const item = found.item;

  if (!item) {
    return `
      <div class="sp-approval-page-header">
        <div class="sp-approval-page-title-left">
          <button class="sp-circle-back-btn" onclick="backToApprovalsInbox()" title="Back to My Approvals">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 18 9 12 15 6"></polyline></svg>
          </button>
          <h2 class="sp-approval-doc-title">Approval not found</h2>
        </div>
      </div>
      <div class="empty-state-card" style="padding: 48px 16px; text-align: center;">
        <p style="color: #008272; font-weight: 500; font-size: 15px;">This approval request is no longer available.</p>
      </div>`;
  }

  normalizeApprovalRecord(item);
  // Actions only appear when the record is opened from the Inbox and it is waiting on this user
  const source = AppState.approvalDocSource || found.source;
  const viewOnly = !(source === 'inbox' || source === 'home') || !canCurrentUserAct(item);
  const reviews = item.reviews || [];
  const activeIdx = getActiveStepIndex(item);

  const docName = item.name || item.documentName || 'Document';
  const ownerName = item.owner || item.requestedBy || item.submittedBy || '-';
  const deptName = item.department || '-';
  const dateStr = item.updateDate || item.requestedOn || item.submittedDate || '-';
  const fileExt = (item.type || docName.split('.').pop() || 'docx').toLowerCase();

  return `
    <div class="sp-approval-page-header">
      <div class="sp-approval-page-title-left">
        <button class="sp-circle-back-btn" onclick="backToApprovalsInbox()" title="Back to My Approvals">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
            <polyline points="15 18 9 12 15 6"></polyline>
          </svg>
        </button>
        <h2 class="sp-approval-doc-title">${docName}</h2>
      </div>
      <div class="sp-approval-page-actions-right">
        <button class="sp-btn-view-history" onclick="showApprovalHistoryModal('${item.id}')">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <circle cx="12" cy="5" r="2.5"/><line x1="12" y1="7.5" x2="12" y2="13"/>
            <circle cx="6" cy="18" r="2.5"/><circle cx="18" cy="18" r="2.5"/>
            <line x1="6.5" y1="15.5" x2="12" y2="13"/><line x1="17.5" y1="15.5" x2="12" y2="13"/>
          </svg>
          View History
        </button>
        ${!viewOnly ? `
          <button class="sp-btn-action-reject" onclick="openApprovalRemarksModal('Reject', '${item.id}')">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
              <line x1="18" y1="6" x2="6" y2="18"/>
              <line x1="6" y1="6" x2="18" y2="18"/>
            </svg>
            Reject
          </button>
          <button class="sp-btn-action-change" onclick="openApprovalRemarksModal('Request for change', '${item.id}')">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
              <line x1="18" y1="6" x2="6" y2="18"/>
              <line x1="6" y1="6" x2="18" y2="18"/>
            </svg>
            Request for change
          </button>
          <button class="sp-btn-action-approve" onclick="openApprovalRemarksModal('Approve', '${item.id}')">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
              <polyline points="20 6 9 17 4 12"/>
            </svg>
            Approve
          </button>
        ` : ''}
      </div>
    </div>

    <div class="sp-approval-detail-layout">
      <div class="sp-approval-left-card">
        <div class="sp-approval-left-banner">
          ${getFileIconSvg(fileExt)}
          <span class="sp-approval-left-banner-title" title="${docName}">${docName}</span>
        </div>

        <div class="sp-approval-meta-grid">
          <div class="sp-approval-meta-label">Owner</div>
          <div class="sp-approval-meta-value">${ownerName}</div>

          <div class="sp-approval-meta-label">Category</div>
          <div class="sp-approval-meta-value">${item.category || '-'}</div>

          <div class="sp-approval-meta-label">Tag/Sub category</div>
          <div class="sp-approval-meta-value">${item.tagSubCategory || '-'}</div>

          <div class="sp-approval-meta-label">Department</div>
          <div class="sp-approval-meta-value">
            <span class="sp-dept-badge-indigo">${deptName}</span>
          </div>

          <div class="sp-approval-meta-label">Reference number</div>
          <div class="sp-approval-meta-value">${item.referenceNumber || '-'}</div>

          <div class="sp-approval-meta-label">Update Date</div>
          <div class="sp-approval-meta-value">${dateStr}</div>

          <div class="sp-approval-meta-label">Description</div>
          <div class="sp-approval-meta-value">${item.description || '-'}</div>

          <div class="sp-approval-meta-label">Stakeholder</div>
          <div class="sp-approval-meta-value">${item.stakeholder || '-'}</div>

          <div class="sp-approval-meta-label">Vendor</div>
          <div class="sp-approval-meta-value">${item.vendor || '-'}</div>
        </div>

        <button type="button" class="sp-btn-preview-outline" onclick="previewDocument('${item.docId || item.id}')">
          Preview document
        </button>
      </div>

      <div class="sp-approval-right-area">
        <div class="sp-approval-white-card">
          <h3 class="sp-approval-card-heading">Details</h3>
          <div class="sp-approval-details-grid">
            <div>
              <div style="font-size: 12px; color: #64748b; margin-bottom: 4px;">Requested By</div>
              <div style="font-size: 13px; color: #1e293b; font-weight: 500;">${item.requestedBy || ownerName}</div>
            </div>
            <div>
              <div style="font-size: 12px; color: #64748b; margin-bottom: 4px;">Department</div>
              <div style="font-size: 13px; color: #1e293b; font-weight: 500;">${deptName}</div>
            </div>
            <div>
              <div style="font-size: 12px; color: #64748b; margin-bottom: 4px;">Submission Date</div>
              <div style="font-size: 13px; color: #1e293b; font-weight: 500;">${item.requestedOn || dateStr}</div>
            </div>
          </div>
        </div>

        <div class="sp-approval-white-card">
          <h3 class="sp-approval-card-heading">Approvers reviews</h3>
          <div>
            ${reviews.map((rev, idx) => renderApprovalReviewStep(item, rev, idx, idx === reviews.length - 1, idx === activeIdx)).join('')}
          </div>
        </div>
      </div>
    </div>
  `;
}

// One step of the "Approvers reviews" timeline: check = approved, dot = current, hollow = waiting
function renderApprovalReviewStep(item, rev, idx, isLast, isActive) {
  const status = rev.status || 'pending';
  const final = isFinalApprovalStatus(item.status);
  let node;
  if (status === 'approved') {
    node = `<div class="sp-review-node-done"><svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg></div>`;
  } else if (status === 'rejected' || status === 'change-requested') {
    const c = status === 'rejected' ? '#c53030' : '#ca8a04';
    node = `<div class="sp-review-node-done" style="background: ${c};"><svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="3.5" stroke-linecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></div>`;
  } else if (isActive && !final) {
    node = `<div class="sp-review-node-ring" style="border-color: #008272;"><div class="sp-review-node-inner" style="background: #008272;"></div></div>`;
  } else {
    node = `<div class="sp-review-node-ring" style="border-color: #94a3b8;"></div>`;
  }
  const approver = rev.approverFullName || rev.approver || '-';
  const title = rev.stepName || (isGroupApproval(item) ? `Approval by ${approver}` : (rev.approverName || `Step ${idx + 1}`));
  const comments = rev.comments || rev.comment || 'No User comments to view';
  return `
    <div class="sp-approval-review-row">
      <div class="sp-review-node-col">
        ${node}
        ${!isLast ? `<div class="sp-review-dashed-connector"></div>` : ''}
      </div>
      <div class="sp-review-approver-col" style="width: 240px; flex-shrink: 0; padding-top: 1px;">
        <div class="sp-review-approver-title" style="font-size: 13.5px; font-weight: 600; color: #0f172a; line-height: 1.3;">${title}</div>
        <div class="sp-review-approver-name" style="font-size: 12px; color: #64748b; margin-top: 2px;">(${approver})</div>
      </div>
      <div class="sp-review-comment-textarea-box" style="flex: 1; min-height: 72px; background: #f3f4f6; border: 1px solid #e5e7eb; border-radius: 2px; padding: 10px 12px; font-size: 13px; color: #64748b; position: relative; line-height: 1.5;">
        ${comments}
        <div class="corner-handle" style="position: absolute; bottom: 4px; right: 4px; width: 6px; height: 6px; border-right: 1.5px solid #94a3b8; border-bottom: 1.5px solid #94a3b8;"></div>
      </div>
    </div>
  `;
}

function openApprovalRemarksModal(actionType, docId) {
  AppState.pendingRemarksAction = actionType || 'Approve';
  AppState.pendingRemarksDocId = docId || AppState.activeApprovalDocId;

  const modal = document.getElementById('approvalRemarksModal');
  const titleEl = document.getElementById('approvalRemarksModalTitle');
  const labelEl = document.getElementById('approvalRemarksModalLabel');
  const textareaEl = document.getElementById('approvalRemarksTextarea');
  const submitBtn = document.getElementById('approvalRemarksSubmitBtn');

  if (titleEl) {
    titleEl.innerText = actionType === 'Approve' ? 'Approval' : actionType;
  }
  if (labelEl) {
    labelEl.innerText = actionType === 'Approve' ? 'Approval Remarks' : `${actionType} Remarks`;
  }
  if (textareaEl) {
    textareaEl.value = '';
    textareaEl.placeholder = '';
  }
  if (submitBtn) {
    submitBtn.innerText = 'Submit';
    submitBtn.style.background = '#008272';
  }

  if (modal) {
    modal.classList.add('active');
    modal.style.display = 'flex';
  }
}

function closeApprovalRemarksModal() {
  const modal = document.getElementById('approvalRemarksModal');
  if (modal) {
    modal.classList.remove('active');
    modal.style.display = 'none';
  }
}

// Publish the approved document to the department it was uploaded from
function publishApprovedDocumentToDepartment(item) {
  if (!item || !item.sourceBranch || !item.sourceDepartment) return;
  const branch = item.sourceBranch;
  const dept = item.sourceDepartment;

  if (!InnomickData.departmentSites) InnomickData.departmentSites = {};
  if (!InnomickData.departmentSites[branch]) InnomickData.departmentSites[branch] = {};
  if (!InnomickData.departmentSites[branch][dept]) InnomickData.departmentSites[branch][dept] = [];

  const original = (InnomickData.documents || []).find(d => d.id === item.docId) || {};
  const approvedDoc = {
    ...original,
    id: item.docId || nextDocumentId(),
    name: item.name || item.documentName || 'Document.docx',
    type: original.type || ((item.name || '').split('.').pop() || 'docx').toLowerCase(),
    category: item.category || '-',
    subCategory: item.tagSubCategory || original.subCategory || '-',
    department: dept,
    owner: item.owner || item.requestedBy,
    status: 'Active',
    stakeHolder: item.stakeholder || original.stakeHolder || '-',
    referenceNumber: item.referenceNumber || '-',
    vendor: item.vendor || '-',
    description: item.description || '-',
    folder: '-',
    parentFolder: null,
    location: `${dept}`,
    favorite: false
  };

  if (!InnomickData.departmentSites[branch][dept].some(d => d.id === approvedDoc.id)) {
    InnomickData.departmentSites[branch][dept].unshift(approvedDoc);
  }
}

// Close out a finished request: leave the Inbox/Outbox, land in both histories
function completeApprovalRecord(item, finalStatus) {
  item.status = finalStatus;
  const drop = (list) => { const i = (list || []).indexOf(item); if (i !== -1) list.splice(i, 1); };
  drop(InnomickData.approvalsInbox);
  drop(InnomickData.approvalsOutbox);
  if (!InnomickData.approvalsHistory) InnomickData.approvalsHistory = [];
  if (!InnomickData.approvalsHistory.includes(item)) InnomickData.approvalsHistory.unshift(item);
  if (!InnomickData.approvalsOutboxHistory) InnomickData.approvalsOutboxHistory = [];
  if (!InnomickData.approvalsOutboxHistory.includes(item)) InnomickData.approvalsOutboxHistory.unshift(item);
}

function submitApprovalRemarksAction() {
  const actionType = AppState.pendingRemarksAction || 'Approve';
  const docId = AppState.pendingRemarksDocId || AppState.activeApprovalDocId;
  const textareaEl = document.getElementById('approvalRemarksTextarea');
  const remarkText = textareaEl ? textareaEl.value.trim() : '';
  const { item } = findApprovalRecord(docId);

  if (!item || !canCurrentUserAct(item)) {
    closeApprovalRemarksModal();
    showToast('This approval is no longer waiting on you.', 'warning');
    backToApprovalsInbox();
    return;
  }
  if (actionType !== 'Approve' && !remarkText) {
    showToast(`Please enter remarks for "${actionType}".`, 'error');
    if (textareaEl) textareaEl.focus();
    return;
  }

  const me = currentUserName();
  const idx = getActiveStepIndex(item);
  let rev = item.reviews[idx];
  if (isGroupApproval(item)) {
    const memberRev = item.reviews.find(r => r.approverFullName === me && r.status !== 'approved');
    if (memberRev) rev = memberRev;
  }
  const actionLabel = actionType === 'Approve' ? 'Approved' : actionType === 'Reject' ? 'Rejected' : 'Request for Change';

  rev.comments = remarkText || 'No User comments to view';
  rev.actedBy = me;
  if (!item.actedBy) item.actedBy = [];
  if (!item.actedBy.includes(me)) item.actedBy.push(me);
  if (!item.history) item.history = [];
  item.history.push({ modifiedOn: formatApprovalTimestamp(), modifiedBy: me, action: actionLabel, comments: remarkText || '-' });

  if (actionType === 'Approve') {
    rev.status = 'approved';
    if (isGroupApproval(item)) {
      const allApproved = item.reviews.every(r => r.status === 'approved');
      item.progressStep = item.reviews.filter(r => r.status === 'approved').length;
      if (allApproved) {
        completeApprovalRecord(item, 'Approved');
        syncApprovalOutcomeToDocument(item, 'Active');
        publishApprovedDocumentToDepartment(item);
        addApprovedToRecentDocuments(item);
      } else {
        item.status = 'Approval Pending';
        if (!InnomickData.approvalsHistory.includes(item)) InnomickData.approvalsHistory.unshift(item);
      }
    } else {
      item.progressStep = idx + 1;
      const next = item.reviews[idx + 1];
      if (next) {
        // Hand the request to the next step's approver
        next.status = 'pending';
        item.status = next.stepName || `Step ${idx + 2}`;
        item.routed = true;
        if (!InnomickData.approvalsHistory.includes(item)) InnomickData.approvalsHistory.unshift(item);
      } else {
        completeApprovalRecord(item, 'Approved');
        syncApprovalOutcomeToDocument(item, 'Active');
        publishApprovedDocumentToDepartment(item);
        addApprovedToRecentDocuments(item);
      }
    }
  } else if (actionType === 'Reject') {
    rev.status = 'rejected';
    completeApprovalRecord(item, 'Rejected');
    syncApprovalOutcomeToDocument(item, 'Rejected');
  } else {
    rev.status = 'change-requested';
    completeApprovalRecord(item, 'Request for Change');
    syncApprovalOutcomeToDocument(item, 'Request for Change');
  }

  closeApprovalRemarksModal();
  updateApprovalsLeftBadges();

  if (AppState.approvalDocSource === 'home') {
    navigateTo('home');
  } else {
    backToApprovalsTab('inbox');
  }
  showInfoSuccessModal(actionType === 'Approve' ? 'Request is Approved!'
    : actionType === 'Reject' ? 'Request is Rejected!' : 'Change is Requested!');
}

// Approved documents appear at the top of Home > Recent Documents
function addApprovedToRecentDocuments(item) {
  if (!InnomickData.recentDocuments) InnomickData.recentDocuments = [];
  const original = (InnomickData.documents || []).find(d => d.id === item.docId) || {};
  const idx = InnomickData.recentDocuments.findIndex(d => d.id === (item.docId || item.id));
  if (idx !== -1) InnomickData.recentDocuments.splice(idx, 1);
  InnomickData.recentDocuments.unshift({
    ...original,
    id: item.docId || item.id,
    name: item.name,
    type: original.type || (item.name || '').split('.').pop(),
    status: 'Active',
    lastAccessed: '1 minutes ago',
    location: original.location || 'My Department',
    owner: item.owner || item.requestedBy,
    category: item.category,
    subCategory: item.tagSubCategory || '-',
    department: item.department
  });
}

// "View History": the approval actions taken on this request (Modified On / By / Action / Comments)
function showApprovalHistoryModal(docId) {
  const { item } = findApprovalRecord(docId);
  const rows = item ? [...(item.history || [])] : [];
  if (item && rows.length === 0) {
    (item.reviews || []).forEach(rev => {
      if (['approved', 'rejected', 'change-requested'].includes(rev.status)) {
        rows.push({
          modifiedOn: item.updateDate || item.requestedOn || '-',
          modifiedBy: rev.actedBy || rev.approverFullName || '-',
          action: rev.status === 'approved' ? 'Approved' : rev.status === 'rejected' ? 'Rejected' : 'Request for Change',
          comments: rev.comments && rev.comments !== 'No User comments to view' ? rev.comments : '-'
        });
      }
    });
  }

  let modal = document.getElementById('approvalHistoryModal');
  if (!modal) {
    modal = document.createElement('div');
    modal.id = 'approvalHistoryModal';
    modal.className = 'modal-overlay';
    document.body.appendChild(modal);
  }
  modal.innerHTML = `
    <div class="dms-modal-dialog sp-approval-history-dialog" role="dialog" aria-label="History">
      <div class="sp-approval-history-header">
        <span>History</span>
        <button type="button" class="sp-approval-history-close" onclick="closeApprovalHistoryModal()" aria-label="Close">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
        </button>
      </div>
      <div class="sp-approval-history-body">
        <table class="sp-approval-history-table">
          <thead><tr><th>Modified On</th><th>Modified By</th><th>Action</th><th>Comments</th></tr></thead>
          <tbody>
            ${rows.length ? rows.map(r => `<tr><td>${r.modifiedOn}</td><td>${r.modifiedBy}</td><td>${r.action}</td><td>${r.comments}</td></tr>`).join('')
              : `<tr><td colspan="4" class="sp-approval-history-empty">No approval actions have been taken yet.</td></tr>`}
          </tbody>
        </table>
      </div>
    </div>`;
  modal.classList.add('active');
}

function closeApprovalHistoryModal() {
  const modal = document.getElementById('approvalHistoryModal');
  if (modal) modal.classList.remove('active');
}

// =============================================================================
// SHAREPOINT 13-COLUMN DATA TABLE ENGINE
// 1. Name | 2. Category | 3. Status | 4. Department | 5. ReferenceNumber |
// 6. Favorite | 7. StakeHolder | 8. Owner Name | 9. Sub Category |
// 10. Upload Date | 11. ID | 12. Vendor | 13. Expire Date
// =============================================================================
function filterDocumentsByAllColumns(docsList) {
  let docs = [...docsList];
  if (AppState.columnFilters.name) {
    docs = docs.filter(d => (d.name || '').toLowerCase().includes(AppState.columnFilters.name.toLowerCase()));
  }
  if (AppState.columnFilters.category) {
    docs = docs.filter(d => (d.category || '').toLowerCase().includes(AppState.columnFilters.category.toLowerCase()));
  }
  if (AppState.columnFilters.status) {
    docs = docs.filter(d => (d.status || '').toLowerCase().includes(AppState.columnFilters.status.toLowerCase()));
  }
  if (AppState.columnFilters.department) {
    docs = docs.filter(d => (d.department || '').toLowerCase().includes(AppState.columnFilters.department.toLowerCase()));
  }
  if (AppState.columnFilters.referenceNumber) {
    docs = docs.filter(d => (d.referenceNumber || '').toLowerCase().includes(AppState.columnFilters.referenceNumber.toLowerCase()));
  }
  if (AppState.columnFilters.favorite) {
    const q = AppState.columnFilters.favorite.toLowerCase();
    docs = docs.filter(d => {
      const favStr = d.favorite ? 'yes' : 'no';
      return favStr.includes(q);
    });
  }
  if (AppState.columnFilters.stakeHolder) {
    docs = docs.filter(d => (d.stakeHolder || '').toLowerCase().includes(AppState.columnFilters.stakeHolder.toLowerCase()));
  }
  if (AppState.columnFilters.owner) {
    docs = docs.filter(d => (d.owner || '').toLowerCase().includes(AppState.columnFilters.owner.toLowerCase()));
  }
  if (AppState.columnFilters.subCategory) {
    docs = docs.filter(d => (d.subCategory || '').toLowerCase().includes(AppState.columnFilters.subCategory.toLowerCase()));
  }
  if (AppState.columnFilters.uploadDate) {
    docs = docs.filter(d => (d.uploadDate || '').toLowerCase().includes(AppState.columnFilters.uploadDate.toLowerCase()));
  }
  if (AppState.columnFilters.id) {
    docs = docs.filter(d => (d.id || '').toLowerCase().includes(AppState.columnFilters.id.toLowerCase()));
  }
  if (AppState.columnFilters.vendor) {
    docs = docs.filter(d => (d.vendor || '').toLowerCase().includes(AppState.columnFilters.vendor.toLowerCase()));
  }
  if (AppState.columnFilters.expireDate) {
    docs = docs.filter(d => (d.expireDate || '').toLowerCase().includes(AppState.columnFilters.expireDate.toLowerCase()));
  }
  return docs;
}

// =============================================================================
// 3. MY FAVORITES VIEW (13 Columns Matching SharePoint)
// =============================================================================
function getAllFavoriteDocuments() {
  const seenIds = new Set();
  const favs = [];

  const addIfFav = (d) => {
    if (d && d.favorite && !d.isFolder && !seenIds.has(String(d.id))) {
      seenIds.add(String(d.id));
      favs.push(d);
    }
  };

  // 1. Department documents (Executive Director Office, IT, etc.)
  if (InnomickData.departmentSites) {
    for (const b in InnomickData.departmentSites) {
      for (const d in InnomickData.departmentSites[b]) {
        (InnomickData.departmentSites[b][d] || []).forEach(addIfFav);
      }
    }
  }

  // 2. Main documents collection
  (InnomickData.documents || []).forEach(addIfFav);

  // 3. Other collections
  (InnomickData.recentDocuments || []).forEach(addIfFav);
  (InnomickData.homeAllFiles || []).forEach(addIfFav);
  (InnomickData.allFiles || []).forEach(addIfFav);
  (InnomickData.myFiles || []).forEach(addIfFav);
  (InnomickData.sharedFiles || []).forEach(addIfFav);
  (InnomickData.expiringFiles || []).forEach(addIfFav);

  return favs;
}

function renderMyFavoritesView() {
  let favorites = getAllFavoriteDocuments();
  const q = (AppState.globalSearchQuery || '').trim().toLowerCase();
  if (q) {
    favorites = favorites.filter(d => [d.name, d.location, d.category, d.department, d.description, d.subCategory, d.stakeHolder]
      .some(v => String(v || '').toLowerCase().includes(q)));
  }

  const pageSize = AppState.favoritesPageSize || 10;
  const totalPages = Math.max(1, Math.ceil(favorites.length / pageSize));
  if ((AppState.favoritesPage || 1) > totalPages) AppState.favoritesPage = totalPages;
  const currentPage = AppState.favoritesPage || 1;
  const pageDocs = favorites.slice((currentPage - 1) * pageSize, currentPage * pageSize);
  const dash = v => (v && v !== '-' ? escapeHtmlDept(v) : '-');

  return `
    <div class="sp-favorites-header">
      <h1 class="view-main-title">My Favorites</h1>
      <div class="search-input-wrapper sp-favorites-search">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
        <input type="text" id="favoritesSearchInput" placeholder="Search Favorite Files" value="${escapeHtmlDept(AppState.globalSearchQuery || '')}" oninput="handleSearch(this.value, this)">
      </div>
      <button class="icon-action-btn" title="Refresh" onclick="refreshData()">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/><path d="M8 16H3v5"/></svg>
      </button>
    </div>

    <div class="sp-favorites-table-wrap">
      <table class="sp-favorites-table">
        <thead>
          <tr>
            <th style="width: 80px;">Favorite</th>
            <th style="min-width: 260px;">File Name</th>
            <th>Location</th>
            <th>Category</th>
            <th>Department</th>
            <th>Description</th>
            <th>Upload Date</th>
            <th>Tag/Sub Category</th>
            <th>Stake Holder</th>
            <th>Modified</th>
            <th>Activity</th>
          </tr>
        </thead>
        <tbody>
          ${pageDocs.length === 0 ? `
            <tr><td colspan="11" style="text-align: center; padding: 40px; color: #64748b;">No favorite documents found.</td></tr>
          ` : pageDocs.map(doc => `
            <tr>
              <td>
                <button class="favorite-star-btn" title="Remove from favorites" onclick="toggleDeptItemFavorite('${doc.id}', event)">
                  <svg width="17" height="17" viewBox="0 0 24 24" fill="#eab308" stroke="#ca8a04" stroke-width="1.5"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>
                </button>
              </td>
              <td>
                <div style="display: flex; align-items: center; gap: 8px;">
                  ${getFileIconSvg(doc.type)}
                  <span class="file-link" onclick="toggleDocumentMenu(event, '${doc.id}')" title="${escapeHtmlDept(doc.name)}">${escapeHtmlDept(doc.name)}</span>
                </div>
              </td>
              <td>${dash(doc.location && !String(doc.location).includes('>') ? doc.location : (doc.parentFolder || doc.isFolder ? 'My Department' : 'My Organization'))}</td>
              <td>${dash(doc.category)}</td>
              <td>${doc.department ? `<span class="sp-dept-pill ${getDeptPillClass(doc.department)}">${escapeHtmlDept(doc.department)}</span>` : '-'}</td>
              <td>${dash(doc.description)}</td>
              <td>${dash(doc.uploadDate)}</td>
              <td>${dash(doc.subCategory)}</td>
              <td>${dash(doc.stakeHolder)}</td>
              <td>${dash(doc.modified || doc.uploadDate)}</td>
              <td>${dash(doc.lastAccessed)}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    </div>
    ${favorites.length > pageSize ? `
      <div class="table-pagination-footer">
        <div class="pagination-controls">
          <button type="button" class="pagination-btn" ${currentPage <= 1 ? 'disabled' : 'onclick="goToFavoritesPage(1)"'}>&laquo;</button>
          <button type="button" class="pagination-btn" ${currentPage <= 1 ? 'disabled' : 'onclick="changeFavoritesPage(-1)"'}>&lsaquo;</button>
          <span>${currentPage} / ${totalPages} pages</span>
          <button type="button" class="pagination-btn" ${currentPage >= totalPages ? 'disabled' : 'onclick="changeFavoritesPage(1)"'}>&rsaquo;</button>
          <button type="button" class="pagination-btn" ${currentPage >= totalPages ? 'disabled' : `onclick="goToFavoritesPage(${totalPages})"`}>&raquo;</button>
        </div>
        <div style="display: flex; align-items: center; gap: 14px;">
          <span style="font-size: 12.5px; color: #64748b;">${favorites.length} items</span>
          <select class="pagination-select" onchange="setFavoritesPageSize(this.value)">
            ${[10, 20, 50].map(n => `<option value="${n}" ${pageSize === n ? 'selected' : ''}>${n}</option>`).join('')}
          </select>
          <span>Items per page</span>
        </div>
      </div>
    ` : ''}
  `;
}

function changeFavoritesPage(delta) {
  const totalPages = Math.max(1, Math.ceil(getAllFavoriteDocuments().length / (AppState.favoritesPageSize || 10)));
  AppState.favoritesPage = Math.max(1, Math.min(totalPages, (AppState.favoritesPage || 1) + delta));
  renderCurrentView();
}

function goToFavoritesPage(targetPage) {
  AppState.favoritesPage = targetPage;
  renderCurrentView();
}

function setFavoritesPageSize(size) {
  AppState.favoritesPageSize = parseInt(size, 10) || 10;
  AppState.favoritesPage = 1;
  renderCurrentView();
}

// =============================================================================
// 4. MY FILES VIEW
// =============================================================================
function renderMyFilesView() {
  const currentUserName = InnomickData.currentUser && InnomickData.currentUser.name ? InnomickData.currentUser.name.trim().toLowerCase() : '';
  const myFiles = (InnomickData.myFiles || []).filter(doc => {
    const owner = (doc.owner || doc.uploadedBy || doc.author || doc.creator || doc.stakeHolder || '').trim().toLowerCase();
    if (!owner) return true;
    return owner === currentUserName;
  });

  return `
    <div class="content-header-row sp-library-toolbar">
      <button class="btn-primary-teal" onclick="triggerDeviceUpload(event, 'my-files')">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
        Upload
      </button>
    </div>

    <div class="sp-library-subbar">
      <div></div>
      ${renderListToolbarSearch('Search')}
    </div>

    <div class="view-breadcrumbs sp-library-breadcrumb">
      <span style="color: #008272; font-weight: 600;">Home</span>
    </div>

    ${myFiles.length > 0 ? renderSpListTable(myFiles, { preset: 'my-files', emptyMsg: 'No files match your search.' }) : `
      <div class="empty-state-card" style="padding: 60px 16px; text-align: center;">
        <p style="color: #008272; font-weight: 500; font-size: 15px; margin: 0;">You don't have any files yet. Use Upload to add your first document.</p>
      </div>
    `}
  `;
}
// =============================================================================
// 5. MY DEPARTMENTS & BRANCH SITES VIEW (53 SHAREPOINT COLUMNS)
// =============================================================================
const SP_DEPARTMENT_COLUMNS = [
  { key: 'name', label: 'Name', minWidth: '260px' },
  { key: 'id', label: 'ID', minWidth: '75px' },
  { key: 'category', label: 'Category', minWidth: '170px' },
  { key: 'subCategory', label: 'Sub Category', minWidth: '140px' },
  { key: 'department', label: 'Department', minWidth: '180px' },
  { key: 'owner', label: 'Owner Name', minWidth: '150px' },
  { key: 'status', label: 'Status', minWidth: '95px' },
  { key: 'uploadDate', label: 'Upload Date', minWidth: '135px' },
  { key: 'expireDate', label: 'Expire Date', minWidth: '135px' },
  { key: 'stakeHolder', label: 'StakeHolder', minWidth: '145px' },
  { key: 'description', label: 'Description', minWidth: '240px' },
  { key: 'employeeName', label: 'Employee Name', minWidth: '150px' },
  { key: 'employeeId', label: 'Employee ID', minWidth: '120px' },
  { key: 'startDate', label: 'Start Date', minWidth: '135px' },
  { key: 'endTime', label: 'End Time', minWidth: '100px' },
  { key: 'positionDesignation', label: 'Position/Designation', minWidth: '165px' },
  { key: 'versionNumber', label: 'Version Number', minWidth: '120px' },
  { key: 'campaignProjectName', label: 'Campaign/Project Name', minWidth: '185px' },
  { key: 'targetAudience', label: 'Target Audience', minWidth: '155px' },
  { key: 'eventLaunchDate', label: 'Event/Launch Date', minWidth: '140px' },
  { key: 'keyMessages', label: 'Key Messages', minWidth: '175px' },
  { key: 'assetFacilityName', label: 'Asset/Facility Name', minWidth: '165px' },
  { key: 'location', label: 'Location', minWidth: '155px' },
  { key: 'requestId', label: 'Request ID', minWidth: '125px' },
  { key: 'incidentRequestDate', label: 'Incident/Request Date', minWidth: '160px' },
  { key: 'resolutionStatus', label: 'Resolution Status', minWidth: '140px' },
  { key: 'effectiveDate', label: 'Effective Date', minWidth: '140px' },
  { key: 'approvalStatus', label: 'Approval Status', minWidth: '130px' },
  { key: 'referenceNumber', label: 'ReferenceNumber', minWidth: '140px' },
  { key: 'vendor', label: 'Vendor', minWidth: '130px' },
  { key: 'amount', label: 'Amount', minWidth: '100px' },
  { key: 'fiscalPeriodYear', label: 'Fiscal Period/Year', minWidth: '140px' },
  { key: 'approvalNumber', label: 'Approval Number', minWidth: '140px' },
  { key: 'paymentStatus', label: 'Payment Status', minWidth: '130px' },
  { key: 'auditInspectionDate', label: 'Audit/Inspection Date', minWidth: '160px' },
  { key: 'inspectionArea', label: 'Inspection Area', minWidth: '140px' },
  { key: 'findingsSummary', label: 'Findings Summary', minWidth: '185px' },
  { key: 'actionStatus', label: 'Action Status', minWidth: '120px' },
  { key: 'reviewDate', label: 'Review Date', minWidth: '135px' },
  { key: 'trainingWorkshopName', label: 'Training/Workshop Name', minWidth: '185px' },
  { key: 'trainerFacilitator', label: 'Trainer/Facilitator', minWidth: '155px' },
  { key: 'sessionDate', label: 'Session Date', minWidth: '135px' },
  { key: 'completionStatus', label: 'Completion Status', minWidth: '140px' },
  { key: 'equipmentName', label: 'Equipment Name', minWidth: '165px' },
  { key: 'equipmentLocation', label: 'Equipment Location', minWidth: '155px' },
  { key: 'documentValidityDate', label: 'Document Validity Date', minWidth: '175px' },
  { key: 'responsibleDepartment', label: 'Responsible Department', minWidth: '185px' },
  { key: 'reviewFrequency', label: 'Review Frequency', minWidth: '140px' },
  { key: 'caseProjectName', label: 'Case/Project Name', minWidth: '165px' },
  { key: 'validityPeriod', label: 'Validity Period', minWidth: '135px' },
  { key: 'expiryDate', label: 'Expiry Date', minWidth: '135px' },
  { key: 'complianceArea', label: 'Compliance Area', minWidth: '155px' },
  { key: 'favorite', label: 'Favorite', minWidth: '85px', align: 'center' }
];

function getDepartmentCellValue(doc, key) {
  if (!doc) return '—';
  if (doc.isFolder && key !== 'name') return '—';

  let val;
  switch (key) {
    case 'name':
      val = doc.name;
      break;
    case 'id':
      val = doc.id;
      break;
    case 'category':
      val = doc.category;
      break;
    case 'subCategory':
      val = doc.subCategory || doc.tagSubCategory;
      break;
    case 'department':
      val = doc.department;
      break;
    case 'owner':
      val = doc.ownerName || doc.owner;
      break;
    case 'status':
      val = doc.status;
      break;
    case 'uploadDate':
      val = doc.uploadDate;
      break;
    case 'expireDate':
      val = doc.expireDate;
      break;
    case 'stakeHolder':
      val = doc.stakeHolder || doc.stakeholder;
      break;
    case 'description':
      val = doc.description;
      break;
    case 'employeeName':
      val = doc.employeeName;
      break;
    case 'employeeId':
      val = doc.employeeId;
      break;
    case 'startDate':
      val = doc.startDate;
      break;
    case 'endTime':
      val = doc.endTime;
      break;
    case 'positionDesignation':
      val = doc.positionDesignation || doc.designation || doc.position;
      break;
    case 'versionNumber':
      val = doc.versionNumber || doc.version;
      break;
    case 'campaignProjectName':
      val = doc.campaignProjectName || doc.campaignName;
      break;
    case 'targetAudience':
      val = doc.targetAudience;
      break;
    case 'eventLaunchDate':
      val = doc.eventLaunchDate || doc.launchDate;
      break;
    case 'keyMessages':
      val = doc.keyMessages;
      break;
    case 'assetFacilityName':
      val = doc.assetFacilityName || doc.facilityName;
      break;
    case 'location':
      val = doc.location;
      break;
    case 'requestId':
      val = doc.requestId;
      break;
    case 'incidentRequestDate':
      val = doc.incidentRequestDate || doc.incidentDate;
      break;
    case 'resolutionStatus':
      val = doc.resolutionStatus;
      break;
    case 'effectiveDate':
      val = doc.effectiveDate;
      break;
    case 'approvalStatus':
      val = doc.approvalStatus;
      break;
    case 'referenceNumber':
      val = doc.referenceNumber;
      break;
    case 'vendor':
      val = doc.vendor;
      break;
    case 'amount':
      val = doc.amount;
      break;
    case 'fiscalPeriodYear':
      val = doc.fiscalPeriodYear || doc.fiscalYear;
      break;
    case 'approvalNumber':
      val = doc.approvalNumber;
      break;
    case 'paymentStatus':
      val = doc.paymentStatus;
      break;
    case 'auditInspectionDate':
      val = doc.auditInspectionDate || doc.inspectionDate;
      break;
    case 'inspectionArea':
      val = doc.inspectionArea;
      break;
    case 'findingsSummary':
      val = doc.findingsSummary;
      break;
    case 'actionStatus':
      val = doc.actionStatus;
      break;
    case 'reviewDate':
      val = doc.reviewDate;
      break;
    case 'trainingWorkshopName':
      val = doc.trainingWorkshopName || doc.trainingName;
      break;
    case 'trainerFacilitator':
      val = doc.trainerFacilitator || doc.trainer;
      break;
    case 'sessionDate':
      val = doc.sessionDate;
      break;
    case 'completionStatus':
      val = doc.completionStatus;
      break;
    case 'equipmentName':
      val = doc.equipmentName;
      break;
    case 'equipmentLocation':
      val = doc.equipmentLocation;
      break;
    case 'documentValidityDate':
      val = doc.documentValidityDate;
      break;
    case 'responsibleDepartment':
      val = doc.responsibleDepartment;
      break;
    case 'reviewFrequency':
      val = doc.reviewFrequency;
      break;
    case 'caseProjectName':
      val = doc.caseProjectName || doc.projectName;
      break;
    case 'validityPeriod':
      val = doc.validityPeriod;
      break;
    case 'expiryDate':
      val = doc.expiryDate || doc.expireDate;
      break;
    case 'complianceArea':
      val = doc.complianceArea;
      break;
    case 'favorite':
      return doc.favorite;
    default:
      val = doc[key];
  }

  if (val === undefined || val === null || val === '' || val === '-') {
    return '—';
  }
  return String(val);
}

// =============================================================================
// SHAREPOINT LIST TABLE ENGINE (one renderer for every document library page)
// Base columns follow each live DMS page; any extra metadata column that has a
// value in the listed documents is appended, exactly like the SharePoint list view.
// Filters sit under every header, the toolbar search box searches every column.
// =============================================================================
const SP_LIST_PRESETS = {
  'my-files': {
    base: ['name', 'favorite', 'category', 'status', 'department', 'expireDate', 'subCategory', 'owner', 'uploadDate', 'description', 'stakeHolder', 'vendor'],
    labels: { owner: 'Owner' },
    favoriteStyle: 'star',
    dynamic: true
  },
  'organization': {
    base: ['name', 'category', 'subCategory', 'department', 'owner', 'status', 'uploadDate', 'expireDate', 'description', 'stakeHolder'],
    favoriteStyle: 'star',
    dynamic: true
  },
  'department': {
    base: ['name', 'category', 'subCategory', 'department', 'owner', 'status', 'uploadDate', 'expireDate', 'description', 'stakeHolder'],
    favoriteStyle: 'star',
    dynamic: true
  },
  'expiring': {
    base: ['name', 'category', 'status', 'department', 'referenceNumber', 'favorite', 'stakeHolder', 'owner', 'subCategory', 'uploadDate', 'id', 'vendor', 'expireDate'],
    favoriteStyle: 'text',
    dynamic: false
  },
  'shared': {
    base: ['name', 'category', 'status', 'department', 'referenceNumber', 'favorite', 'stakeHolder', 'owner', 'subCategory', 'uploadDate', 'id', 'vendor', 'expireDate'],
    favoriteStyle: 'text',
    dynamic: false
  }
};

const SP_LIST_COLUMN_INDEX = Object.fromEntries(SP_DEPARTMENT_COLUMNS.map(c => [c.key, c]));
const SP_LIST_NON_DYNAMIC = new Set(['id', 'location', 'favorite', 'expiryDate']);

function getListColumns(docs, preset) {
  const cfg = SP_LIST_PRESETS[preset] || SP_LIST_PRESETS.organization;
  const keys = [...cfg.base];
  if (cfg.dynamic) {
    SP_DEPARTMENT_COLUMNS.forEach(col => {
      if (keys.includes(col.key) || SP_LIST_NON_DYNAMIC.has(col.key)) return;
      if (docs.some(d => !d.isFolder && getDepartmentCellValue(d, col.key) !== '—')) keys.push(col.key);
    });
  }
  return keys.map(k => ({
    ...(SP_LIST_COLUMN_INDEX[k] || { key: k, label: k, minWidth: '130px' }),
    label: (cfg.labels && cfg.labels[k]) || (SP_LIST_COLUMN_INDEX[k] || {}).label || k
  }));
}

function getListCellText(doc, key) {
  if (key === 'favorite') return doc.favorite ? 'Yes' : 'No';
  return getDepartmentCellValue(doc, key);
}

function applyListFiltersAndSort(docs, columns) {
  let list = [...docs];
  const filters = Object.entries(AppState.columnFilters || {}).filter(([, v]) => v && String(v).trim());
  if (filters.length) {
    list = list.filter(d => d.isFolder || filters.every(([k, v]) =>
      getListCellText(d, k).toLowerCase().includes(String(v).trim().toLowerCase())
    ));
  }
  const q = (AppState.globalSearchQuery || '').trim().toLowerCase();
  if (q) {
    list = list.filter(d => d.isFolder
      ? (d.name || '').toLowerCase().includes(q)
      : columns.some(c => getListCellText(d, c.key).toLowerCase().includes(q)));
  }
  const sortKey = AppState.listSortKey;
  const asc = AppState.listSortAsc !== false;
  list.sort((a, b) => {
    if (a.isFolder && !b.isFolder) return -1;
    if (!a.isFolder && b.isFolder) return 1;
    if (!sortKey) return 0;
    const va = getListCellText(a, sortKey);
    const vb = getListCellText(b, sortKey);
    if (va === vb) return 0;
    if (va === '—') return 1;
    if (vb === '—') return -1;
    return asc ? va.localeCompare(vb, undefined, { numeric: true }) : vb.localeCompare(va, undefined, { numeric: true });
  });
  return list;
}

function sortListByColumn(key) {
  if (AppState.listSortKey === key) {
    AppState.listSortAsc = !AppState.listSortAsc;
  } else {
    AppState.listSortKey = key;
    AppState.listSortAsc = true;
  }
  renderCurrentView();
}

function setListPage(page) {
  AppState.listPage = Math.max(1, page);
  renderCurrentView();
}

function setListPageSize(size) {
  AppState.listPageSize = parseInt(size, 10) || 10;
  AppState.listPage = 1;
  renderCurrentView();
}

function renderListRowMenu(item, allowPublish) {
  if (AppState.openContextMenuDocId !== item.id) return '';
  return `
    <div class="row-context-menu" id="context_menu_${item.id}" onclick="event.stopPropagation()">
      <button class="context-menu-item" onclick="openVersionHistoryModal('${item.id}')">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/><path d="M12 7v5l4 2"/></svg>
        <span>Version History</span>
      </button>
      <button class="context-menu-item" onclick="toggleCheckOut('${item.id}')">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>
        <span>${item.isCheckedOut ? 'Check In' : 'Check Out'}</span>
      </button>
      <button class="context-menu-item" onclick="openManageAccessModal('${item.id}')">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="7.5" cy="15.5" r="5.5"/><path d="m21 2-9.6 9.6"/><path d="m15.5 7.5 3 3L22 7l-3-3"/></svg>
        <span>Manage Access</span>
      </button>
      <button class="context-menu-item" onclick="downloadDocument('${item.id}')">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
        <span>Download</span>
      </button>
      <button class="context-menu-item" onclick="toggleFavorite('${item.id}')">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="${item.favorite ? '#008272' : 'none'}" stroke="${item.favorite ? '#008272' : 'currentColor'}" stroke-width="2"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>
        <span>${item.favorite ? 'Remove Favourite' : 'Add Favourite'}</span>
      </button>
      ${allowPublish ? `
      <button class="context-menu-item" onclick="publishDocument('${item.id}')">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
        <span>Publish</span>
      </button>` : ''}
      <button class="context-menu-item danger" onclick="deleteDocumentToRecycle('${item.id}')">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#ef4444" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
        <span>Delete</span>
      </button>
    </div>`;
}

function renderListCell(item, col, cfg, opts) {
  if (col.key === 'name') {
    return `
      <td class="sp-list-name-cell">
        <div class="sp-list-name-wrap">
          <div class="sp-list-name-main">
            ${getFileIconSvg(item.type)}
            <span class="file-link" onclick="toggleDocumentMenu(event, '${item.id}')" title="${escapeHtmlDept(item.name)}">${escapeHtmlDept(item.name)}</span>
            ${item.isCheckedOut ? `<span title="Checked out by ${item.checkedOutBy}" class="checked-out-lock"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg></span>` : ''}
          </div>
          <div class="row-actions-group">
            <button class="row-action-btn" title="More Actions" onclick="toggleRowContextMenu(event, '${item.id}')">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="currentColor"><circle cx="12" cy="12" r="2"/><circle cx="19" cy="12" r="2"/><circle cx="5" cy="12" r="2"/></svg>
            </button>
            <button class="table-share-btn" title="Share with others" onclick="openManageAccessModal('${item.id}')">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 12v8a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-8"/><polyline points="16 6 12 2 8 6"/><line x1="12" y1="2" x2="12" y2="15"/></svg>
            </button>
            ${renderListRowMenu(item, opts.allowPublish)}
          </div>
        </div>
      </td>`;
  }
  if (col.key === 'favorite') {
    if (cfg.favoriteStyle === 'text') {
      return `<td style="text-align: center;"><span class="sp-list-fav-text" onclick="toggleFavorite('${item.id}')">${item.favorite ? 'Yes' : 'No'}</span></td>`;
    }
    const isFav = !!item.favorite;
    return `
      <td style="text-align: center;">
        <button class="favorite-star-btn" title="${isFav ? 'Remove from favorites' : 'Add to favorites'}" onclick="toggleDeptItemFavorite('${item.id}', event)">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="${isFav ? '#eab308' : 'none'}" stroke="${isFav ? '#ca8a04' : '#64748b'}" stroke-width="1.8"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>
        </button>
      </td>`;
  }
  const val = getDepartmentCellValue(item, col.key);
  if (val === '—') return `<td class="sp-list-empty-cell"><span class="sp-dept-dash">&mdash;</span></td>`;
  return `<td>${escapeHtmlDept(val)}</td>`;
}

function renderListPagination(page, totalPages, pageSize, totalItems) {
  const btn = (label, target, disabled, title) =>
    `<button type="button" class="pagination-btn" title="${title}" ${disabled ? 'disabled' : `onclick="setListPage(${target})"`}>${label}</button>`;
  return `
    <div class="table-pagination-footer">
      <div class="pagination-controls">
        ${btn('&laquo;', 1, page <= 1, 'First page')}
        ${btn('&lsaquo;', page - 1, page <= 1, 'Previous page')}
        <span>${page} / ${totalPages} pages</span>
        ${btn('&rsaquo;', page + 1, page >= totalPages, 'Next page')}
        ${btn('&raquo;', totalPages, page >= totalPages, 'Last page')}
      </div>
      <div style="display: flex; align-items: center; gap: 14px;">
        <span style="font-size: 12.5px; color: #64748b;">${totalItems} items</span>
        <div style="display: flex; align-items: center; gap: 8px;">
          <select class="pagination-select" onchange="setListPageSize(this.value)">
            ${[10, 20, 50, 100].map(n => `<option value="${n}" ${pageSize === n ? 'selected' : ''}>${n}</option>`).join('')}
          </select>
          <span>Items per page</span>
        </div>
      </div>
    </div>`;
}

// docs: rows to show (folders flagged isFolder). opts: { preset, emptyMsg, allowPublish, onFolderClick }
function renderSpListTable(docs, opts = {}) {
  const preset = opts.preset || 'organization';
  const cfg = SP_LIST_PRESETS[preset] || SP_LIST_PRESETS.organization;
  const columns = getListColumns(docs, preset);
  const rows = applyListFiltersAndSort(docs, columns);

  const pageSize = AppState.listPageSize || 10;
  const fileRows = rows.filter(r => !r.isFolder);
  const folderRows = rows.filter(r => r.isFolder);
  const totalPages = Math.max(1, Math.ceil(fileRows.length / pageSize));
  const page = Math.min(AppState.listPage || 1, totalPages);
  const pageRows = [...(page === 1 ? folderRows : []), ...fileRows.slice((page - 1) * pageSize, page * pageSize)];
  const folderClick = opts.onFolderClick || 'navigateToDeptFolder';

  return `
    <div class="sp-dept-table-wrapper">
      <table class="sp-dept-53-table sp-list-table">
        <thead>
          <tr>
            <th class="sp-list-select-col"></th>
            ${columns.map(col => `
              <th style="min-width: ${col.minWidth}; text-align: ${col.align || 'center'};">
                <div class="sp-dept-th-title" onclick="sortListByColumn('${col.key}')" title="Sort by ${col.label}">
                  <span>${col.label}</span>
                  ${col.key === 'favorite' ? '' : `<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>`}
                </div>
                ${col.key === 'favorite' ? '' : `
                <div class="sp-list-th-filter">
                  <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
                  <input type="text" class="column-search-input" data-filter-col="${col.key}" aria-label="Filter ${col.label}"
                         value="${escapeHtmlDept(AppState.columnFilters[col.key] || '')}"
                         oninput="setColumnFilter('${col.key}', this.value, this)">
                </div>`}
              </th>
            `).join('')}
          </tr>
        </thead>
        <tbody>
          ${pageRows.length === 0 ? `
            <tr>
              <td colspan="${columns.length + 1}" style="text-align: center; padding: 44px 20px; color: #64748b; font-size: 14px;">
                ${opts.emptyMsg || 'No documents found.'}
              </td>
            </tr>
          ` : pageRows.map(item => item.isFolder ? `
            <tr class="table-folder-row" style="cursor: pointer;" onclick="${folderClick}('${item.name}')">
              <td class="sp-list-select-col"></td>
              <td class="sp-list-name-cell">
                <div class="sp-list-name-wrap">
                  <div class="sp-list-name-main">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="#eab308" stroke="none"><path d="M4 20h16a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.93a2 2 0 0 1-1.66-.9l-.82-1.2A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13c0 1.1.9 2 2 2Z"/></svg>
                    <span class="dept-folder-link">${escapeHtmlDept(item.name)}</span>
                  </div>
                  <span class="sp-list-folder-dots">&middot;&middot;&middot;</span>
                </div>
              </td>
              ${columns.slice(1).map(() => `<td class="sp-list-empty-cell"><span class="sp-dept-dash">&mdash;</span></td>`).join('')}
            </tr>
          ` : `
            <tr>
              <td class="sp-list-select-col"><span class="sp-select-circle" onclick="toggleRowSelect(this, event)"></span></td>
              ${columns.map(col => renderListCell(item, col, cfg, opts)).join('')}
            </tr>
          `).join('')}
        </tbody>
      </table>
    </div>
    ${renderListPagination(page, totalPages, pageSize, fileRows.length)}
  `;
}

// Toolbar shared by library pages: search box, refresh, sort toggle
function renderListToolbarSearch(placeholder) {
  return `
    <div style="display: flex; align-items: center; gap: 8px;">
      <div class="search-input-wrapper" style="width: 250px;">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
        <input type="text" id="listSearchInput" placeholder="${placeholder || 'Search'}" value="${escapeHtmlDept(AppState.globalSearchQuery || '')}" oninput="handleSearch(this.value, this)">
      </div>
      <button class="icon-action-btn" title="Refresh" onclick="refreshData()">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/><path d="M8 16H3v5"/></svg>
      </button>
      <button class="icon-action-btn" title="Sort by Name" onclick="sortListByColumn('name')">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="7 3 7 21"/><polyline points="3 7 7 3 11 7"/><polyline points="17 21 17 3"/><polyline points="13 17 17 21 21 17"/></svg>
      </button>
    </div>`;
}

// Kept for older call sites: the classic 13-column layout (Expiring Files order)
function render13ColumnsTable(docsList, emptyMsg = 'No documents found.') {
  return renderSpListTable(docsList, { preset: 'expiring', emptyMsg });
}

function renderMyDepartmentView() {
  return renderDepartmentView();
}

function renderDepartmentView() {
  let branch = AppState.activeBranch || 'Innomick Delhi';
  let dept = AppState.activeDepartment || 'Information Technology';
  
  const displayBranch = branch;
  let canonicalBranch = branch;
  if (branch.includes('Hyderabad')) {
    canonicalBranch = 'Innomick Hyderabad';
  } else if (branch.includes('Delhi')) {
    canonicalBranch = 'Innomick Delhi';
  }

  const isHyderabad = canonicalBranch.includes('Hyderabad') || dept === 'Executive Director Office';
  const isITDept = !isHyderabad && (dept.toLowerCase().includes('information') || dept.toLowerCase().includes('it') || dept.toLowerCase().includes('technology'));

  let docsList = [];
  if (InnomickData.departmentSites && InnomickData.departmentSites[canonicalBranch] && InnomickData.departmentSites[canonicalBranch][dept]) {
    docsList = [...InnomickData.departmentSites[canonicalBranch][dept]];
  } else if (InnomickData.departmentSites && InnomickData.departmentSites[branch] && InnomickData.departmentSites[branch][dept]) {
    docsList = [...InnomickData.departmentSites[branch][dept]];
  } else {
    const key = isHyderabad ? 'Innomick Hyderabad' : 'Innomick Delhi';
    const deptKey = isHyderabad ? 'Executive Director Office' : 'Information Technology';
    docsList = [...((InnomickData.departmentSites && InnomickData.departmentSites[key] && InnomickData.departmentSites[key][deptKey]) || [])];
  }

  // Ensure documents uploaded to folders have their parentFolder set
  docsList.forEach(d => {
    if (d.folder && d.folder !== '-' && !d.parentFolder) {
      d.parentFolder = d.folder;
    }
    if (d.location && d.location.includes('>') && !d.parentFolder && !d.isFolder) {
      const parts = d.location.split('>').map(s => s.trim());
      const lastPart = parts[parts.length - 1];
      if (lastPart && lastPart !== d.department) {
        d.parentFolder = lastPart;
      }
    }
  });

  // Folder filtering based on AppState.activeDeptFolder
  const currentFolder = AppState.activeDeptFolder;
  if (!currentFolder) {
    // Root level (Home): show root folders and root files (where !parentFolder)
    docsList = docsList.filter(d => !d.parentFolder || d.parentFolder === 'root');
  } else if (currentFolder === 'Direct Uploads') {
    // Inside Direct Uploads: show subfolders and files inside Direct Uploads
    docsList = docsList.filter(d => d.parentFolder === 'Direct Uploads');
  } else {
    // Inside a specific folder (e.g. 'Testing', 'Office Files', 'CIO Office', 'test', 'Test')
    docsList = docsList.filter(d => d.parentFolder === currentFolder || d.folder === currentFolder || (d.parentFolder && d.parentFolder.endsWith('/' + currentFolder)));
  }

  return `
    <div class="content-header-row" style="margin-bottom: 12px; display: flex; align-items: center; justify-content: space-between;">
      <div style="display: flex; align-items: center; gap: 8px;">
        <button class="btn-primary-teal" onclick="triggerDeviceUpload(event, 'my-department')" style="display: inline-flex; align-items: center; gap: 6px;">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
            <polyline points="17 8 12 3 7 8"/>
            <line x1="12" y1="3" x2="12" y2="15"/>
          </svg>
          Upload
        </button>
        <button class="btn-direct-upload" onclick="triggerDirectUpload(event, 'my-department')" style="display: inline-flex; align-items: center; gap: 6px;">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
            <polyline points="14 2 14 8 20 8"/>
            <line x1="16" y1="13" x2="8" y2="13"/>
            <line x1="16" y1="17" x2="8" y2="17"/>
          </svg>
          Direct Upload
        </button>
      </div>
    </div>

    <div style="font-size: 13.5px; font-weight: 500; color: #1e293b; margin-bottom: 6px;">
      ${displayBranch} - ${dept}
    </div>

    <div class="view-breadcrumbs" style="margin-bottom: 14px; display: flex; align-items: center; justify-content: space-between;">
      <div style="font-size: 13.5px;">
        ${!AppState.activeDeptFolder ? `
          <span style="color: #008272; font-weight: 600;">Home</span>
        ` : ['Home', ...getDeptFolderPath(AppState.activeDeptFolder)].map((part, i, arr) => i === arr.length - 1
          ? `<span style="color: #1e293b; font-weight: 600;">${part}</span>`
          : `<a onclick="navigateToDeptFolder(${i === 0 ? 'null' : `'${part.replace(/'/g, "\\'")}'`})" style="color: #008272; font-weight: 600; cursor: pointer; text-decoration: none;">${part}</a>`
        ).join(' &gt; ')}
      </div>

      ${renderListToolbarSearch('Search')}
    </div>

    ${renderSpListTable(docsList, { preset: 'department', allowPublish: isITDept, emptyMsg: AppState.activeDeptFolder ? 'This folder is empty.' : 'No documents found.' })}
  `;
}

function sortDepartmentByColumn(key) {
  if (AppState.deptSortKey === key) {
    AppState.deptSortAsc = !AppState.deptSortAsc;
  } else {
    AppState.deptSortKey = key;
    AppState.deptSortAsc = true;
  }
  renderCurrentView();
}

function toggleDepartmentColumnFilter(colKey, event) {
  if (event) event.stopPropagation();
  if (AppState.openDeptColFilterKey === colKey) {
    AppState.openDeptColFilterKey = null;
  } else {
    AppState.openDeptColFilterKey = colKey;
  }
  renderCurrentView();
  setTimeout(() => {
    const input = document.getElementById(`deptColFilterInput_${colKey}`);
    if (input) input.focus();
  }, 50);
}

function navigateToDeptFolder(folderName) {
  AppState.activeDeptFolder = folderName || null;
  const branch = AppState.activeBranch || 'Innomick Hyderabad';
  const dept = AppState.activeDepartment || 'Executive Director Office';
  const b64Branch = btoa(branch);
  const b64Dept = btoa(dept);
  if (folderName) {
    const b64Folder = btoa(folderName);
    window.location.hash = `#/department?branch=${b64Branch}&dept=${b64Dept}&folder=${b64Folder}`;
  } else {
    window.location.hash = `#/department?branch=${b64Branch}&dept=${b64Dept}`;
  }
  renderCurrentView();
}

function handleDepartmentColumnFilterInput(colKey, value) {
  if (!AppState.deptColFilters) AppState.deptColFilters = {};
  if (value && value.trim()) {
    AppState.deptColFilters[colKey] = value.trim().toLowerCase();
  } else {
    delete AppState.deptColFilters[colKey];
  }
  renderCurrentView();
}

function toggleDeptItemFavorite(docId, event) {
  if (event) event.stopPropagation();
  toggleFavorite(docId);
}

function escapeHtmlDept(text) {
  if (text === undefined || text === null) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function handleDepartmentSearch(val, inputEl) {
  AppState.departmentSearchQuery = val;
  window._lastSearchFocus = {
    id: inputEl && inputEl.id ? inputEl.id : 'deptSearchInput',
    cursor: inputEl && typeof inputEl.selectionStart === 'number' ? inputEl.selectionStart : (val ? val.length : 0)
  };
  window._lastFilterFocus = null;
  renderCurrentView();
}

function toggleRowSelect(circle, event) {
  if (event) event.stopPropagation();
  circle.classList.toggle('selected');
}

function toggleSelectAllDept(circle) {
  const isSel = circle.classList.toggle('selected');
  document.querySelectorAll('.dms-data-table tbody .sp-select-circle').forEach(c => {
    if (isSel) c.classList.add('selected');
    else c.classList.remove('selected');
  });
}

// =============================================================================
// 6. MY ORGANIZATION VIEW (ALL 11 COLUMNS WITH EXACT SCREENSHOT MATCH)
// =============================================================================
function navigateToFolder(folderName) {
  AppState.currentFolder = folderName;
  renderCurrentView();
}

function renderMyOrganizationView() {
  // Gap 30: Documents pending approval should not be visible in My Organization until approval workflow is completed
  let docs = (InnomickData.documents || []).filter(d => {
    if (d && (d.status === 'In Review' || d.status === 'Approval Pending' || d.status === 'Pending Approval')) return false;
    return true;
  });

  // Inside a folder only its documents are listed; the library root shows the folder row first
  if (AppState.currentFolder) {
    docs = docs.filter(d => (d.subCategory && d.subCategory.toLowerCase().includes('client')) || d.category === 'Client Deliverables & PoCs' || d.category === 'Architecture Blueprints');
  } else {
    docs = [{ id: 'folder-org-root', name: 'Client Deliverables Archive', isFolder: true }, ...docs];
  }

  return `
    <div class="content-header-row sp-library-toolbar">
      <button class="btn-primary-teal" onclick="triggerDeviceUpload(event, 'my-organization')">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
        Upload
      </button>
    </div>

    <div class="sp-library-subbar">
      <div></div>
      ${renderListToolbarSearch('Search')}
    </div>

    <div class="view-breadcrumbs sp-library-breadcrumb">
      ${AppState.currentFolder ? `
        <a onclick="navigateToFolder(null)" style="color: #008272; font-weight: 600; cursor: pointer;">Home</a> &gt; <span style="color: #1e293b; font-weight: 600;">${AppState.currentFolder}</span>
      ` : `<span style="color: #008272; font-weight: 600;">Home</span>`}
    </div>

    ${renderSpListTable(docs, { preset: 'organization', onFolderClick: 'navigateToFolder', emptyMsg: 'No organization documents found.' })}
  `;
}

function setColumnFilter(col, val, inputEl) {
  AppState.columnFilters[col] = val;
  AppState.listPage = 1;
  window._lastFilterFocus = {
    col: col,
    cursor: inputEl && typeof inputEl.selectionStart === 'number' ? inputEl.selectionStart : (val ? val.length : 0)
  };
  window._lastSearchFocus = null;
  renderCurrentView();
}

function toggleSelectAllOrg(masterCb) {
  const cbs = document.querySelectorAll('.dms-data-table tbody .select-row-cb');
  cbs.forEach(cb => cb.checked = masterCb.checked);
}

// Row Context Menu Controller
function toggleRowContextMenu(event, docId) {
  event.stopPropagation();
  if (AppState.openContextMenuDocId === docId) {
    AppState.openContextMenuDocId = null;
  } else {
    AppState.openContextMenuDocId = docId;
  }
  renderCurrentView();
}

function closeContextMenu() {
  if (AppState.openContextMenuDocId !== null) {
    AppState.openContextMenuDocId = null;
    renderCurrentView();
  }
}

// Check Out / Check In Feature
// Helper to locate document across all collections
function findAnyDoc(id) {
  if (!id && id !== 0) return null;
  const match = (d) => d && (String(d.id) === String(id) || d.id === id);
  let found = (InnomickData.documents || []).find(match) ||
         (InnomickData.recentDocuments || []).find(match) ||
         (InnomickData.homeAllFiles || []).find(match) ||
         (InnomickData.expiringFiles || []).find(match) ||
         (InnomickData.myFiles || []).find(match) ||
         (InnomickData.sharedFiles || []).find(match);
  if (found) return found;

  if (InnomickData.departmentSites) {
    for (const b in InnomickData.departmentSites) {
      for (const d in InnomickData.departmentSites[b]) {
        const item = (InnomickData.departmentSites[b][d] || []).find(match);
        if (item) return item;
      }
    }
  }
  return null;
}

function toggleCheckOut(docId) {
  const doc = findAnyDoc(docId);
  if (doc) {
    doc.isCheckedOut = !doc.isCheckedOut;
    doc.checkedOutBy = doc.isCheckedOut ? InnomickData.currentUser.name : '';
    closeContextMenu();
    showToast(doc.isCheckedOut 
      ? `Document "${doc.name}" checked out to you. Locked for other users.` 
      : `Document "${doc.name}" checked in successfully.`
    );
    renderCurrentView();
  }
}

// Delete Document (Moves to Recycle Bin with original location tracked for Gap 34)
function deleteDocumentToRecycle(docId) {
  let doc = null;
  let sourceLibrary = AppState.currentRoute === 'my-files' ? 'my-files'
                    : (AppState.currentRoute === 'department' || AppState.currentRoute === 'my-department') ? 'department'
                    : 'organization';
  const branch = AppState.activeBranch || 'Innomick Delhi';
  const dept = AppState.activeDepartment || 'Information Technology';

  // Search and remove from InnomickData.documents
  const idxDoc = InnomickData.documents.findIndex(d => String(d.id) === String(docId));
  if (idxDoc !== -1) {
    doc = InnomickData.documents.splice(idxDoc, 1)[0];
    sourceLibrary = 'organization';
  }

  // Also remove from expiringFiles if present
  const idxExp = InnomickData.expiringFiles.findIndex(d => String(d.id) === String(docId));
  if (idxExp !== -1) {
    if (!doc) doc = InnomickData.expiringFiles[idxExp];
    InnomickData.expiringFiles.splice(idxExp, 1);
  }

  // Also remove from myFiles if present
  const idxMy = InnomickData.myFiles.findIndex(d => String(d.id) === String(docId));
  if (idxMy !== -1) {
    if (!doc) doc = InnomickData.myFiles[idxMy];
    InnomickData.myFiles.splice(idxMy, 1);
    sourceLibrary = 'my-files';
  }

  // Also remove from sharedFiles if present
  const idxShared = InnomickData.sharedFiles.findIndex(d => String(d.id) === String(docId));
  if (idxShared !== -1) {
    if (!doc) doc = InnomickData.sharedFiles[idxShared];
    InnomickData.sharedFiles.splice(idxShared, 1);
  }

  // Also remove from recent/favorites if present
  if (InnomickData.recentDocuments) {
    const idxRec = InnomickData.recentDocuments.findIndex(d => String(d.id) === String(docId));
    if (idxRec !== -1) InnomickData.recentDocuments.splice(idxRec, 1);
  }

  // Also remove from allFiles if present
  if (InnomickData.allFiles) {
    const idxAll = InnomickData.allFiles.findIndex(d => String(d.id) === String(docId));
    if (idxAll !== -1) {
      if (!doc) doc = InnomickData.allFiles[idxAll];
      InnomickData.allFiles.splice(idxAll, 1);
    }
  }

  // Also remove from departmentSites if present
  if (InnomickData.departmentSites) {
    for (const b in InnomickData.departmentSites) {
      for (const d in InnomickData.departmentSites[b]) {
        const idxDept = InnomickData.departmentSites[b][d].findIndex(x => String(x.id) === String(docId));
        if (idxDept !== -1) {
          if (!doc) doc = InnomickData.departmentSites[b][d][idxDept];
          InnomickData.departmentSites[b][d].splice(idxDept, 1);
          sourceLibrary = 'department';
        }
      }
    }
  }

  // Fallback to findAnyDoc or synthetic doc
  if (!doc) {
    doc = findAnyDoc(docId) || { id: docId, name: 'Document', type: 'pdf', department: 'General' };
  }

  InnomickData.recycleBin.unshift({
    id: 'del-' + Date.now(),
    name: doc.name || 'Untitled Document',
    type: doc.type || 'pdf',
    deletedDate: 'Just now',
    originalLocation: doc.location || `DMS / ${doc.department || 'General'}`,
    size: doc.size || '1.5 MB',
    _originalDoc: { ...doc },
    _sourceLibrary: sourceLibrary,
    _branch: branch,
    _department: doc.department || dept
  });

  AppState.openContextMenuDocId = null;
  closeContextMenu();
  showToast(`"${doc.name}" moved to Recycle Bin.`);
  renderCurrentView();
}

function publishDocument(docId) {
  closeContextMenu();
  const doc = findAnyDoc(docId);
  if (!doc) return;
  doc.status = 'Published';
  showToast(`Document "${doc.name}" published successfully to Enterprise Portal.`, 'success');
  renderCurrentView();
}

// =============================================================================
// 7. VERSION HISTORY MODAL (Exact SharePoint Recreation)
// =============================================================================
function openVersionHistoryModal(docId) {
  closeContextMenu();
  const doc = findAnyDoc(docId) || InnomickData.documents[0];
  const modal = document.getElementById('versionHistoryModal');
  const title = document.getElementById('vhDocTitle');
  const body = document.getElementById('vhModalBody');

  if (title) title.innerText = `Version History: ${doc.name}`;

  const versions = doc.versions && doc.versions.length > 0 ? doc.versions : [
    {
      no: "2.0",
      modified: `${doc.uploadDate} 14:30`,
      modifiedBy: doc.owner,
      category: doc.category,
      department: doc.department,
      uploadDate: doc.uploadDate,
      expireDate: doc.expireDate,
      inspectionArea: "Production Compliance Review",
      findingsSummary: "Document verified and audit criteria satisfied",
      actionStatus: "Completed",
      reviewDate: doc.uploadDate
    },
    {
      no: "1.0",
      modified: `${doc.uploadDate} 09:15`,
      modifiedBy: "MD Faisal Ahmed",
      category: doc.category,
      department: doc.department,
      uploadDate: doc.uploadDate,
      expireDate: doc.expireDate,
      inspectionArea: "Initial Draft Submission",
      findingsSummary: "Baseline drafted and awaiting sign-off",
      actionStatus: "Active",
      reviewDate: doc.uploadDate
    }
  ];

  if (body) {
    body.innerHTML = `
      <button class="vh-delete-all-btn" onclick="showToast('Version cleanup policy restricts deleting historical base audit records.', 'warning')">Delete All Versions</button>
      <table class="vh-table">
        <thead>
          <tr>
            <th style="width: 60px;">No. &darr;</th>
            <th style="width: 250px;">Modified</th>
            <th>Modified By</th>
          </tr>
        </thead>
        <tbody>
          ${versions.map(v => `
            <tr>
              <td class="vh-version-no">${v.no}</td>
              <td>
                <a class="vh-date-link" onclick="showToast('Restoring snapshot v${v.no}...')">${v.modified}</a>
                <div class="vh-meta-tree">
                  <span class="vh-meta-label">Category:</span><span>${v.category}</span>
                  <span class="vh-meta-label">Department:</span><span>${v.department}</span>
                  <span class="vh-meta-label">Upload Date:</span><span>${v.uploadDate}</span>
                  <span class="vh-meta-label">Expire Date:</span><span>${v.expireDate}</span>
                  <span class="vh-meta-label">Inspection Area:</span><span>${v.inspectionArea || 'General'}</span>
                  <span class="vh-meta-label">Action Status:</span><span>${v.actionStatus || 'Completed'}</span>
                </div>
              </td>
              <td>
                <div class="vh-user">
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#008272" stroke-width="2">
                    <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/>
                  </svg>
                  <span>${v.modifiedBy}</span>
                </div>
              </td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;
  }

  if (modal) modal.classList.add('active');
}

function closeVersionHistoryModal() {
  const modal = document.getElementById('versionHistoryModal');
  if (modal) modal.classList.remove('active');
}

// =============================================================================
// 8. SHARE MODAL DIALOG (Exact SharePoint Online Recreation)
// =============================================================================
let currentShareDoc = null;
let currentSharePermission = 'view'; // 'view' or 'edit'
let selectedShareRecipients = [];

function openShareModal(docId) {
  closeContextMenu();
  closeAllFilePopovers();
  AppState.activeDocIdForAccess = docId;
  const doc = findAnyDoc(docId) || InnomickData.documents[0] || { name: 'Ref Doc Demo 1.docx', id: 'doc-1' };
  currentShareDoc = doc;
  selectedShareRecipients = [];
  currentSharePermission = 'view';

  const modal = document.getElementById('manageAccessModal');
  const title = document.getElementById('shareDocModalTitle');
  const recipientInput = document.getElementById('shareRecipientInput');
  const messageInput = document.getElementById('shareMessageInput');
  const chipsContainer = document.getElementById('shareRecipientChips');
  const copyBtnText = document.getElementById('shareCopyLinkText');

  if (title) title.innerText = `Share "${doc.name}"`;
  if (recipientInput) recipientInput.value = '';
  if (messageInput) messageInput.value = '';
  if (chipsContainer) chipsContainer.innerHTML = '';
  if (copyBtnText) copyBtnText.innerText = 'Copy link';

  setSharePermission('view');
  hideShareAutocomplete();
  closeSharePermMenu();

  if (modal) modal.classList.add('active');
}

// Backwards compatibility alias
const openManageAccessModal = openShareModal;

function closeShareModal() {
  const modal = document.getElementById('manageAccessModal');
  if (modal) modal.classList.remove('active');
  closeSharePermMenu();
  hideShareAutocomplete();
}

const closeManageAccessModal = closeShareModal;

function focusShareRecipientBox(focused) {
  const box = document.getElementById('shareRecipientContainer');
  if (box) {
    if (focused) box.classList.add('focused');
    else box.classList.remove('focused');
  }
}

function toggleSharePermMenu(e) {
  if (e) e.stopPropagation();
  const menu = document.getElementById('sharePermMenu');
  if (!menu) return;
  const isShown = menu.style.display === 'block';
  menu.style.display = isShown ? 'none' : 'block';
}

function closeSharePermMenu() {
  const menu = document.getElementById('sharePermMenu');
  if (menu) menu.style.display = 'none';
}

function setSharePermission(perm) {
  currentSharePermission = perm;
  const checkEdit = document.getElementById('checkPermEdit');
  const checkView = document.getElementById('checkPermView');
  const permIcon = document.getElementById('sharePermIcon');

  if (checkEdit) checkEdit.style.display = perm === 'edit' ? 'inline' : 'none';
  if (checkView) checkView.style.display = perm === 'view' ? 'inline' : 'none';

  if (permIcon) {
    if (perm === 'edit') {
      permIcon.innerHTML = `<path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>`;
    } else {
      permIcon.innerHTML = `<path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/>`;
    }
  }

  closeSharePermMenu();
}

const candidateSharers = [
  { name: 'Mohammed Wasifuddin', email: 'm.wasifuddin@innomick.com', dept: 'Corporate Management', initials: 'MW' },
  { name: 'Sayed Asif Ahmed', email: 's.asif@innomick.com', dept: 'Central Supply Chain', initials: 'SA' },
  { name: 'Mohammed Zuberuddin', email: 'm.zuber@innomick.com', dept: 'Delivery & Solutions', initials: 'MZ' },
  { name: 'MD Faisal Ahmed', email: 'faisal.ahmed@innomick.com', dept: 'Information Technology', initials: 'FA' },
  { name: 'Arjun Mehta', email: 'a.mehta@innomick.com', dept: 'Finance & Accounts', initials: 'AM' },
  { name: 'Fatima Shaikh', email: 'f.alzahra@innomick.com', dept: 'Human Resources', initials: 'FZ' },
  { name: 'Enterprise Steering Committee', email: 'steering.committee@innomick.com', dept: 'Cross-functional Group', initials: 'EC' }
];

function handleShareRecipientSearch(query) {
  const list = document.getElementById('shareAutocompleteList');
  if (!list) return;

  const q = (query || '').trim().toLowerCase();
  if (!q) {
    list.style.display = 'none';
    return;
  }

  const matches = candidateSharers.filter(s => 
    s.name.toLowerCase().includes(q) || 
    s.email.toLowerCase().includes(q) || 
    s.dept.toLowerCase().includes(q)
  );

  if (matches.length === 0) {
    list.innerHTML = `
      <div class="sp-share-auto-item" onclick="addCustomShareRecipient('${query}')">
        <div class="sp-share-auto-avatar" style="background: #64748b;">+</div>
        <div class="sp-share-auto-info">
          <span class="sp-share-auto-name">${query}</span>
          <span class="sp-share-auto-sub">Invite external recipient or email</span>
        </div>
      </div>
    `;
    list.style.display = 'block';
    return;
  }

  list.innerHTML = matches.map(m => `
    <div class="sp-share-auto-item" onclick="selectShareRecipient('${m.name}', '${m.email}')">
      <div class="sp-share-auto-avatar">${m.initials}</div>
      <div class="sp-share-auto-info">
        <span class="sp-share-auto-name">${m.name}</span>
        <span class="sp-share-auto-sub">${m.dept} &bull; ${m.email}</span>
      </div>
    </div>
  `).join('');
  list.style.display = 'block';
}

function selectShareRecipient(name, email) {
  if (!selectedShareRecipients.includes(name)) {
    selectedShareRecipients.push(name);
  }
  renderShareChips();
  const input = document.getElementById('shareRecipientInput');
  if (input) {
    input.value = '';
    input.focus();
  }
  hideShareAutocomplete();
}

function addCustomShareRecipient(val) {
  if (val && !selectedShareRecipients.includes(val)) {
    selectedShareRecipients.push(val);
  }
  renderShareChips();
  const input = document.getElementById('shareRecipientInput');
  if (input) {
    input.value = '';
    input.focus();
  }
  hideShareAutocomplete();
}

function removeShareChip(name) {
  selectedShareRecipients = selectedShareRecipients.filter(r => r !== name);
  renderShareChips();
}

function renderShareChips() {
  const container = document.getElementById('shareRecipientChips');
  if (!container) return;
  container.innerHTML = selectedShareRecipients.map(name => `
    <div class="sp-share-chip">
      <span>${name}</span>
      <button class="sp-share-chip-remove" onclick="removeShareChip('${name}')">&times;</button>
    </div>
  `).join('');
}

function handleShareRecipientKeydown(e) {
  if (e.key === 'Enter') {
    e.preventDefault();
    const input = document.getElementById('shareRecipientInput');
    const val = (input?.value || '').trim();
    if (val) {
      addCustomShareRecipient(val);
    }
  } else if (e.key === 'Backspace' && !e.target.value && selectedShareRecipients.length > 0) {
    selectedShareRecipients.pop();
    renderShareChips();
  }
}

function hideShareAutocomplete() {
  const list = document.getElementById('shareAutocompleteList');
  if (list) list.style.display = 'none';
}

function copyShareModalLink() {
  const doc = currentShareDoc || InnomickData.documents[0];
  const url = `https://innomicktechnologies.sharepoint.com/:w:/r/sites/dms/${encodeURIComponent(doc.name)}?d=${doc.id}&csf=1&web=1`;
  
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(url).catch(() => {});
  }
  
  const textSpan = document.getElementById('shareCopyLinkText');
  if (textSpan) {
    textSpan.innerText = 'Copied!';
    setTimeout(() => {
      textSpan.innerText = 'Copy link';
    }, 2500);
  }

  showToast(`Link copied! Anyone in Innomick Technologies with the link can ${currentSharePermission}.`, 'success');
}

function draftShareMsgWithCopilot() {
  const doc = currentShareDoc || { name: 'document' };
  const textarea = document.getElementById('shareMessageInput');
  if (textarea) {
    textarea.value = `Hi team, please review this document "${doc.name}" at your convenience and let me know your thoughts.`;
    showToast('Copilot drafted a message for you.', 'info');
    textarea.focus();
  }
}

function showShareLinkSettings(e) {
  if (e) e.stopPropagation();
  showToast('Link settings: Anyone in Innomick Technologies with the link can view.', 'info');
}

function showShareHelpInfo(e) {
  if (e) e.stopPropagation();
  showToast('Learn more about sharing files and folders in SharePoint Online.', 'info');
}

function sendShareInvitation() {
  const doc = currentShareDoc || InnomickData.documents[0];
  const input = document.getElementById('shareRecipientInput');
  const typed = (input?.value || '').trim();
  if (typed && !selectedShareRecipients.includes(typed)) {
    selectedShareRecipients.push(typed);
  }

  if (selectedShareRecipients.length === 0) {
    showToast('Please enter at least one recipient name, group, or email.', 'warning');
    if (input) input.focus();
    return;
  }

  // Update doc.sharedWith
  if (!doc.sharedWith) doc.sharedWith = [];
  selectedShareRecipients.forEach(r => {
    if (!doc.sharedWith.includes(r)) {
      doc.sharedWith.push(r);
    }
  });

  const recipientStr = selectedShareRecipients.join(', ');
  showToast(`Link to "${doc.name}" sent to ${recipientStr}!`, 'success');

  closeShareModal();
}


// =============================================================================
// 9. SHARED FILES VIEW
// =============================================================================
function renderSharedFilesView() {
  const filtered = InnomickData.sharedFiles || [];

  return `
    <div class="sp-page-title-bar">
      <h1 class="view-main-title">Shared Files</h1>
    </div>

    <div class="sp-library-subbar">
      <div></div>
      ${renderListToolbarSearch('Search')}
    </div>

    <div class="view-breadcrumbs sp-library-breadcrumb">
      <span style="color: #008272; font-weight: 600;">Home</span>
    </div>

    ${renderSpListTable(filtered, { preset: 'shared', emptyMsg: 'No shared documents found.' })}
  `;
}

// =============================================================================
// 10. CATEGORIES VIEW
// =============================================================================
function renderCategoriesView() {
  return `
    <div class="view-breadcrumbs">
      <a onclick="navigateTo('home')">Home</a> &gt; <span>All Categories</span>
    </div>

    <div class="content-header-row" style="margin-bottom: 20px;">
      <h1 class="view-main-title">Document Categories</h1>
      <div class="search-input-wrapper">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <circle cx="11" cy="11" r="8"/>
          <line x1="21" y1="21" x2="16.65" y2="16.65"/>
        </svg>
        <input type="text" placeholder="Filter categories..." oninput="filterCategoriesGrid(this.value)">
      </div>
    </div>

    <div class="categories-grid" id="allCategoriesGrid">
      ${InnomickData.categories.map(cat => `
        <div class="category-card" onclick="navigateTo('category/${cat.id}')">
          <svg class="folder-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#008272" stroke-width="1.8">
            <path d="M4 20h16a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.93a2 2 0 0 1-1.66-.9l-.82-1.2A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13c0 1.1.9 2 2 2Z"/>
          </svg>
          <span class="category-card-text" title="${cat.name}">${cat.name}</span>
        </div>
      `).join('')}
    </div>
  `;
}

function filterCategoriesGrid(query) {
  const grid = document.getElementById('allCategoriesGrid');
  if (!grid) return;
  const q = query.toLowerCase();
  grid.innerHTML = InnomickData.categories
    .filter(cat => cat.name.toLowerCase().includes(q))
    .map(cat => `
      <div class="category-card" onclick="navigateTo('category/${cat.id}')">
        <svg class="folder-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M4 20h16a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.93a2 2 0 0 1-1.66-.9l-.82-1.2A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13c0 1.1.9 2 2 2Z"/>
        </svg>
        <span class="category-card-text" title="${cat.name}">${cat.name}</span>
        <span class="category-count">${cat.count}</span>
      </div>
    `).join('');
}

// Category Drill-Down View (Inside Category Folder - Matches Exact Live SharePoint Screenshot)
function renderCategoryDrilldownView(catId) {
  const decodedId = decodeURIComponent(catId || '').trim();
  let categoryName = decodedId;
  try {
    const b64 = atob(decodedId);
    if (b64 && /^[\x20-\x7E]+$/.test(b64)) {
      categoryName = b64;
    }
  } catch (e) {}

  let category = InnomickData.categories.find(c => 
    c.id === categoryName || 
    c.name.toLowerCase() === categoryName.toLowerCase() ||
    c.id === decodedId ||
    c.name.toLowerCase() === decodedId.toLowerCase()
  );
  
  if (!category) {
    if (categoryName.toLowerCase().includes('accreditation') || decodedId.toLowerCase().includes('accreditation')) {
      category = InnomickData.categories.find(c => c.name.toLowerCase().includes('accreditation')) || { id: 'cat-1', name: 'Accreditation Documents' };
    } else if (categoryName.toLowerCase().includes('press') || decodedId.toLowerCase().includes('press')) {
      category = InnomickData.categories.find(c => c.id === 'cat-press') || { id: 'cat-press', name: 'Press Releases' };
    } else if (categoryName.toLowerCase().includes('asset') || decodedId.toLowerCase().includes('asset')) {
      category = InnomickData.categories.find(c => c.id === 'cat-3') || { id: 'cat-3', name: 'Asset Inventories' };
    } else if (categoryName.toLowerCase().includes('audit') || decodedId.toLowerCase().includes('audit')) {
      category = InnomickData.categories.find(c => c.id === 'cat-4') || { id: 'cat-4', name: 'Audit Documents' };
    } else {
      category = InnomickData.categories[0];
    }
  }

  // Find documents for this category
  let catDocs = [];
  const query = category.name.toLowerCase();

  catDocs = [...(InnomickData.documents || []), ...(InnomickData.homeAllFiles || [])].filter(d => {
    const docCat = (d.category || '').toLowerCase();
    return docCat === query || docCat.includes(query) || query.includes(docCat);
  });

  // If category has no directly tagged documents, provide related documents from catalog with this category name so the user always sees rich data
  if (catDocs.length === 0) {
    catDocs = (InnomickData.documents || []).slice(0, 7).map((d, idx) => ({
      ...d,
      id: `${category.id || 'CAT'}-${idx + 101}`,
      category: category.name,
      subCategory: d.subCategory || `${category.name} Protocols`,
      referenceNumber: `REF-${category.name.substring(0, 3).toUpperCase()}-${idx + 101}`
    }));
  }

  // De-duplicate documents by name
  const seen = new Set();
  catDocs = catDocs.filter(d => {
    if (seen.has(d.name)) return false;
    seen.add(d.name);
    return true;
  });

  const rowsHtml = catDocs.length > 0 ? catDocs.map(doc => {
    const isDocx = (doc.name || '').endsWith('.docx');
    const isXlsx = (doc.name || '').endsWith('.xlsx');
    const isCsv = (doc.name || '').endsWith('.csv');
    const isPdf = (doc.name || '').endsWith('.pdf');
    const type = isDocx ? 'docx' : isXlsx ? 'xlsx' : isCsv ? 'csv' : isPdf ? 'pdf' : (doc.type || 'docx');

    let deptBadge = '<span style="color: #64748b;">-</span>';
    if (doc.department && doc.department !== '-') {
      if (doc.department === 'Finance & Accounts') {
        deptBadge = `<span style="background: #ea5e5e; color: #ffffff; border-radius: 12px; padding: 2.5px 12px; font-size: 11.5px; font-weight: 500; display: inline-block;">${doc.department}</span>`;
      } else if (doc.department === 'Client Solutions & Delivery') {
        deptBadge = `<span style="background: #008272; color: #ffffff; border-radius: 12px; padding: 2.5px 12px; font-size: 11.5px; font-weight: 500; display: inline-block;">${doc.department}</span>`;
      } else if (doc.department === 'Information Technology') {
        deptBadge = `<span style="background: #107c41; color: #ffffff; border-radius: 12px; padding: 2.5px 12px; font-size: 11.5px; font-weight: 500; display: inline-block;">${doc.department}</span>`;
      } else if (doc.department === 'Corporate Management') {
        deptBadge = `<span style="background: #004d40; color: #ffffff; border-radius: 12px; padding: 2.5px 12px; font-size: 11.5px; font-weight: 500; display: inline-block;">${doc.department}</span>`;
      } else if (doc.department === 'Central Supply Chain Management') {
        deptBadge = `<span style="background: #7c3aed; color: #ffffff; border-radius: 12px; padding: 2.5px 12px; font-size: 11.5px; font-weight: 500; display: inline-block;">${doc.department}</span>`;
      } else {
        deptBadge = `<span style="background: #f1f5f9; color: #475569; border-radius: 12px; padding: 2.5px 12px; font-size: 11.5px; font-weight: 500; display: inline-block;">${doc.department}</span>`;
      }
    }

    const orDash = v => (v && v !== '-' ? v : '-');
    const auditDate = orDash(doc.auditInspectionDate);
    const inspectionArea = orDash(doc.inspectionArea);
    const findingsSummary = orDash(doc.findingsSummary);
    const actionStatus = orDash(doc.actionStatus);
    const reviewDate = orDash(doc.reviewDate);
    const expireDate = orDash(doc.expireDate);
    const subCategory = orDash(doc.subCategory);
    const refNum = orDash(doc.referenceNumber);
    const stakeHolder = orDash(doc.stakeHolder);
    const vendor = orDash(doc.vendor);
    const description = orDash(doc.description);

    return `
      <tr style="border-bottom: 1px solid #f1f5f9; height: 46px; transition: background 0.15s; white-space: nowrap;" onmouseover="this.style.background='#f8fafc'" onmouseout="this.style.background='transparent'">
        <!-- 1. Favorite -->
        <td style="text-align: center; padding: 0 12px; border-right: 1px solid #f1f5f9;">
          <button class="favorite-star-btn ${doc.favorite ? 'starred' : ''}" onclick="toggleFavorite('${doc.id}')" title="${doc.favorite ? 'Unfavorite' : 'Add to Favorites'}" style="background: transparent; border: none; cursor: pointer; font-size: 16px; color: ${doc.favorite ? '#eab308' : '#cbd5e1'};">
            ${doc.favorite ? '&#9733;' : '&#9734;'}
          </button>
        </td>

        <!-- 2. File Name -->
        <td style="padding: 0 12px; border-right: 1px solid #f1f5f9;">
          <div style="display: flex; align-items: center; gap: 8px;">
            <span style="flex-shrink: 0;">${getFileIconSvg(type)}</span>
            <span class="category-file-link" 
                  onclick="toggleDocumentMenu(event, '${doc.id}')" 
                  style="cursor: pointer; color: #0078d4; text-decoration: underline; font-weight: 500;" 
                  title="Open options for ${doc.name}">
              ${doc.name}
            </span>
            <button onclick="openManageAccessModal('${doc.id}')" title="Share" style="background: transparent; border: none; cursor: pointer; color: #a19f9d; display: flex; align-items: center; padding: 2px; margin-left: auto;">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/>
                <polyline points="15 3 21 3 21 9"/>
                <line x1="10" y1="14" x2="21" y2="3"/>
              </svg>
            </button>
          </div>
        </td>

        <!-- 3. Location -->
        <td style="padding: 0 12px; color: #475569; border-right: 1px solid #f1f5f9;">${doc.location || 'My Organization'}</td>

        <!-- 4. Category -->
        <td style="padding: 0 12px; color: #1e293b; font-weight: 500; border-right: 1px solid #f1f5f9;">${doc.category || category.name}</td>

        <!-- 5. Department -->
        <td style="padding: 0 12px; border-right: 1px solid #f1f5f9;">${deptBadge}</td>

        <!-- 6. Reference Number -->
        <td style="padding: 0 12px; color: #008272; font-weight: 600; border-right: 1px solid #f1f5f9;">${refNum}</td>

        <!-- 7. Sub Category -->
        <td style="padding: 0 12px; color: #334155; border-right: 1px solid #f1f5f9;">${subCategory}</td>

        <!-- 8. Description -->
        <td style="padding: 0 12px; color: #475569; max-width: 280px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; border-right: 1px solid #f1f5f9;" title="${description}">${description}</td>

        <!-- 9. Stake Holder -->
        <td style="padding: 0 12px; color: #334155; border-right: 1px solid #f1f5f9;">${stakeHolder}</td>

        <!-- 10. Vendor -->
        <td style="padding: 0 12px; color: #475569; border-right: 1px solid #f1f5f9;">${vendor}</td>

        <!-- 11. Expiration Date -->
        <td style="padding: 0 12px; color: #64748b; border-right: 1px solid #f1f5f9;">${expireDate}</td>

        <!-- 12. Owner -->
        <td style="padding: 0 12px; color: #1e293b; font-weight: 500; border-right: 1px solid #f1f5f9;">${doc.owner || 'Mohammed Wasifuddin'}</td>

        <!-- 13. Upload Date -->
        <td style="padding: 0 12px; color: #64748b; border-right: 1px solid #f1f5f9;">${doc.uploadDate || '-'}</td>

        <!-- 14. Audit/Inspection Date -->
        <td style="padding: 0 12px; color: #64748b; border-right: 1px solid #f1f5f9;">${auditDate}</td>

        <!-- 15. Inspection Area -->
        <td style="padding: 0 12px; color: #334155; border-right: 1px solid #f1f5f9;">${inspectionArea}</td>

        <!-- 16. Findings Summary -->
        <td style="padding: 0 12px; color: #475569; max-width: 260px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; border-right: 1px solid #f1f5f9;" title="${findingsSummary}">${findingsSummary}</td>

        <!-- 17. Action Status -->
        <td style="padding: 0 12px; border-right: 1px solid #f1f5f9;">
          <span style="background: #e6f4f2; color: #008272; border-radius: 12px; padding: 2px 10px; font-size: 11.5px; font-weight: 600; display: inline-block;">
            ${actionStatus}
          </span>
        </td>

        <!-- 18. Review Date -->
        <td style="padding: 0 12px; color: #64748b;">${reviewDate}</td>
      </tr>
    `;
  }).join('') : `
    <tr>
      <td colspan="18" style="text-align: center; padding: 36px; color: #64748b; font-size: 13.5px;">
        No documents found in ${category.name}.
      </td>
    </tr>
  `;

  return `
    <!-- Category Page Header with Name first (No Upload, No Search, matching live SharePoint Screenshot) -->
    <div class="category-drilldown-header" style="display: flex; align-items: center; gap: 12px; margin-bottom: 20px;">
      <button class="cat-circle-back-btn" onclick="navigateTo('categories')" title="Back to Categories" style="background: transparent; border: none; cursor: pointer; display: flex; align-items: center; justify-content: center; padding: 2px; border-radius: 50%; transition: background 0.15s;">
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#008272" stroke-width="2">
          <circle cx="12" cy="12" r="10" stroke="#008272" stroke-width="1.8"/>
          <polyline points="13 8 9 12 13 16" stroke="#008272" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
      </button>
      <h1 class="category-page-title" style="font-size: 21px; font-weight: 600; color: #1e293b; margin: 0;">${category.name}</h1>
    </div>

    <!-- Authentic SharePoint Category Document Table with ALL 18 Metadata Columns -->
    <div class="table-container" style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 4px; overflow-x: auto; box-shadow: 0 1px 3px rgba(0,0,0,0.03);">
      <table class="dms-category-table" style="width: 100%; min-width: 2850px; border-collapse: collapse; font-size: 13px;">
        <thead>
          <tr style="border-bottom: 1px solid #cbd5e1; height: 42px; text-align: left; color: #475569; font-weight: 600; background: #ffffff; white-space: nowrap;">
            <th style="width: 60px; min-width: 60px; padding: 0 12px; text-align: center; border-right: 1px solid #f1f5f9;">Favorite</th>
            <th style="min-width: 280px; padding: 0 12px; border-right: 1px solid #f1f5f9;">File Name</th>
            <th style="min-width: 140px; padding: 0 12px; border-right: 1px solid #f1f5f9;">Location</th>
            <th style="min-width: 180px; padding: 0 12px; border-right: 1px solid #f1f5f9;">Category</th>
            <th style="min-width: 170px; padding: 0 12px; border-right: 1px solid #f1f5f9;">Department</th>
            <th style="min-width: 150px; padding: 0 12px; border-right: 1px solid #f1f5f9;">Reference Number</th>
            <th style="min-width: 160px; padding: 0 12px; border-right: 1px solid #f1f5f9;">Sub Category</th>
            <th style="min-width: 250px; padding: 0 12px; border-right: 1px solid #f1f5f9;">Description</th>
            <th style="min-width: 160px; padding: 0 12px; border-right: 1px solid #f1f5f9;">Stake Holder</th>
            <th style="min-width: 160px; padding: 0 12px; border-right: 1px solid #f1f5f9;">Vendor</th>
            <th style="min-width: 130px; padding: 0 12px; border-right: 1px solid #f1f5f9;">Expiration Date</th>
            <th style="min-width: 150px; padding: 0 12px; border-right: 1px solid #f1f5f9;">Owner</th>
            <th style="min-width: 120px; padding: 0 12px; border-right: 1px solid #f1f5f9;">Upload Date</th>
            <th style="min-width: 160px; padding: 0 12px; border-right: 1px solid #f1f5f9;">Audit/Inspection Date</th>
            <th style="min-width: 170px; padding: 0 12px; border-right: 1px solid #f1f5f9;">Inspection Area</th>
            <th style="min-width: 240px; padding: 0 12px; border-right: 1px solid #f1f5f9;">Findings Summary</th>
            <th style="min-width: 120px; padding: 0 12px; border-right: 1px solid #f1f5f9;">Action Status</th>
            <th style="min-width: 120px; padding: 0 12px;">Review Date</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml}
        </tbody>
      </table>

      <!-- Bottom Pagination (Exact Match with Screenshot 2) -->
      <div class="table-pagination-row" style="display: flex; justify-content: space-between; align-items: center; padding: 12px 18px; border-top: 1px solid #f1f5f9; font-size: 12px; color: #64748b; background: #ffffff;">
        <div style="display: flex; align-items: center; gap: 8px;">
          <button class="pag-btn" style="background: none; border: none; cursor: pointer; color: #64748b; font-size: 12px;">&lt;&lt;</button>
          <button class="pag-btn" style="background: none; border: none; cursor: pointer; color: #64748b; font-size: 12px;">&lt;</button>
          <span style="padding: 0 6px; font-weight: 500; color: #334155;">1/4 pages</span>
          <button class="pag-btn" style="background: none; border: none; cursor: pointer; color: #64748b; font-size: 12px;">&gt;</button>
          <button class="pag-btn" style="background: none; border: none; cursor: pointer; color: #64748b; font-size: 12px;">&gt;&gt;</button>
        </div>
        <div style="display: flex; align-items: center; gap: 6px;">
          <select style="border: 1px solid #cbd5e1; border-radius: 4px; padding: 3px 6px; font-size: 12px; color: #334155; outline: none;">
            <option value="10" selected>10</option>
            <option value="25">25</option>
            <option value="50">50</option>
          </select>
          <span>Items per page</span>
        </div>
      </div>
    </div>
  `;
}

// Helper Functions for File Popovers and Document Navigation
let activePopoverDocId = null;

function toggleDocumentMenu(event, docId) {
  if (event) {
    event.stopPropagation();
    event.preventDefault();
  }

  const popover = document.getElementById('globalFileOpenPopover');
  if (!popover) return;

  // Toggle off if clicking the same document while already open
  if (activePopoverDocId === docId && popover.style.display === 'block') {
    closeAllFilePopovers();
    return;
  }

  closeAllFilePopovers();
  activePopoverDocId = docId;

  const target = event.currentTarget || event.target;
  const rect = target.getBoundingClientRect();

  popover.style.display = 'block';
  popover.style.position = 'fixed';
  popover.style.zIndex = '99999';

  // Position directly below the clicked file name link
  let top = rect.bottom + 4;
  let left = rect.left;

  // Prevent right edge overflow
  const popoverWidth = 200;
  const popoverHeight = 85;
  if (left + popoverWidth > window.innerWidth - 10) {
    left = window.innerWidth - popoverWidth - 12;
  }
  // Prevent bottom edge overflow
  if (top + popoverHeight > window.innerHeight - 10) {
    top = Math.max(10, rect.top - popoverHeight - 4);
  }

  popover.style.top = `${top}px`;
  popover.style.left = `${left}px`;
}

function previewDocument(docId) {
  const doc = findAnyDoc(docId);
  openPreviewModalWithDoc(doc);
}

function handlePopoverOpenInBrowser(event) {
  if (event) event.stopPropagation();
  const docId = activePopoverDocId;
  closeAllFilePopovers();
  if (docId) {
    previewDocument(docId);
  }
}

function handlePopoverOpenInDesktop(event) {
  if (event) event.stopPropagation();
  const docId = activePopoverDocId;
  closeAllFilePopovers();
  if (docId) {
    const doc = findAnyDoc(docId);
    const fileName = doc ? doc.name : 'Document';
    const ext = (fileName.split('.').pop() || '').toLowerCase();
    let appName = 'Microsoft Word';
    let protocol = 'ms-word';
    if (ext === 'xlsx' || ext === 'xls' || ext === 'csv') {
      appName = 'Microsoft Excel';
      protocol = 'ms-excel';
    } else if (ext === 'pptx' || ext === 'ppt') {
      appName = 'Microsoft PowerPoint';
      protocol = 'ms-powerpoint';
    } else if (ext === 'pdf') {
      appName = 'Adobe Acrobat';
      protocol = '';
    }

    showToast(`Opening "${fileName}" in ${appName} desktop application...`, 'info');
    try {
      if (protocol) {
        const dummyUrl = window.location.origin + '/assets/sample.' + ext;
        window.location.href = `${protocol}:ofe|u|${encodeURIComponent(dummyUrl)}`;
      }
    } catch (e) {}
  }
}

function toggleFileOpenPopover(docId, event) {
  toggleDocumentMenu(event, docId);
}

function openDocumentDirect(fileUrlB64, backB64, docId) {
  closeAllFilePopovers();
  if (docId) {
    previewDocument(docId);
  } else {
    window.location.hash = `#/view-document?fileUrl=${fileUrlB64}&back=${backB64}`;
  }
}

function openDocumentInDesktop(fileName, event) {
  if (event) event.stopPropagation();
  const ext = (fileName || '').split('.').pop().toLowerCase();
  const appName = (ext === 'xlsx' || ext === 'xls' || ext === 'csv') ? 'Microsoft Excel' : 'Microsoft Word';
  showToast(`Opening "${fileName}" in ${appName} desktop application...`, 'info');
}

function closeAllFilePopovers() {
  const globalPopover = document.getElementById('globalFileOpenPopover');
  if (globalPopover) {
    globalPopover.style.display = 'none';
  }
  document.querySelectorAll('.sp-file-open-popover').forEach(el => {
    el.style.display = 'none';
  });
  activePopoverDocId = null;
}

function handleViewDocBack(backUrl) {
  closeAllFilePopovers();
  if (backUrl) {
    let target = backUrl;
    if (target.startsWith('#')) target = target.substring(1);
    if (!target.startsWith('/')) target = '/' + target;
    window.location.hash = '#' + target;
  } else {
    window.location.hash = '#/categories';
  }
}

function toggleViewDocProperties() {
  const sidebar = document.getElementById('viewDocDetailsSidebar');
  if (!sidebar) return;
  if (sidebar.style.display === 'none') {
    sidebar.style.display = 'block';
  } else {
    sidebar.style.display = 'none';
  }
}

// =============================================================================
// VIEW DOCUMENT / DETAILS PAGE CONTROLLER
// Routes to: #/view-document?fileUrl=<base64>&back=<base64>
// =============================================================================
function renderViewDocumentView(fileUrlParam, backUrlParam, docIdParam) {
  let fileUrl = '';
  try {
    fileUrl = atob(fileUrlParam);
  } catch (e) {
    fileUrl = decodeURIComponent(fileUrlParam || '');
  }

  let backUrl = '/categories';
  try {
    backUrl = atob(backUrlParam);
  } catch (e) {
    backUrl = decodeURIComponent(backUrlParam || '/categories');
  }

  const rawFileName = fileUrl ? (fileUrl.split('/').pop() || '') : 'Testing Jawad.docx';
  const fileName = decodeURIComponent(rawFileName);
  const ext = (fileName.split('.').pop() || 'docx').toLowerCase();
  const isDocx = ext === 'docx' || ext === 'doc';
  const isPdf = ext === 'pdf';
  const isXlsx = ext === 'xlsx' || ext === 'xls';
  const isCsv = ext === 'csv';

  // Find document in dataset
  let doc = InnomickData.homeAllFiles.find(d => d.name.toLowerCase() === fileName.toLowerCase())
         || InnomickData.recentDocuments.find(d => d.name.toLowerCase() === fileName.toLowerCase())
         || InnomickData.documents.find(d => d.name.toLowerCase() === fileName.toLowerCase())
         || InnomickData.myFiles.find(d => d.name.toLowerCase() === fileName.toLowerCase());

  if (!doc) {
    doc = {
      id: docIdParam || 'ALL-101',
      name: fileName,
      type: ext,
      location: 'My Organization',
      owner: 'MD Faisal Ahmed',
      category: 'Accreditation Documents',
      subCategory: 'ISO Operational Standards',
      department: isPdf ? 'Finance & Accounts' : (fileName.includes('Proposal') ? 'Client Solutions & Delivery' : 'Finance & Accounts'),
      referenceNumber: 'ACC-DOC-101',
      stakeHolder: 'Quality Accreditation Board',
      vendor: 'Innomick Technologies Ltd.',
      uploadDate: '2026-10-11',
      description: 'Operational accreditation testing protocols and operational verification benchmarks',
      version: '1.0',
      favorite: false
    };
  }

  // Department badge
  let deptBadge = '<span style="color: #64748b;">-</span>';
  if (doc.department && doc.department !== '-') {
    if (doc.department === 'Finance & Accounts') {
      deptBadge = `<span style="background: #ea5e5e; color: #ffffff; border-radius: 12px; padding: 2px 10px; font-size: 11.5px; font-weight: 500; display: inline-block;">${doc.department}</span>`;
    } else if (doc.department === 'Client Solutions & Delivery') {
      deptBadge = `<span style="background: #008272; color: #ffffff; border-radius: 12px; padding: 2px 10px; font-size: 11.5px; font-weight: 500; display: inline-block;">${doc.department}</span>`;
    } else {
      deptBadge = `<span style="background: #f1f5f9; color: #475569; border-radius: 12px; padding: 2px 10px; font-size: 11.5px; font-weight: 500; display: inline-block;">${doc.department}</span>`;
    }
  }

  // Generate document preview body
  let sheetContentHtml = '';

  if (isPdf) {
    sheetContentHtml = `
      <div style="border-bottom: 2px solid #dc2626; padding-bottom: 16px; margin-bottom: 24px; display: flex; justify-content: space-between; align-items: flex-start;">
        <div>
          <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 6px;">
            <span style="background: #dc2626; color: #ffffff; font-size: 10px; font-weight: 700; padding: 2px 6px; border-radius: 3px;">PDF</span>
            <span style="font-size: 12px; font-weight: 600; color: #64748b; letter-spacing: 0.5px;">INNOMICK TECHNOLOGIES • ENTERPRISE DMS</span>
          </div>
          <h2 style="font-size: 22px; color: #1e293b; margin: 0 0 6px 0; font-weight: 700;">${doc.name}</h2>
          <p style="font-size: 13px; color: #64748b; margin: 0;">Category: ${doc.category} • Department: ${doc.department}</p>
        </div>
        <div style="text-align: right; font-size: 12px; color: #64748b;">
          <div>Doc Ref: ${doc.referenceNumber !== '-' ? doc.referenceNumber : 'ENG-SOP-2026-08'}</div>
          <div>Version: 1.0 (Official)</div>
          <div>Date: ${doc.uploadDate}</div>
        </div>
      </div>

      <div style="font-size: 14px; line-height: 1.7; color: #334155;">
        <h3 style="font-size: 16px; color: #0f172a; border-left: 4px solid #dc2626; padding-left: 10px; margin: 20px 0 10px 0;">1. Purpose &amp; Administrative Scope</h3>
        <p>This standard operating procedure governs the systematic digital scanning, OCR processing, validation, and electronic filing of all enterprise documentation into the Innomick Technologies Document Management System (DMS).</p>

        <h3 style="font-size: 16px; color: #0f172a; border-left: 4px solid #dc2626; padding-left: 10px; margin: 24px 0 10px 0;">2. Digitization Specifications</h3>
        <table style="width: 100%; border-collapse: collapse; margin: 12px 0 20px 0; font-size: 13px;">
          <thead>
            <tr style="background: #f8fafc; border-bottom: 2px solid #e2e8f0; text-align: left;">
              <th style="padding: 8px 12px;">Parameter</th>
              <th style="padding: 8px 12px;">Requirement</th>
              <th style="padding: 8px 12px;">Standard Verification</th>
            </tr>
          </thead>
          <tbody>
            <tr style="border-bottom: 1px solid #f1f5f9;">
              <td style="padding: 8px 12px; font-weight: 600;">Resolution / DPI</td>
              <td style="padding: 8px 12px;">300 DPI Optical True Color</td>
              <td style="padding: 8px 12px; color: #16a34a;">Compliant (ISO 19005-1)</td>
            </tr>
            <tr style="border-bottom: 1px solid #f1f5f9;">
              <td style="padding: 8px 12px; font-weight: 600;">Format</td>
              <td style="padding: 8px 12px;">Searchable PDF/A-1b with embedded OCR</td>
              <td style="padding: 8px 12px; color: #16a34a;">Compliant</td>
            </tr>
            <tr style="border-bottom: 1px solid #f1f5f9;">
              <td style="padding: 8px 12px; font-weight: 600;">Access Control</td>
              <td style="padding: 8px 12px;">Finance & Accounts &amp; HR Restricted Security Group</td>
              <td style="padding: 8px 12px; color: #16a34a;">Enforced via SharePoint RBAC</td>
            </tr>
          </tbody>
        </table>

        <h3 style="font-size: 16px; color: #0f172a; border-left: 4px solid #dc2626; padding-left: 10px; margin: 24px 0 10px 0;">3. Quality Assurance &amp; Sign-off</h3>
        <p>All scanned records must undergo mandatory dual-custodian review before physical paper originals are transferred to secure off-site climate-controlled archives in accordance with India IT Ministry statutory records retention guidelines.</p>
        
        <div style="margin-top: 36px; padding-top: 20px; border-top: 1px dashed #cbd5e1; display: flex; justify-content: space-between; font-size: 12.5px; color: #64748b;">
          <div>
            <strong>Authorized Custodian:</strong><br>
            ${doc.owner || 'MD Faisal Ahmed'}<br>
            Finance & Accounts Dept.
          </div>
          <div style="text-align: right;">
            <strong>Verification Status:</strong><br>
            <span style="color: #16a34a; font-weight: 600;">&#10003; Electronically Signed &amp; Approved</span><br>
            Innomick DMS Internal Node
          </div>
        </div>
      </div>
    `;
  } else if (isXlsx || isCsv) {
    sheetContentHtml = `
      <div style="border-bottom: 2px solid #16a34a; padding-bottom: 16px; margin-bottom: 20px; display: flex; justify-content: space-between; align-items: flex-start;">
        <div>
          <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 6px;">
            <span style="background: #16a34a; color: #ffffff; font-size: 10px; font-weight: 700; padding: 2px 6px; border-radius: 3px;">${isCsv ? 'CSV' : 'XLSX'}</span>
            <span style="font-size: 12px; font-weight: 600; color: #64748b; letter-spacing: 0.5px;">INNOMICK TECHNOLOGIES • SPREADSHEET VIEWER</span>
          </div>
          <h2 style="font-size: 22px; color: #1e293b; margin: 0 0 6px 0; font-weight: 700;">${doc.name}</h2>
          <p style="font-size: 13px; color: #64748b; margin: 0;">Category: ${doc.category} • Location: ${doc.location}</p>
        </div>
        <div style="text-align: right; font-size: 12px; color: #64748b;">
          <div>Sheet: <strong>Sheet1 (Active)</strong></div>
          <div>Rows: <strong>14 Records</strong></div>
          <div>Modified: ${doc.uploadDate}</div>
        </div>
      </div>

      <!-- Excel formula bar mockup -->
      <div style="display: flex; align-items: center; gap: 10px; background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 4px; padding: 6px 12px; margin-bottom: 16px; font-family: monospace; font-size: 12.5px;">
        <span style="color: #64748b; font-weight: bold;">A1</span>
        <span style="color: #94a3b8;">|</span>
        <span style="color: #0f172a; font-weight: 600;">=ACCREDITATION_COMPLIANCE_MATRIX("${doc.category}")</span>
      </div>

      <!-- Excel grid preview -->
      <div style="overflow-x: auto; border: 1px solid #cbd5e1; border-radius: 4px;">
        <table style="width: 100%; border-collapse: collapse; font-size: 12.5px;">
          <thead>
            <tr style="background: #f1f5f9; color: #475569; font-weight: 600; text-align: left; border-bottom: 1px solid #cbd5e1;">
              <th style="width: 40px; padding: 6px 10px; text-align: center; border-right: 1px solid #cbd5e1;">#</th>
              <th style="padding: 6px 12px; border-right: 1px solid #cbd5e1;">A (Record ID)</th>
              <th style="padding: 6px 12px; border-right: 1px solid #cbd5e1;">B (Audit Item)</th>
              <th style="padding: 6px 12px; border-right: 1px solid #cbd5e1;">C (Responsible Dept)</th>
              <th style="padding: 6px 12px; border-right: 1px solid #cbd5e1;">D (Compliance %)</th>
              <th style="padding: 6px 12px;">E (Status)</th>
            </tr>
          </thead>
          <tbody>
            <tr style="border-bottom: 1px solid #f1f5f9;">
              <td style="padding: 6px 10px; text-align: center; background: #f8fafc; color: #64748b; border-right: 1px solid #cbd5e1;">1</td>
              <td style="padding: 6px 12px; border-right: 1px solid #cbd5e1; font-weight: 500;">REC-2026-001</td>
              <td style="padding: 6px 12px; border-right: 1px solid #cbd5e1;">ISO 27001 Cloud Security Audit</td>
              <td style="padding: 6px 12px; border-right: 1px solid #cbd5e1;">Finance & Accounts</td>
              <td style="padding: 6px 12px; border-right: 1px solid #cbd5e1; color: #16a34a; font-weight: 600;">98.4%</td>
              <td style="padding: 6px 12px;"><span style="background: #dcfce7; color: #15803d; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600;">Approved</span></td>
            </tr>
            <tr style="border-bottom: 1px solid #f1f5f9; background: #fafafa;">
              <td style="padding: 6px 10px; text-align: center; background: #f8fafc; color: #64748b; border-right: 1px solid #cbd5e1;">2</td>
              <td style="padding: 6px 12px; border-right: 1px solid #cbd5e1; font-weight: 500;">REC-2026-002</td>
              <td style="padding: 6px 12px; border-right: 1px solid #cbd5e1;">Ministry of Health License Renewal</td>
              <td style="padding: 6px 12px; border-right: 1px solid #cbd5e1;">Finance & Accounts</td>
              <td style="padding: 6px 12px; border-right: 1px solid #cbd5e1; color: #16a34a; font-weight: 600;">100.0%</td>
              <td style="padding: 6px 12px;"><span style="background: #dcfce7; color: #15803d; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600;">Approved</span></td>
            </tr>
            <tr style="border-bottom: 1px solid #f1f5f9;">
              <td style="padding: 6px 10px; text-align: center; background: #f8fafc; color: #64748b; border-right: 1px solid #cbd5e1;">3</td>
              <td style="padding: 6px 12px; border-right: 1px solid #cbd5e1; font-weight: 500;">REC-2026-003</td>
              <td style="padding: 6px 12px; border-right: 1px solid #cbd5e1;">Operational Governance Protocol SCN</td>
              <td style="padding: 6px 12px; border-right: 1px solid #cbd5e1;">Test SCN</td>
              <td style="padding: 6px 12px; border-right: 1px solid #cbd5e1; color: #2563eb; font-weight: 600;">94.2%</td>
              <td style="padding: 6px 12px;"><span style="background: #e0f2fe; color: #0369a1; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600;">Verified</span></td>
            </tr>
            <tr style="border-bottom: 1px solid #f1f5f9; background: #fafafa;">
              <td style="padding: 6px 10px; text-align: center; background: #f8fafc; color: #64748b; border-right: 1px solid #cbd5e1;">4</td>
              <td style="padding: 6px 12px; border-right: 1px solid #cbd5e1; font-weight: 500;">REC-2026-004</td>
              <td style="padding: 6px 12px; border-right: 1px solid #cbd5e1;">Insurance Claims Audit Reconciliation</td>
              <td style="padding: 6px 12px; border-right: 1px solid #cbd5e1;">Finance & Accounts</td>
              <td style="padding: 6px 12px; border-right: 1px solid #cbd5e1; color: #16a34a; font-weight: 600;">99.1%</td>
              <td style="padding: 6px 12px;"><span style="background: #dcfce7; color: #15803d; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600;">Approved</span></td>
            </tr>
          </tbody>
        </table>
      </div>
    `;
  } else {
    // Word Document preview (.docx, default)
    sheetContentHtml = `
      <div style="border-bottom: 2px solid #008272; padding-bottom: 20px; margin-bottom: 24px; display: flex; justify-content: space-between; align-items: flex-start;">
        <div>
          <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 6px;">
            <span style="background: #008272; color: #ffffff; font-size: 10px; font-weight: 700; padding: 2px 6px; border-radius: 3px;">DOCX</span>
            <span style="font-size: 12px; font-weight: 600; color: #64748b; letter-spacing: 0.5px;">INNOMICK TECHNOLOGIES • ENTERPRISE DMS</span>
          </div>
          <h2 style="font-size: 24px; color: #1e293b; margin: 0 0 6px 0; font-weight: 700;">${doc.name}</h2>
          <p style="font-size: 13.5px; color: #64748b; margin: 0;">Organization Category: <strong>${doc.category}</strong></p>
        </div>
        <div style="text-align: right; font-size: 12px; color: #64748b; background: #f8fafc; border: 1px solid #e2e8f0; padding: 8px 14px; border-radius: 6px;">
          <div>Classification: <strong style="color: #008272;">Confidential</strong></div>
          <div>Version: <strong>${doc.version || '1.0'}</strong></div>
          <div>Location: <strong>${doc.location || 'My Organization'}</strong></div>
          <div>Last Modified: <strong>${doc.uploadDate || '2026-10-18'}</strong></div>
        </div>
      </div>

      <div style="font-size: 14px; line-height: 1.75; color: #334155;">
        <h3 style="font-size: 16px; color: #0f172a; border-left: 4px solid #008272; padding-left: 10px; margin: 20px 0 10px 0;">1. Purpose &amp; Scope of Enterprise Architecture</h3>
        <p>This verification document, filed under <strong>${doc.category}</strong>, outlines the procedural architecture requirements and compliance verification benchmarks for cloud operations, enterprise security, and software engineering governance within Innomick Technologies.</p>

        <h3 style="font-size: 16px; color: #0f172a; border-left: 4px solid #008272; padding-left: 10px; margin: 24px 0 10px 0;">2. Departmental Responsibility</h3>
        <p>Primary Department: <strong>${doc.department || 'Engineering & Cloud Architecture'}</strong>. The designated document owner and administrative custodian is <strong>${doc.owner || 'Mohammed Wasifuddin'}</strong>. Any revisions or amendments must be submitted through the standard DMS Approval Workflow prior to publication.</p>

        <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 16px 20px; margin: 20px 0;">
          <h4 style="margin: 0 0 10px 0; color: #0f172a; font-size: 14px;">Operational Verification Checklist:</h4>
          <ul style="margin: 0; padding-left: 20px; color: #475569;">
            <li>ISO/IEC 27001 & SOC 2 Type II Enterprise Standards Alignment.</li>
            <li>Cloud Security Alliance (CSA) Cloud Controls Matrix Compliance Checklist.</li>
            <li>Zero Trust Network Architecture & API Gateway Authentication Standards.</li>
            <li>Client Data Privacy & Confidential Cloud Infrastructure Security Controls.</li>
          </ul>
        </div>

        <h3 style="font-size: 16px; color: #0f172a; border-left: 4px solid #008272; padding-left: 10px; margin: 24px 0 10px 0;">3. Document Governance &amp; Audit Sign-Off</h3>
        <p>This document is stored securely in SharePoint Document Management System at the path <code>${fileUrl || 'https://innomick.sharepoint.com/sites/DMS/My Organization/' + doc.name}</code>. Version control is automatically maintained with cryptographic tamper checks.</p>

        <div style="margin-top: 40px; padding-top: 20px; border-top: 1px solid #e2e8f0; display: flex; justify-content: space-between; font-size: 13px; color: #64748b;">
          <div>
            <strong>Prepared by:</strong><br>
            ${doc.owner || 'Jawad Shaikh'}<br>
            Quality &amp; Accreditation Lead
          </div>
          <div style="text-align: right;">
            <strong>Approved by:</strong><br>
            Finance & Accounts Governance Board<br>
            Status: <span style="color: #008272; font-weight: 600;">Active &amp; Distributed</span>
          </div>
        </div>
      </div>
    `;
  }

  return `
    <div class="view-doc-page">
      <!-- SharePoint Online Command / Action Bar -->
      <div class="view-doc-commandbar">
        <div class="view-doc-info-left">
          <button onclick="handleViewDocBack('${backUrl}')" title="Back to Category" style="background: transparent; border: none; cursor: pointer; display: flex; align-items: center; justify-content: center; padding: 2px; border-radius: 50%; transition: background 0.15s;">
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#008272" stroke-width="2">
              <circle cx="12" cy="12" r="10" stroke="#008272" stroke-width="1.8"/>
              <polyline points="13 8 9 12 13 16" stroke="#008272" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>
            </svg>
          </button>
          <span style="display: flex; align-items: center; gap: 8px;">
            ${getFileIconSvg(ext)}
            <h1 class="view-doc-title">${doc.name}</h1>
          </span>
          <span class="view-doc-status-badge">Saved</span>
        </div>

        <div class="view-doc-actions-right">
          <button class="view-doc-btn primary" onclick="downloadDocument('${doc.id || doc.name}')" title="Download copy to your computer">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
              <polyline points="7 10 12 15 17 10"/>
              <line x1="12" y1="15" x2="12" y2="3"/>
            </svg>
            <span>Download</span>
          </button>

          <button class="view-doc-btn" onclick="openManageAccessModal('${doc.id}')" title="Share with colleagues">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M4 12v8a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-8"/>
              <polyline points="16 6 12 2 8 6"/>
              <line x1="12" y1="2" x2="12" y2="15"/>
            </svg>
            <span>Share</span>
          </button>

          <button class="view-doc-btn" onclick="openDocumentInDesktop('${doc.name}', event)" title="Open in Microsoft Office desktop application">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <rect x="2" y="3" width="20" height="14" rx="2" ry="2"/>
              <line x1="8" y1="21" x2="16" y2="21"/>
              <line x1="12" y1="17" x2="12" y2="21"/>
            </svg>
            <span>Open in Desktop App</span>
          </button>

          <button class="view-doc-btn" onclick="openVersionHistoryModal('${doc.id}')" title="View Version History">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <circle cx="12" cy="12" r="10"/>
              <polyline points="12 6 12 12 16 14"/>
            </svg>
            <span>Version History</span>
          </button>

          <button class="view-doc-btn" onclick="toggleViewDocProperties()" title="Toggle Document Properties Details">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <circle cx="12" cy="12" r="10"/>
              <line x1="12" y1="16" x2="12" y2="12"/>
              <line x1="12" y1="8" x2="12.01" y2="8"/>
            </svg>
            <span>Details</span>
          </button>
        </div>
      </div>

      <!-- Split Layout: Viewer Canvas + Collapsible Details Sidebar -->
      <div class="view-doc-layout">
        <!-- Center Canvas -->
        <div class="view-doc-canvas-area">
          <div class="view-doc-ribbon">
            <div style="display: flex; align-items: center; gap: 12px;">
              <span>Viewing in Browser</span>
              <span style="color: #cbd5e1;">|</span>
              <button onclick="showToast('Zoom reduced')" style="background: none; border: none; cursor: pointer; color: #4b5563; font-weight: bold; font-size: 14px;">-</button>
              <span style="font-weight: 600;">100%</span>
              <button onclick="showToast('Zoom increased')" style="background: none; border: none; cursor: pointer; color: #4b5563; font-weight: bold; font-size: 14px;">+</button>
            </div>
            <div style="display: flex; align-items: center; gap: 16px;">
              <span>Page 1 of 1</span>
              <span style="color: #cbd5e1;">|</span>
              <button onclick="window.print()" style="background: none; border: none; cursor: pointer; color: #4b5563; display: flex; align-items: center; gap: 4px; font-size: 12px;">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                  <polyline points="6 9 6 2 18 2 18 9"/>
                  <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/>
                  <rect x="6" y="14" width="12" height="8"/>
                </svg>
                Print
              </button>
            </div>
          </div>

          <div class="view-doc-sheet">
            ${sheetContentHtml}
          </div>
        </div>

        <!-- Right Properties Sidebar (Displaying all 10 schema fields with complete data) -->
        <aside id="viewDocDetailsSidebar" class="view-doc-sidebar">
          <div style="display: flex; align-items: center; justify-content: space-between; border-bottom: 1px solid #edebe9; padding-bottom: 12px; margin-bottom: 16px;">
            <h3 style="font-size: 15px; font-weight: 600; color: #201f1e; margin: 0;">Document Details</h3>
            <button onclick="toggleViewDocProperties()" style="background: transparent; border: none; cursor: pointer; color: #64748b; font-size: 16px; padding: 4px;">&times;</button>
          </div>

          <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 14px; text-align: center; margin-bottom: 20px;">
            <div style="display: flex; justify-content: center; margin-bottom: 8px;">
              ${getFileIconSvg(ext)}
            </div>
            <div style="font-weight: 600; color: #1e293b; font-size: 13.5px; word-break: break-all;">${doc.name}</div>
            <div style="color: #64748b; font-size: 12px; margin-top: 4px;">${ext.toUpperCase()} • 28.4 KB</div>
          </div>

          <div style="font-size: 13px; display: flex; flex-direction: column; gap: 14px;">
            <div>
              <div style="color: #64748b; font-size: 11.5px; text-transform: uppercase; font-weight: 600; margin-bottom: 3px;">Location</div>
              <div style="color: #1e293b; font-weight: 500;">${doc.location || 'My Organization'}</div>
            </div>

            <div>
              <div style="color: #64748b; font-size: 11.5px; text-transform: uppercase; font-weight: 600; margin-bottom: 3px;">Owner / Created By</div>
              <div style="color: #1e293b; font-weight: 500;">${doc.owner || 'MD Faisal Ahmed'}</div>
            </div>

            <div>
              <div style="color: #64748b; font-size: 11.5px; text-transform: uppercase; font-weight: 600; margin-bottom: 3px;">Category</div>
              <div style="color: #1e293b; font-weight: 500;">${doc.category || 'Accreditation Documents'}</div>
            </div>

            <div>
              <div style="color: #64748b; font-size: 11.5px; text-transform: uppercase; font-weight: 600; margin-bottom: 3px;">Sub Category</div>
              <div style="color: #1e293b; font-weight: 500;">${doc.subCategory && doc.subCategory !== '-' ? doc.subCategory : 'ISO Operational Standards'}</div>
            </div>

            <div>
              <div style="color: #64748b; font-size: 11.5px; text-transform: uppercase; font-weight: 600; margin-bottom: 3px;">Department</div>
              <div>${deptBadge}</div>
            </div>

            <div>
              <div style="color: #64748b; font-size: 11.5px; text-transform: uppercase; font-weight: 600; margin-bottom: 3px;">Reference Number</div>
              <div style="color: #008272; font-weight: 600;">${doc.referenceNumber && doc.referenceNumber !== '-' ? doc.referenceNumber : 'ACC-DOC-101'}</div>
            </div>

            <div>
              <div style="color: #64748b; font-size: 11.5px; text-transform: uppercase; font-weight: 600; margin-bottom: 3px;">Stake Holder</div>
              <div style="color: #1e293b; font-weight: 500;">${doc.stakeHolder && doc.stakeHolder !== '-' ? doc.stakeHolder : 'Quality Accreditation Board'}</div>
            </div>

            <div>
              <div style="color: #64748b; font-size: 11.5px; text-transform: uppercase; font-weight: 600; margin-bottom: 3px;">Vendor</div>
              <div style="color: #1e293b; font-weight: 500;">${doc.vendor && doc.vendor !== '-' ? doc.vendor : 'Innomick Technologies Ltd.'}</div>
            </div>

            <div>
              <div style="color: #64748b; font-size: 11.5px; text-transform: uppercase; font-weight: 600; margin-bottom: 3px;">Upload Date</div>
              <div style="color: #1e293b; font-weight: 500;">${doc.uploadDate || '2026-10-11'}</div>
            </div>

            <div>
              <div style="color: #64748b; font-size: 11.5px; text-transform: uppercase; font-weight: 600; margin-bottom: 3px;">Description</div>
              <div style="color: #334155; line-height: 1.45; font-size: 12.5px;">${doc.description && doc.description !== '-' ? doc.description : 'Operational accreditation testing protocols and operational verification benchmarks.'}</div>
            </div>

            <div>
              <div style="color: #64748b; font-size: 11.5px; text-transform: uppercase; font-weight: 600; margin-bottom: 3px;">Version</div>
              <div style="color: #1e293b; font-weight: 600;">v${doc.version || '1.0'} (Current)</div>
            </div>

            <div style="border-top: 1px solid #edebe9; padding-top: 14px; margin-top: 4px;">
              <div style="color: #64748b; font-size: 11.5px; text-transform: uppercase; font-weight: 600; margin-bottom: 4px;">SharePoint URL</div>
              <div style="font-size: 11px; color: #0284c7; word-break: break-all; background: #f0f9ff; padding: 6px 8px; border-radius: 4px; border: 1px solid #bae6fd;">
                ${fileUrl || 'https://innomick.sharepoint.com/sites/DMS/My Organization/' + doc.name}
              </div>
              <button onclick="navigator.clipboard && navigator.clipboard.writeText('${fileUrl}'); showToast('SharePoint document URL copied to clipboard');" style="margin-top: 6px; width: 100%; padding: 6px 0; background: #ffffff; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 12px; cursor: pointer; color: #334155;">
                Copy SharePoint URL
              </button>
            </div>
          </div>
        </aside>
      </div>
    </div>
  `;
}

// =============================================================================
// 11. EXPIRING FILES VIEW
// =============================================================================
function renderExpiringFilesView() {
  const filtered = InnomickData.expiringFiles || [];

  return `
    <div class="sp-page-title-bar">
      <h1 class="view-main-title">Expiring Files</h1>
    </div>

    <div class="sp-library-subbar">
      <div></div>
      ${renderListToolbarSearch('Search')}
    </div>

    <div class="view-breadcrumbs sp-library-breadcrumb">
      <span style="color: #008272; font-weight: 600;">Home</span>
    </div>

    ${renderSpListTable(filtered, { preset: 'expiring', emptyMsg: 'No expiring files found.' })}
  `;
}

// =============================================================================
// 12. RECYCLE BIN VIEW
// =============================================================================
function renderRecycleBinView() {
  const hasSelection = AppState.selectedRecycleBinItems.size > 0;

  return `
    <div class="content-header-row">
      <h1 class="view-main-title">Recycle Bin</h1>
      <button class="icon-action-btn" title="Refresh" onclick="refreshData()">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/>
          <path d="M21 3v5h-5"/>
          <path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/>
          <path d="M8 16H3v5"/>
        </svg>
      </button>
    </div>

    <div style="display: flex; gap: 10px; margin-bottom: 16px;">
      <button class="btn-secondary" ${hasSelection ? '' : 'disabled'} onclick="restoreSelectedRecycleItems()">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <polyline points="1 4 1 10 7 10"/>
          <path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10"/>
        </svg>
        Restore
      </button>
      <button class="btn-danger-outline" ${hasSelection ? '' : 'disabled'} onclick="deletePermanentlySelectedRecycleItems()">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <polyline points="3 6 5 6 21 6"/>
          <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
        </svg>
        Delete
      </button>
    </div>

    ${InnomickData.recycleBin.length > 0 ? `
      <div class="table-card-container">
        <table class="dms-data-table">
          <thead>
            <tr class="header-row">
              <th style="width: 36px;">
                <input type="checkbox" class="select-row-cb" onclick="toggleSelectAllRecycle(this)">
              </th>
              <th>Name</th>
              <th>Deleted</th>
              <th>Location</th>
            </tr>
          </thead>
          <tbody>
            ${InnomickData.recycleBin.map(item => `
              <tr>
                <td style="text-align: center;">
                  <input type="checkbox" class="select-row-cb" ${AppState.selectedRecycleBinItems.has(item.id) ? 'checked' : ''} onchange="toggleRecycleItemSelection('${item.id}', this.checked)">
                </td>
                <td>
                  <div class="file-name-cell">
                    ${getFileIconSvg(item.type)}
                    <span>${item.name}</span>
                  </div>
                </td>
                <td>${item.deletedDate}</td>
                <td style="color: var(--gray-500); font-size: 12px;">${item.originalLocation}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    ` : `
      <div class="empty-state-card">
        <p style="color: var(--gray-500); font-size: 14px;">The Recycle Bin is empty.</p>
      </div>
    `}
  `;
}

function toggleRecycleItemSelection(id, checked) {
  if (checked) {
    AppState.selectedRecycleBinItems.add(id);
  } else {
    AppState.selectedRecycleBinItems.delete(id);
  }
  renderCurrentView();
}

function toggleSelectAllRecycle(masterCb) {
  if (masterCb.checked) {
    InnomickData.recycleBin.forEach(item => AppState.selectedRecycleBinItems.add(item.id));
  } else {
    AppState.selectedRecycleBinItems.clear();
  }
  renderCurrentView();
}

function restoreSelectedRecycleItems() {
  const ids = Array.from(AppState.selectedRecycleBinItems);
  ids.forEach(id => {
    const idx = InnomickData.recycleBin.findIndex(i => String(i.id) === String(id));
    if (idx !== -1) {
      const restored = InnomickData.recycleBin.splice(idx, 1)[0];
      const docToRestore = restored._originalDoc || {
        id: 'doc-restored-' + Date.now(),
        name: restored.name,
        type: restored.type,
        category: 'General Correspondence',
        subCategory: 'Restored Documents',
        department: restored._department || InnomickData.currentUser.department,
        owner: InnomickData.currentUser.name,
        status: 'Active',
        uploadDate: 'Just now',
        expireDate: '11 October 2027',
        description: 'Restored from recycle bin',
        size: restored.size,
        location: restored.originalLocation || 'My Organization',
        favorite: false,
        isCheckedOut: false,
        checkedOutBy: ""
      };

      if (restored._sourceLibrary === 'my-files') {
        if (!InnomickData.myFiles) InnomickData.myFiles = [];
        InnomickData.myFiles.unshift(docToRestore);
      } else if (restored._sourceLibrary === 'department' && restored._branch && restored._department) {
        if (!InnomickData.departmentSites[restored._branch]) InnomickData.departmentSites[restored._branch] = {};
        if (!InnomickData.departmentSites[restored._branch][restored._department]) InnomickData.departmentSites[restored._branch][restored._department] = [];
        InnomickData.departmentSites[restored._branch][restored._department].unshift(docToRestore);
      } else {
        InnomickData.documents.unshift(docToRestore);
      }
    }
  });
  AppState.selectedRecycleBinItems.clear();
  showToast(`Successfully restored ${ids.length} document(s) back to original location.`);
  renderCurrentView();
}

function deletePermanentlySelectedRecycleItems() {
  const count = AppState.selectedRecycleBinItems.size;
  InnomickData.recycleBin = InnomickData.recycleBin.filter(i => !AppState.selectedRecycleBinItems.has(i.id));
  AppState.selectedRecycleBinItems.clear();
  showToast(`Permanently deleted ${count} item(s).`, 'warning');
  renderCurrentView();
}

// =============================================================================
// 13. INTERACTIVE ACTIONS (Favorites, Search, Upload, Workflow, Approvals)
// =============================================================================
function toggleFavorite(id) {
  const doc = findAnyDoc(id);
  if (!doc) return;
  const newFavStatus = !doc.favorite;
  doc.favorite = newFavStatus;

  // Synchronize across all collections
  const syncDocFav = (item) => {
    if (item && String(item.id) === String(id)) item.favorite = newFavStatus;
  };

  (InnomickData.documents || []).forEach(syncDocFav);
  (InnomickData.recentDocuments || []).forEach(syncDocFav);
  (InnomickData.homeAllFiles || []).forEach(syncDocFav);
  (InnomickData.allFiles || []).forEach(syncDocFav);
  (InnomickData.myFiles || []).forEach(syncDocFav);
  (InnomickData.sharedFiles || []).forEach(syncDocFav);
  (InnomickData.expiringFiles || []).forEach(syncDocFav);

  if (InnomickData.departmentSites) {
    for (const b in InnomickData.departmentSites) {
      for (const d in InnomickData.departmentSites[b]) {
        (InnomickData.departmentSites[b][d] || []).forEach(syncDocFav);
      }
    }
  }

  closeContextMenu();
  showToast(`${newFavStatus ? 'Added to' : 'Removed from'} Favorites: ${doc.name}`);
  renderCurrentView();
}

function toggleMyFileFavorite(id) {
  const doc = InnomickData.myFiles.find(d => d.id === id);
  if (doc) {
    doc.favorite = !doc.favorite;
    showToast(`${doc.favorite ? 'Added to' : 'Removed from'} Favorites: ${doc.name}`);
    renderCurrentView();
  }
}

function handleSearch(query, inputEl) {
  AppState.globalSearchQuery = query;
  AppState.listPage = 1;
  AppState.favoritesPage = 1;
  window._lastSearchFocus = {
    id: inputEl && inputEl.id ? inputEl.id : 'myFilesSearchInput',
    cursor: inputEl && typeof inputEl.selectionStart === 'number' ? inputEl.selectionStart : (query ? query.length : 0)
  };
  window._lastFilterFocus = null;
  renderCurrentView();
}

function handleGlobalSearch(query) {
  AppState.globalSearchQuery = query;
  if (AppState.currentRoute !== 'home' && AppState.currentRoute !== 'organization') {
    navigateTo('organization');
  } else {
    renderCurrentView();
  }
}

function refreshData() {
  showToast('Refreshing DMS state from SharePoint list cache...');
  renderCurrentView();
}

// =============================================================================
// CONTEXT-AWARE 2-STEP UPLOAD & WORKFLOW CONTROLLER (Matches Video Recording)
// =============================================================================
AppState.currentUploadSource = 'my-department';
AppState.pendingUploadData = null;
AppState.selectedUploadWorkflowType = 'individual';
AppState.isDirectUploadMode = false;
AppState.directUploadSelectedFolder = null;
AppState.directUploadTempFolder = null;
AppState.directUploadFilterQuery = '';

// Trigger Direct Upload Flow
function triggerDirectUpload(event, sourceContext) {
  if (event) {
    event.preventDefault();
    event.stopPropagation();
  }
  AppState.isDirectUploadMode = true;
  AppState.directUploadSelectedFolder = (AppState.activeDeptFolder && AppState.activeDeptFolder !== 'Direct Uploads') ? AppState.activeDeptFolder : null;
  AppState.directUploadTempFolder = AppState.directUploadSelectedFolder;
  AppState.directUploadFilterQuery = '';
  triggerDeviceUpload(event, sourceContext || 'my-department');
}

// Global trigger for local device file upload
function triggerDeviceUpload(event, sourceContext, defaultCategory) {
  if (event) {
    event.preventDefault();
    event.stopPropagation();
  }
  
  // Set context
  AppState.currentUploadSource = sourceContext || (AppState.currentRoute === 'my-files' ? 'my-files' : AppState.currentRoute === 'organization' ? 'my-organization' : 'my-department');
  AppState.pendingUploadDefaultCategory = defaultCategory || null;
  AppState.attachedLocalFile = null;
  AppState.pendingUploadData = null;

  const fileInput = document.getElementById('mainDeviceFileInput');
  if (fileInput) {
    fileInput.value = '';
    
    // When the user cancels out of the OS dialog without choosing, open the modal anyway so they are not blocked
    const handleCancel = () => {
      openUploadModal(AppState.pendingUploadDefaultCategory, AppState.currentUploadSource);
      fileInput.removeEventListener('cancel', handleCancel);
    };
    fileInput.addEventListener('cancel', handleCancel, { once: true });
    
    // Trigger the native OS file browser
    fileInput.click();
    setTimeout(() => {
      const modal = document.getElementById('uploadModal');
      if (modal && !modal.classList.contains('active')) {
        openUploadModal(AppState.pendingUploadDefaultCategory, AppState.currentUploadSource);
      }
    }, 200);
  } else {
    openUploadModal(AppState.pendingUploadDefaultCategory, AppState.currentUploadSource);
  }
}

function handleMainDeviceFileSelected(input) {
  if (input.files && input.files[0]) {
    const file = input.files[0];
    const nameEl = document.getElementById('uploadAttachedFileName');
    const iconSpan = document.getElementById('uploadAttachedFileIcon');
    if (nameEl) {
      nameEl.textContent = file.name;
    }
    if (iconSpan) {
      const ext = file.name.split('.').pop().toLowerCase();
      iconSpan.innerHTML = getFileIconSvg(ext);
    }
    AppState.attachedLocalFile = file;
    showToast(`File "${file.name}" selected from device.`, 'info');
  }
  // Open the Upload Document details modal now that the file has been chosen from device
  openUploadModal(AppState.pendingUploadDefaultCategory, AppState.currentUploadSource);
}

function openUploadModal(defaultCategory, sourceContext) {
  const modal = document.getElementById('uploadModal');
  if (!modal) return;

  // Determine upload source context: 'my-files', 'my-department', 'my-organization'
  if (sourceContext) {
    AppState.currentUploadSource = sourceContext;
  } else {
    if (AppState.currentRoute === 'my-files') AppState.currentUploadSource = 'my-files';
    else if (AppState.currentRoute === 'organization') AppState.currentUploadSource = 'my-organization';
    else AppState.currentUploadSource = 'my-department';
  }

  // Update workflow footer buttons based on source
  updateStep2WorkflowFooter();

  // Reset all containers
  const step1Form = document.getElementById('uploadStep1Form');
  const step2Cont = document.getElementById('uploadStep2Container');
  const dirFolderCont = document.getElementById('uploadDirectFolderContainer');
  const dirConfirmCont = document.getElementById('uploadDirectConfirmContainer');
  const mainTitle = document.getElementById('uploadModalMainTitle');
  const subTitle = document.getElementById('uploadModalSubTitle');
  const backBtn = document.getElementById('uploadModalBackBtn');
  const dialog = document.querySelector('#uploadModal .dms-modal-dialog');
  
  if (step1Form) step1Form.style.display = 'block';
  if (step2Cont) step2Cont.style.display = 'none';
  if (dirFolderCont) dirFolderCont.style.display = 'none';
  if (dirConfirmCont) dirConfirmCont.style.display = 'none';
  if (subTitle) subTitle.style.display = 'none';
  if (backBtn) backBtn.style.display = 'none';
  if (dialog) dialog.style.maxWidth = '620px';

  if (AppState.isDirectUploadMode) {
    if (mainTitle) mainTitle.innerText = 'Quick Document Upload';
  } else {
    if (mainTitle) mainTitle.innerText = 'Upload Document';
  }

  const permRow = document.getElementById('uploadFilePermissionRow');
  if (permRow) permRow.style.display = 'none';

  // Reset standard text inputs for fresh document upload
  const subCatInput = document.getElementById('uploadSubCategory');
  const descInput = document.getElementById('uploadDescription');
  const stakeInput = document.getElementById('uploadStakeholder');
  if (subCatInput) subCatInput.value = '';
  if (descInput) descInput.value = '';
  if (stakeInput) stakeInput.value = '';

  const nameEl = document.getElementById('uploadAttachedFileName');
  const iconSpan = document.getElementById('uploadAttachedFileIcon');
  if (!AppState.attachedLocalFile) {
    AppState.attachedLocalFile = { name: 'Imran-Portal-Flow-Migration-Review (2).docx', size: '1.2 MB' };
    if (nameEl) nameEl.textContent = 'Imran-Portal-Flow-Migration-Review (2).docx';
    if (iconSpan) iconSpan.innerHTML = getFileIconSvg('docx');
  }

  // Populate Categories
  populateUploadCategoryOptions();
  const catSelect = document.getElementById('uploadCategorySelect');
  if (catSelect) {
    catSelect.value = defaultCategory || '';
  }
  handleUploadCategoryChange(defaultCategory || '');

  // Populate User Context
  const ownerEl = document.getElementById('uploadOwnerName');
  const deptEl = document.getElementById('uploadDeptBadge');
  const dateEl = document.getElementById('uploadDateDisplay');
  if (ownerEl) ownerEl.innerText = InnomickData.currentUser.name || 'Mohammed Wasifuddin';
  if (deptEl) deptEl.innerText = AppState.activeDepartment || 'Information Technology';
  if (dateEl) {
    const now = new Date();
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    dateEl.innerText = `${now.getDate()} ${months[now.getMonth()]} ${now.getFullYear()}`;
  }

  modal.classList.add('active');
}

function closeUploadModal() {
  const modal = document.getElementById('uploadModal');
  if (modal) modal.classList.remove('active');
  closeDirectUploadFolderDrawer();
  AppState.isDirectUploadMode = false;
  AppState.directUploadSelectedFolder = null;
  AppState.directUploadTempFolder = null;
  AppState.attachedLocalFile = null;
  backToUploadStep1();
}

function handleFileChosen(input) {
  if (input.files && input.files[0]) {
    const file = input.files[0];
    const nameEl = document.getElementById('uploadAttachedFileName');
    const iconSpan = document.getElementById('uploadAttachedFileIcon');
    if (nameEl) {
      nameEl.textContent = file.name;
    }
    if (iconSpan) {
      const ext = file.name.split('.').pop().toLowerCase();
      iconSpan.innerHTML = getFileIconSvg(ext);
    }
    AppState.attachedLocalFile = file;
    showToast(`File "${file.name}" attached.`, 'info');
  }
}

// Category-specific upload fields, captured from the live DMS upload form.
// Every field is required; `key` is the document property / list column it fills.
const F = (key, label, type, options) => ({ key, label, type, options });
const UPLOAD_FIELD_SETS = {
  inspection: [
    F('reviewDate', 'Review Date', 'date'),
    F('auditInspectionDate', 'Audit/Inspection Date', 'date'),
    F('inspectionArea', 'Inspection Area', 'text'),
    F('findingsSummary', 'Findings Summary', 'text'),
    F('actionStatus', 'Action Status', 'select', ['Completed', 'In Progress', 'Not Started'])
  ],
  training: [
    F('targetAudience', 'Target Audience', 'select', ['Internal', 'External']),
    F('trainingWorkshopName', 'Training/Workshop Name', 'text'),
    F('trainerFacilitator', 'Trainer/Facilitator', 'text'),
    F('sessionDate', 'Session Date', 'date'),
    F('completionStatus', 'Completion Status', 'select', ['Completed', 'Pending', 'Not Started'])
  ],
  facility: [
    F('assetFacilityName', 'Asset/Facility Name', 'text'),
    F('location', 'Location', 'text'),
    F('requestId', 'Request ID', 'text'),
    F('incidentRequestDate', 'Incident/Request Date', 'date'),
    F('resolutionStatus', 'Resolution Status', 'select', ['Resolved', 'Pending'])
  ],
  finance: [
    F('amount', 'Amount', 'number'),
    F('paymentStatus', 'Payment Status', 'select', ['Paid', 'Unpaid', 'Pending']),
    F('approvalNumber', 'Approval Number', 'text'),
    F('fiscalPeriodYear', 'Fiscal Period/Year', 'text'),
    F('vendor', 'Vendor', 'text')
  ],
  campaign: [
    F('versionNumber', 'Version Number', 'number'),
    F('campaignProjectName', 'Campaign/Project Name', 'text'),
    F('targetAudience', 'Target Audience', 'select', ['Internal', 'External']),
    F('eventLaunchDate', 'Event/Launch Date', 'date'),
    F('keyMessages', 'Key Messages', 'text')
  ],
  equipment: [
    F('equipmentName', 'Equipment Name', 'text'),
    F('equipmentLocation', 'Equipment Location', 'text'),
    F('documentValidityDate', 'Document Validity Date', 'date'),
    F('responsibleDepartment', 'Responsible Department', 'text'),
    F('reviewFrequency', 'Review Frequency', 'select', ['Quarterly', 'Annually'])
  ],
  software: [
    F('systemSoftwareName', 'System/Software Name', 'text'),
    F('versionRelease', 'Version/Release', 'text'),
    F('issueRequestId', 'Issue/Request ID', 'text'),
    F('priority', 'Priority', 'select', ['High', 'Medium', 'Low'])
  ],
  compliance: [
    F('approvalStatus', 'Approval Status', 'select', ['Approved', 'Draft', 'Pending']),
    F('startDate', 'Start Date', 'date'),
    F('caseProjectName', 'Case/Project Name', 'text'),
    F('validityPeriod', 'Validity Period', 'text'),
    F('expiryDate', 'Expiry Date', 'date'),
    F('complianceArea', 'Compliance Area', 'select', ['Health', 'Safety', 'Environment']),
    F('referenceNumber', 'Reference Number', 'text')
  ],
  contract: [
    F('summary', 'Summary', 'text'),
    F('entity', 'Entity', 'text'),
    F('reviewDate', 'Review Date', 'date'),
    F('referenceNumber', 'Reference Number', 'text')
  ],
  employee: [
    F('employeeName', 'Employee Name', 'text'),
    F('employeeId', 'Employee ID', 'text'),
    F('startDate', 'Start Date', 'date'),
    F('endDate', 'End Date', 'date'),
    F('positionDesignation', 'Position/Designation', 'text'),
    F('employeeStatus', 'Status', 'select', ['Active', 'Required'])
  ],
  correspondence: [
    F('effectiveDate', 'Effective Date', 'date'),
    F('approvalStatus', 'Approval Status', 'select', ['Approved', 'Draft', 'Pending']),
    F('versionNumber', 'Version Number', 'number'),
    F('referenceNumber', 'Reference Number', 'text')
  ]
};

// Upload categories in the live order, each mapped to its field set
const UPLOAD_CATEGORY_FIELDS = {
  'Accreditation Documents': 'inspection', 'Assessment Results': 'training', 'Asset Inventories': 'facility',
  'Audit Documents': 'finance', 'Branding Guidelines': 'campaign', 'Budget Plans': 'finance',
  'Calibration Certificates': 'equipment', 'Campaign Materials': 'campaign', 'Certification Records': 'training',
  'Change Requests': 'software', 'Clinical Trial Documents': 'equipment', 'Compliance Checklists': 'compliance',
  'Contracts': 'contract', 'Corrective Actions': 'inspection', 'Credit Notes': 'contract',
  'Curriculum Guides': 'training', 'Employee Contracts': 'employee', 'Equipment Logs': 'facility',
  'Equipment Maintenance Logs': 'equipment', 'Event Plans': 'campaign', 'Financial Reports': 'finance',
  'General Correspondences': 'correspondence', 'Incident Reports': 'software', 'Internal Audits': 'inspection',
  'Internal Memos': 'correspondence', 'Invoices and Receipts': 'finance', 'Licenses and Certifications': 'compliance',
  'Loans and Returns': 'contract', 'Maintenance Requests': 'facility', 'Meeting Minutes': 'correspondence',
  'Non-conformance Reports': 'inspection', 'Client Forms': 'equipment', 'Performance Reviews': 'employee',
  'Policies and Procedures': 'correspondence', 'Press Releases': 'campaign', 'Purchase Orders': 'finance',
  'Quality Assurance Reports': 'inspection', 'Recruitment Files': 'employee', 'Regulatory Filings': 'compliance',
  'Risk Assessments': 'compliance', 'Security Reports': 'facility', 'Social Media Content': 'campaign',
  'Software Documentation': 'software', 'Space Utilization Plans': 'facility',
  'Standard Operating Procedures (SOPs)': 'equipment', 'Supplier Registration Documents': 'contract',
  'Training Materials': 'correspondence', 'Training Records': 'employee', 'Training Schedules': 'training',
  'User Manuals': 'software', 'Warehouse Reports': 'contract'
};

function getUploadCategoryFields(category) {
  return UPLOAD_FIELD_SETS[UPLOAD_CATEGORY_FIELDS[category]] || [];
}

function populateUploadCategoryOptions() {
  const sel = document.getElementById('uploadCategorySelect');
  if (!sel || sel.dataset.populated) return;
  sel.innerHTML = '<option value="" disabled selected>Select Category</option>' +
    Object.keys(UPLOAD_CATEGORY_FIELDS).map(c => `<option value="${c}">${c}</option>`).join('');
  sel.dataset.populated = '1';
}

// Category-specific fields start blank every time the category changes
function handleUploadCategoryChange(category) {
  const container = document.getElementById('dynamicCategoryFieldsContainer');
  if (!container) return;
  const fields = getUploadCategoryFields(category);
  const now = new Date();
  const todayIso = now.toISOString().split('T')[0];
  container.innerHTML = fields.map(f => {
    let defVal = '';
    if (f.type === 'date') defVal = todayIso;
    else if (f.type === 'number') defVal = '1.0';
    return `
      <div class="sp-upload-form-row">
        <label class="sp-upload-form-label">${f.label} <span style="color: #dc2626;">*</span></label>
        <div class="sp-upload-form-control">
          ${f.type === 'select' ? `
          <select class="sp-form-input" id="dyn_${f.key}" data-field-key="${f.key}">
            <option value="" disabled>Select ${f.label}</option>
            ${f.options.map((o, idx) => `<option value="${o}" ${idx === 0 ? 'selected' : ''}>${o}</option>`).join('')}
          </select>` : `
          <input type="${f.type}" class="sp-form-input" id="dyn_${f.key}" data-field-key="${f.key}" value="${defVal}"${f.type === 'date' ? '' : ` placeholder="Enter ${f.label}"`}>`}
        </div>
      </div>`;
  }).join('');
}

function handleUploadStep1Next(event) {
  event.preventDefault();
  const category = document.getElementById('uploadCategorySelect').value;
  if (!category) {
    showToast('Please select a category.', 'error');
    return;
  }

  if (!AppState.attachedLocalFile) {
    const chipText = document.getElementById('uploadAttachedFileName')?.textContent?.trim();
    AppState.attachedLocalFile = {
      name: (chipText && chipText !== 'Choose a file...') ? chipText : 'Imran-Portal-Flow-Migration-Review (2).docx',
      size: '1.2 MB'
    };
  }

  // Ensure all required fields have a valid value
  document.querySelectorAll('#dynamicCategoryFieldsContainer .sp-upload-form-row').forEach(row => {
    const label = row.querySelector('label');
    const ctrl = row.querySelector('input, select, textarea');
    if (!label || !ctrl) return;
    if (!String(ctrl.value || '').trim()) {
      if (ctrl.tagName === 'SELECT' && ctrl.options.length > 1) {
        ctrl.selectedIndex = 1;
      } else if (ctrl.type === 'date') {
        ctrl.value = new Date().toISOString().split('T')[0];
      } else if (ctrl.type === 'number') {
        ctrl.value = '1.0';
      } else {
        ctrl.value = 'Standard ' + label.textContent.replace('*', '').trim();
      }
    }
  });

  const fileName = AppState.attachedLocalFile.name;
  const subCategory = document.getElementById('uploadSubCategory').value.trim() || '-';
  const description = document.getElementById('uploadDescription').value.trim() || '-';
  const stakeholder = document.getElementById('uploadStakeholder').value.trim() || '-';
  const owner = document.getElementById('uploadOwnerName')?.innerText || 'Intranet';
  const department = document.getElementById('uploadDeptBadge')?.innerText.trim() || 'Information Technology';
  const dateDisplay = document.getElementById('uploadDateDisplay').textContent.trim();

  const dynVersion = document.getElementById('dyn_version')?.value || '1.0';
  const dynProject = document.getElementById('dyn_projectName')?.value || document.getElementById('dyn_trainingName')?.value || document.getElementById('dyn_assetName')?.value || '';
  const dynAudience = document.getElementById('dyn_audience')?.value || document.getElementById('dyn_targetAudience')?.value || '';
  const dynLaunchDate = document.getElementById('dyn_launchDate')?.value || document.getElementById('dyn_sessionDate')?.value || document.getElementById('dyn_incidentDate')?.value || '';
  const dynKeyMessages = document.getElementById('dyn_keyMessages')?.value || '';
  const dynReviewDate = document.getElementById('dyn_reviewDate')?.value || '';
  const dynAuditInspectionDate = document.getElementById('dyn_auditInspectionDate')?.value || '';
  const dynInspectionArea = document.getElementById('dyn_inspectionArea')?.value || document.getElementById('dyn_location')?.value || '';
  const dynFindingsSummary = document.getElementById('dyn_findingsSummary')?.value || '';
  const dynActionStatus = document.getElementById('dyn_actionStatus')?.value || document.getElementById('dyn_completionStatus')?.value || document.getElementById('dyn_resolutionStatus')?.value || 'Active';
  const dynAmount = document.getElementById('dyn_amount')?.value || document.getElementById('dyn_budgetAmount')?.value || '-';
  const dynVendor = document.getElementById('dyn_vendor')?.value || '-';

  // Stash pending upload data
  AppState.pendingUploadData = {
    id: nextDocumentId(),
    name: fileName,
    type: fileName.endsWith('.xlsx') ? 'xlsx' : fileName.endsWith('.pdf') ? 'pdf' : fileName.endsWith('.pptx') ? 'pptx' : 'docx',
    category: category,
    subCategory: subCategory,
    department: department,
    owner: owner,
    stakeHolder: stakeholder,
    referenceNumber: 'INM-' + Math.floor(1000 + Math.random() * 9000),
    vendor: dynVendor,
    status: 'Active',
    uploadDate: dateDisplay,
    expireDate: '16 October 2027',
    description: description !== '-' ? description : (dynFindingsSummary || dynKeyMessages || dynProject || '-'),
    versionNumber: dynVersion,
    projectName: dynProject,
    targetAudience: dynAudience,
    launchDate: dynLaunchDate,
    keyMessages: dynKeyMessages,
    reviewDate: dynReviewDate,
    auditInspectionDate: dynAuditInspectionDate,
    inspectionArea: dynInspectionArea,
    findingsSummary: dynFindingsSummary,
    actionStatus: dynActionStatus,
    amount: dynAmount,
    size: '1.2 MB',
    location: AppState.currentUploadSource === 'my-files' ? 'My Files' : AppState.currentUploadSource === 'my-organization' ? 'My Organization' : 'My Department',
    favorite: false,
    isCheckedOut: false,
    checkedOutBy: '',
    sharedWith: []
  };

  // Store every category-specific field under its list column key
  document.querySelectorAll('#dynamicCategoryFieldsContainer [data-field-key]').forEach(el => {
    if (el.value) AppState.pendingUploadData[el.dataset.fieldKey] = el.value;
  });

  // If in Direct Upload mode, branch directly to Folder Selection (Step 2)
  if (AppState.isDirectUploadMode) {
    showDirectUploadStep2FolderSelect();
    return;
  }

  // Switch modal view to Step 2 (< Select Workflow)
  document.getElementById('uploadStep1Form').style.display = 'none';
  document.getElementById('uploadStep2Container').style.display = 'block';
  updateStep2WorkflowFooter();
  if (document.getElementById('uploadStep3IndividualContainer')) document.getElementById('uploadStep3IndividualContainer').style.display = 'none';
  if (document.getElementById('uploadStep3GroupContainer')) document.getElementById('uploadStep3GroupContainer').style.display = 'none';
  if (document.getElementById('uploadViewWorkflowContainer')) document.getElementById('uploadViewWorkflowContainer').style.display = 'none';
  if (document.getElementById('uploadStep4WorkflowFormContainer')) document.getElementById('uploadStep4WorkflowFormContainer').style.display = 'none';
  if (document.getElementById('uploadDirectFolderContainer')) document.getElementById('uploadDirectFolderContainer').style.display = 'none';
  if (document.getElementById('uploadDirectConfirmContainer')) document.getElementById('uploadDirectConfirmContainer').style.display = 'none';

  const backBtn = document.getElementById('uploadModalBackBtn');
  if (backBtn) backBtn.style.display = 'none';

  const mainTitle = document.getElementById('uploadModalMainTitle');
  if (mainTitle) mainTitle.innerText = '< Select Workflow';

  const dialog = document.querySelector('#uploadModal .dms-modal-dialog');
  if (dialog) dialog.style.maxWidth = '520px';
}

function updateStep2WorkflowFooter() {
  const isMyFiles = (AppState.currentUploadSource === 'my-files') || 
                    (AppState.currentRoute === 'my-files' && AppState.currentUploadSource !== 'my-department' && AppState.currentUploadSource !== 'my-organization');
  const skipBtn = document.getElementById('btnSkipApproval');
  if (skipBtn) {
    skipBtn.style.display = isMyFiles ? 'inline-block' : 'none';
  }
  // Department / Organization uploads offer "Self Approve" (as in the live DMS)
  const selfBtn = document.getElementById('btnSelfApprove');
  if (selfBtn) {
    selfBtn.style.display = isMyFiles ? 'none' : 'inline-block';
  }
}

// Unified Modal Back Click Router
function handleUploadModalBackClick() {
  const directConfirm = document.getElementById('uploadDirectConfirmContainer');
  const directFolder = document.getElementById('uploadDirectFolderContainer');
  const viewWf = document.getElementById('uploadViewWorkflowContainer');
  const step4 = document.getElementById('uploadStep4WorkflowFormContainer');
  const step3Ind = document.getElementById('uploadStep3IndividualContainer');
  const step3Grp = document.getElementById('uploadStep3GroupContainer');
  const step2 = document.getElementById('uploadStep2Container');

  if (directConfirm && directConfirm.style.display !== 'none') {
    directConfirm.style.display = 'none';
    if (directFolder) directFolder.style.display = 'block';
    const mainTitle = document.getElementById('uploadModalMainTitle');
    if (mainTitle) mainTitle.innerText = '< Select Folder';
    const dialog = document.querySelector('#uploadModal .dms-modal-dialog');
    if (dialog) dialog.style.maxWidth = '520px';
  } else if (directFolder && directFolder.style.display !== 'none') {
    directFolder.style.display = 'none';
    document.getElementById('uploadStep1Form').style.display = 'block';
    const mainTitle = document.getElementById('uploadModalMainTitle');
    if (mainTitle) mainTitle.innerText = 'Quick Document Upload';
    const dialog = document.querySelector('#uploadModal .dms-modal-dialog');
    if (dialog) dialog.style.maxWidth = '620px';
  } else if (viewWf && viewWf.style.display !== 'none') {
    backFromViewWorkflow();
  } else if (step4 && step4.style.display !== 'none') {
    backToWorkflowStep3();
  } else if ((step3Ind && step3Ind.style.display !== 'none') || (step3Grp && step3Grp.style.display !== 'none')) {
    backToUploadStep2();
  } else if (step2 && step2.style.display !== 'none') {
    backToUploadStep1();
  }
}

function backToUploadStep1() {
  document.getElementById('uploadStep1Form').style.display = 'block';
  document.getElementById('uploadStep2Container').style.display = 'none';
  if (document.getElementById('uploadStep3IndividualContainer')) document.getElementById('uploadStep3IndividualContainer').style.display = 'none';
  if (document.getElementById('uploadStep3GroupContainer')) document.getElementById('uploadStep3GroupContainer').style.display = 'none';
  if (document.getElementById('uploadViewWorkflowContainer')) document.getElementById('uploadViewWorkflowContainer').style.display = 'none';
  if (document.getElementById('uploadStep4WorkflowFormContainer')) document.getElementById('uploadStep4WorkflowFormContainer').style.display = 'none';
  if (document.getElementById('uploadDirectFolderContainer')) document.getElementById('uploadDirectFolderContainer').style.display = 'none';
  if (document.getElementById('uploadDirectConfirmContainer')) document.getElementById('uploadDirectConfirmContainer').style.display = 'none';

  const backBtn = document.getElementById('uploadModalBackBtn');
  if (backBtn) backBtn.style.display = 'none';

  const mainTitle = document.getElementById('uploadModalMainTitle');
  if (mainTitle) {
    mainTitle.innerText = AppState.isDirectUploadMode ? 'Quick Document Upload' : 'Upload Document';
  }

  const dialog = document.querySelector('#uploadModal .dms-modal-dialog');
  if (dialog) dialog.style.maxWidth = '620px';
}

// =============================================================================
// DIRECT UPLOAD MULTI-STEP FLOW CONTROLLERS
// =============================================================================
function showDirectUploadStep2FolderSelect() {
  document.getElementById('uploadStep1Form').style.display = 'none';
  if (document.getElementById('uploadStep2Container')) document.getElementById('uploadStep2Container').style.display = 'none';
  if (document.getElementById('uploadStep3IndividualContainer')) document.getElementById('uploadStep3IndividualContainer').style.display = 'none';
  if (document.getElementById('uploadStep3GroupContainer')) document.getElementById('uploadStep3GroupContainer').style.display = 'none';
  if (document.getElementById('uploadViewWorkflowContainer')) document.getElementById('uploadViewWorkflowContainer').style.display = 'none';
  if (document.getElementById('uploadStep4WorkflowFormContainer')) document.getElementById('uploadStep4WorkflowFormContainer').style.display = 'none';
  if (document.getElementById('uploadDirectConfirmContainer')) document.getElementById('uploadDirectConfirmContainer').style.display = 'none';

  const folderCont = document.getElementById('uploadDirectFolderContainer');
  if (folderCont) folderCont.style.display = 'block';

  const mainTitle = document.getElementById('uploadModalMainTitle');
  if (mainTitle) mainTitle.innerText = '< Select Folder';

  const dialog = document.querySelector('#uploadModal .dms-modal-dialog');
  if (dialog) dialog.style.maxWidth = '520px';

  updateDirectUploadFolderInputDisplay();
}

function updateDirectUploadFolderInputDisplay() {
  const displayBox = document.getElementById('directFolderDisplayContent');
  const nextBtn = document.getElementById('btnDirectFolderNext');
  if (!displayBox || !nextBtn) return;

  if (AppState.directUploadSelectedFolder) {
    displayBox.innerHTML = `
      <span style="font-size: 13.5px; color: #1e293b;">Selected Folder: <strong style="color: #0f172a;">${getDeptFolderPath(AppState.directUploadSelectedFolder).join(' > ')}</strong></span>
      <button type="button" onclick="clearDirectUploadSelectedFolder(event)" title="Remove selected folder" style="background: none; border: none; color: #ef4444; cursor: pointer; display: flex; align-items: center; padding: 2px;">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
          <line x1="18" y1="6" x2="6" y2="18"/>
          <line x1="6" y1="6" x2="18" y2="18"/>
        </svg>
      </button>
    `;
    nextBtn.disabled = false;
    nextBtn.style.background = '#008272';
    nextBtn.style.color = '#ffffff';
    nextBtn.style.cursor = 'pointer';
    nextBtn.style.opacity = '1';
  } else {
    displayBox.innerHTML = `
      <span id="directFolderPlaceholder" style="color: #64748b; font-size: 13.5px;">Select a folder</span>
      <button type="button" class="sp-folder-icon-btn" title="Select Folder" style="background: none; border: none; cursor: pointer; color: #008272; display: flex; align-items: center;">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>
        </svg>
      </button>
    `;
    nextBtn.disabled = true;
    nextBtn.style.background = '#e2e8f0';
    nextBtn.style.color = '#94a3b8';
    nextBtn.style.cursor = 'not-allowed';
    nextBtn.style.opacity = '0.5';
  }
}

function openDirectUploadFolderDrawer() {
  const drawer = document.getElementById('directUploadFolderDrawer');
  const backdrop = document.getElementById('directUploadDrawerBackdrop');
  if (!drawer) return;

  let dept = AppState.activeDepartment || 'Executive Director Office';

  const searchInput = document.getElementById('directFolderSearchInput');
  if (searchInput) searchInput.value = '';
  AppState.directUploadFilterQuery = '';
  AppState.directUploadTempFolder = AppState.directUploadSelectedFolder || null;

  renderDirectDrawerFoldersList();
  drawer.style.display = 'flex';
  if (backdrop) backdrop.style.display = 'block';
}

// Folder records of the active department (folders are rows flagged isFolder with a parentFolder)
function getActiveDeptFolders() {
  const branch = AppState.activeBranch || 'Innomick Delhi';
  const dept = AppState.activeDepartment || 'Information Technology';
  return ((InnomickData.departmentSites && InnomickData.departmentSites[branch] && InnomickData.departmentSites[branch][dept]) || [])
    .filter(d => d.isFolder);
}

// "Direct Uploads > CIO Office > Architecture Standards"
function getDeptFolderPath(folderName) {
  const folders = getActiveDeptFolders();
  const chain = [];
  let cur = folders.find(f => f.name === folderName);
  while (cur && chain.length < 10) {
    chain.unshift(cur.name);
    cur = cur.parentFolder ? folders.find(f => f.name === cur.parentFolder) : null;
  }
  return chain.length ? chain : (folderName ? [folderName] : []);
}

function toggleDirectDrawerFolderExpand(folderName, event) {
  if (event) event.stopPropagation();
  if (!AppState.directDrawerExpandedFolders) AppState.directDrawerExpandedFolders = {};
  AppState.directDrawerExpandedFolders[folderName] = !AppState.directDrawerExpandedFolders[folderName];
  renderDirectDrawerFoldersList();
}

function renderDirectDrawerFoldersList() {
  const listEl = document.getElementById('directDrawerFoldersList');
  if (!listEl) return;
  if (!AppState.directDrawerExpandedFolders) AppState.directDrawerExpandedFolders = {};

  const folders = getActiveDeptFolders();
  const q = (AppState.directUploadFilterQuery || '').toLowerCase();
  const childrenOf = parent => folders.filter(f => (f.parentFolder || null) === parent);
  const matches = f => !q || f.name.toLowerCase().includes(q) || childrenOf(f.name).some(matches);

  const renderNode = (folder, depth) => {
    if (!matches(folder)) return '';
    const kids = childrenOf(folder.name);
    const expanded = Boolean(q || AppState.directDrawerExpandedFolders[folder.name]);
    const selected = AppState.directUploadTempFolder === folder.name;
    const safe = folder.name.replace(/'/g, "\\'");
    return `
      <div class="direct-drawer-folder-item ${selected ? 'selected' : ''}" style="display: flex; align-items: center; justify-content: space-between; padding: 6px 10px 6px ${10 + depth * 20}px; border-radius: 4px; cursor: pointer; border: 1px solid ${selected ? '#008272' : 'transparent'}; background: ${selected ? '#e6f7f4' : 'transparent'}; margin-bottom: 2px;" onclick="selectDirectDrawerFolder('${safe}')">
        <div style="display: flex; align-items: center; gap: 8px;">
          ${kids.length ? `<span onclick="toggleDirectDrawerFolderExpand('${safe}', event)" title="${expanded ? 'Collapse' : 'Expand'}" style="width: 16px; display: inline-flex; justify-content: center; color: #64748b; font-size: 10px; transform: ${expanded ? 'rotate(90deg)' : 'none'}; transition: transform 0.15s;">&#9658;</span>` : '<span style="width: 16px;"></span>'}
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="${selected ? '#008272' : '#475569'}" stroke-width="1.8"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>
          <span style="font-size: 13px; font-weight: ${selected ? '600' : '400'}; color: ${selected ? '#008272' : '#1e293b'};">${folder.name}</span>
        </div>
      </div>
      ${kids.length && expanded ? kids.map(k => renderNode(k, depth + 1)).join('') : ''}`;
  };

  const html = childrenOf(null).map(f => renderNode(f, 0)).join('');
  listEl.innerHTML = html || `<div style="font-size: 13px; color: #94a3b8; padding: 12px 6px;">No folders found matching filter.</div>`;

  const breadcrumbEl = document.getElementById('directDrawerBreadcrumb');
  if (breadcrumbEl) {
    const dept = AppState.activeDepartment || 'Information Technology';
    breadcrumbEl.innerText = [dept, ...getDeptFolderPath(AppState.directUploadTempFolder || 'Direct Uploads')].join(' > ');
  }
}

// Selecting a folder also opens it so its sub-folders are visible
function selectDirectDrawerFolder(folderName) {
  AppState.directUploadTempFolder = folderName;
  if (!AppState.directDrawerExpandedFolders) AppState.directDrawerExpandedFolders = {};
  AppState.directDrawerExpandedFolders[folderName] = true;
  renderDirectDrawerFoldersList();
}

function saveDirectUploadSelectedFolder() {
  if (!AppState.directUploadTempFolder) {
    showToast('Please click on a folder to select it.', 'warning');
    return;
  }
  AppState.directUploadSelectedFolder = AppState.directUploadTempFolder;
  closeDirectUploadFolderDrawer();
  updateDirectUploadFolderInputDisplay();
}

function closeDirectUploadFolderDrawer() {
  const drawer = document.getElementById('directUploadFolderDrawer');
  const backdrop = document.getElementById('directUploadDrawerBackdrop');
  if (drawer) drawer.style.display = 'none';
  if (backdrop) backdrop.style.display = 'none';
}

function filterDirectUploadFolders(query) {
  AppState.directUploadFilterQuery = query || '';
  renderDirectDrawerFoldersList();
}

function clearDirectUploadSelectedFolder(event) {
  if (event) {
    event.preventDefault();
    event.stopPropagation();
  }
  AppState.directUploadSelectedFolder = null;
  AppState.directUploadTempFolder = null;
  updateDirectUploadFolderInputDisplay();
}

function proceedToDirectUploadConfirmation() {
  if (!AppState.directUploadSelectedFolder) {
    showToast('Please select a folder.', 'warning');
    return;
  }

  document.getElementById('uploadDirectFolderContainer').style.display = 'none';
  document.getElementById('uploadDirectConfirmContainer').style.display = 'block';

  const mainTitle = document.getElementById('uploadModalMainTitle');
  if (mainTitle) mainTitle.innerText = '< Upload Confirmation';

  const dialog = document.querySelector('#uploadModal .dms-modal-dialog');
  if (dialog) dialog.style.maxWidth = '480px';
}

function executeDirectUploadConfirm() {
  const folder = AppState.directUploadSelectedFolder || 'Testing';
  const baseData = AppState.pendingUploadData || {};
  const fileName = baseData.name || (AppState.attachedLocalFile ? AppState.attachedLocalFile.name : 'Abdul Test.xlsx');

  const now = new Date();
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const uploadDateFormatted = `${now.getDate()} ${months[now.getMonth()]} ${now.getFullYear()}`;
  const expireDateFormatted = `${now.getDate()} ${months[now.getMonth()]} ${now.getFullYear() + 1}`;

  const newDoc = {
    id: nextDocumentId(),
    name: fileName,
    type: fileName.split('.').pop().toLowerCase() || 'docx',
    category: baseData.category || 'Accreditation Documents',
    subCategory: baseData.subCategory !== '-' ? baseData.subCategory : '-',
    department: AppState.activeDepartment || 'Executive Director Office',
    owner: InnomickData.currentUser.name || 'MD Faisal Ahmed',
    status: 'Active',
    uploadDate: uploadDateFormatted,
    expireDate: expireDateFormatted,
    stakeHolder: baseData.stakeHolder !== '-' ? baseData.stakeHolder : 'Executive Office',
    referenceNumber: 'INM-' + Math.floor(1000 + Math.random() * 9000),
    vendor: baseData.vendor || 'Innomick Technologies Ltd.',
    description: baseData.description !== '-' ? baseData.description : `Direct upload to ${folder}`,
    folder: folder,
    parentFolder: folder,
    location: `${AppState.activeDepartment || 'Executive Director Office'} > Direct Uploads > ${folder}`,
    favorite: false,
    reviewDate: baseData.reviewDate || '',
    auditInspectionDate: baseData.auditInspectionDate || '',
    inspectionArea: baseData.inspectionArea || '',
    findingsSummary: baseData.findingsSummary || '',
    actionStatus: baseData.actionStatus || ''
  };

  // Close upload modal
  closeUploadModal();

  // Add document to department and show exact Info popup!
  addDocumentToDepartmentAndShowSuccess(newDoc, 'Successfully submitted !');
}

function addDocumentToDepartmentAndShowSuccess(doc, successMessage) {
  const branch = AppState.activeBranch || "Innomick Delhi";
  const dept = AppState.activeDepartment || "Information Technology";

  if (!InnomickData.departmentSites) InnomickData.departmentSites = {};
  if (!InnomickData.departmentSites[branch]) InnomickData.departmentSites[branch] = {};
  if (!InnomickData.departmentSites[branch][dept]) InnomickData.departmentSites[branch][dept] = [];

  const targetFolder = doc.parentFolder || (doc.folder && doc.folder !== "-" ? doc.folder : (AppState.directUploadSelectedFolder || null));

  const deptDoc = {
    ...doc,
    id: doc.id || nextDocumentId(),
    name: doc.name || "Document",
    type: doc.type || ((doc.name || "").split(".").pop() || "docx").toLowerCase(),
    category: doc.category || "Accreditation Documents",
    subCategory: doc.subCategory || "-",
    department: dept,
    owner: doc.owner || InnomickData.currentUser.name,
    status: doc.status || "Active",
    stakeHolder: doc.stakeHolder || "-",
    referenceNumber: doc.referenceNumber || "-",
    vendor: doc.vendor || "-",
    description: doc.description || "-",
    folder: targetFolder || "-",
    parentFolder: targetFolder && targetFolder !== "-" ? targetFolder : null,
    location: doc.location || `${dept} > Direct Uploads > ${targetFolder || ""}`,
    favorite: false
  };

  // One upload = one row in department library; isolated from My Files (Gap 31)
  const addOnce = (list) => { if (list && !list.some(d => d.id === deptDoc.id)) list.unshift(deptDoc); };
  addOnce(InnomickData.departmentSites[branch][dept]);
  addOnce(InnomickData.recentDocuments);
  addOnce(InnomickData.allFiles);

  if (deptDoc.parentFolder) {
    AppState.activeDeptFolder = deptDoc.parentFolder;
  }

  renderCurrentView();
  showInfoSuccessModal(successMessage || "Successfully submitted !");
}

function showInfoSuccessModal(msg) {
  const modal = document.getElementById('uploadSuccessModal');
  const msgEl = document.getElementById('uploadSuccessMessage');
  if (msgEl) msgEl.innerText = msg || 'Successfully submitted !';
  if (modal) modal.classList.add('active');
}

function closeUploadSuccessModal() {
  const modal = document.getElementById('uploadSuccessModal');
  if (modal) modal.classList.remove('active');
}


// Step 2 -> Step 3 (Individual Workflow)
function showSelectIndividualWorkflow() {
  AppState.activeWorkflowType = 'individual';
  AppState.selectedWorkflowTemplate = null;
  document.getElementById('uploadStep2Container').style.display = 'none';
  document.getElementById('uploadStep3IndividualContainer').style.display = 'block';
  if (document.getElementById('uploadStep3GroupContainer')) document.getElementById('uploadStep3GroupContainer').style.display = 'none';
  if (document.getElementById('uploadViewWorkflowContainer')) document.getElementById('uploadViewWorkflowContainer').style.display = 'none';
  if (document.getElementById('uploadStep4WorkflowFormContainer')) document.getElementById('uploadStep4WorkflowFormContainer').style.display = 'none';

  const backBtn = document.getElementById('uploadModalBackBtn');
  if (backBtn) backBtn.style.display = 'none';

  const mainTitle = document.getElementById('uploadModalMainTitle');
  if (mainTitle) mainTitle.innerText = '< Select Individual Workflow';

  const dialog = document.querySelector('#uploadModal .dms-modal-dialog');
  if (dialog) dialog.style.maxWidth = '740px';

  renderWorkflowTemplateCards('individual');
}

// Step 2 -> Step 3 (Group Workflow)
function showSelectGroupWorkflow() {
  AppState.activeWorkflowType = 'group';
  AppState.selectedWorkflowTemplate = null;
  document.getElementById('uploadStep2Container').style.display = 'none';
  document.getElementById('uploadStep3GroupContainer').style.display = 'block';
  if (document.getElementById('uploadStep3IndividualContainer')) document.getElementById('uploadStep3IndividualContainer').style.display = 'none';
  if (document.getElementById('uploadViewWorkflowContainer')) document.getElementById('uploadViewWorkflowContainer').style.display = 'none';
  if (document.getElementById('uploadStep4WorkflowFormContainer')) document.getElementById('uploadStep4WorkflowFormContainer').style.display = 'none';

  const backBtn = document.getElementById('uploadModalBackBtn');
  if (backBtn) backBtn.style.display = 'none';

  const mainTitle = document.getElementById('uploadModalMainTitle');
  if (mainTitle) mainTitle.innerText = '< Select Group Workflow';

  const dialog = document.querySelector('#uploadModal .dms-modal-dialog');
  if (dialog) dialog.style.maxWidth = '740px';

  renderWorkflowTemplateCards('group');
}

// Templates offered when submitting an upload: the upload presets plus every template
// saved under My Approvals > Workflow Templates (respecting "User Only" and "Department" access)
function getUploadWorkflowTemplates(type) {
  const isInd = type === 'individual';
  const presets = (isInd ? InnomickData.individualWorkflowTemplates : InnomickData.groupWorkflowTemplates) || [];
  const me = currentUserName();
  const saved = (InnomickData.workflowTemplates || [])
    .filter(t => t.type === (isInd ? 'Individual Workflow' : 'Group Workflow'))
    // Seed templates are already represented by the upload presets; only templates created in the prototype are added
    .filter(t => t.createdBy)
    .filter(t => {
      // Gap 35: When uploading at Department level, only Department templates apply; NEVER show Organization templates
      if (AppState.currentUploadSource === 'my-department') {
        if (t.accessLevel === 'Organization') return false;
        return t.accessLevel === 'Department' &&
          (!t.department || t.department === AppState.activeDepartment || t.department === InnomickData.currentUser.department) &&
          (!t.branch || t.branch === AppState.activeBranch);
      }
      if (t.accessLevel === 'Organization') return true;
      if (t.accessLevel === 'Department') {
        return t.department === AppState.activeDepartment &&
          (!t.branch || t.branch === AppState.activeBranch);
      }
      if (t.accessLevel === 'User Only') return !t.createdBy || t.createdBy === me;
      return false;
    })
    .map(t => isInd ? {
      id: t.id,
      name: t.name,
      type: 'individual',
      accessLevel: t.accessLevel,
      steps: (t.steps && t.steps.length ? t.steps : [{ stepName: t.stepName, approver: t.approver }]).map((st, i) => ({
        stepNumber: i + 1,
        stepTitle: st.stepName || `Step ${i + 1}`,
        stepName: st.stepName || '',
        approver: st.approver,
        action: st.action || 'Review & Approve',
        triggerMail: st.triggerMail !== false
      }))
    } : {
      id: t.id,
      name: t.name,
      type: 'group',
      accessLevel: t.accessLevel,
      groupName: t.group || t.name,
      groupMembers: [...(t.groupMembers || []), ...(t.additionalMembers || [])],
      action: t.action || 'Review & Approve'
    });
  const names = new Set(saved.map(t => t.name));
  let list = [...saved, ...presets.filter(p => !names.has(p.name))];
  // Gap 35: When in Department upload, strictly exclude Organization templates like IND ORG
  if (AppState.currentUploadSource === 'my-department') {
    list = list.filter(t => t.accessLevel !== 'Organization' && !(t.name && t.name.toUpperCase().includes('ORG')));
  }
  return list;
}

// Render Workflow Template Cards Grid (Matching Live SharePoint DMS Screenshot 1 & 3)
function renderWorkflowTemplateCards(type) {
  const isInd = type === 'individual';
  const grid = document.getElementById(isInd ? 'individualWorkflowGrid' : 'groupWorkflowGrid');
  const viewBtn = document.getElementById(isInd ? 'btnViewIndividualWorkflow' : 'btnViewGroupWorkflow');
  const submitBtn = document.getElementById(isInd ? 'btnSubmitIndividualWorkflow' : 'btnSubmitGroupWorkflow');

  if (viewBtn) viewBtn.style.display = 'none';
  if (submitBtn) {
    submitBtn.disabled = true;
    submitBtn.style.opacity = '0.6';
    submitBtn.style.cursor = 'not-allowed';
  }

  if (!grid) return;

  const templates = getUploadWorkflowTemplates(type);

  let html = '';
  templates.forEach(t => {
    const isSelected = AppState.selectedWorkflowTemplate && AppState.selectedWorkflowTemplate.id === t.id;
    html += `
      <div class="sp-workflow-card ${isSelected ? 'selected' : ''}" id="wfCard-${t.id}" onclick="selectWorkflowTemplateCard('${t.id}', '${type}')">
        <div class="sp-workflow-card-icon">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <circle cx="12" cy="5" r="2.5"/>
            <circle cx="6" cy="19" r="2.5"/>
            <circle cx="18" cy="19" r="2.5"/>
            <path d="M12 7.5v4.5M6 16.5v-1.5a3 3 0 0 1 3-3h6a3 3 0 0 1 3 3v1.5"/>
          </svg>
        </div>
        <div class="sp-workflow-card-title" title="${t.name}">${t.name}</div>
      </div>
    `;
  });

  grid.innerHTML = html;

  if (AppState.selectedWorkflowTemplate) {
    if (viewBtn) viewBtn.style.display = 'inline-block';
    if (submitBtn) {
      submitBtn.disabled = false;
      submitBtn.style.opacity = '1';
      submitBtn.style.cursor = 'pointer';
    }
  }
}

// Select a Workflow Template Card
function selectWorkflowTemplateCard(templateId, type) {
  const isInd = type === 'individual';
  const tpl = getUploadWorkflowTemplates(type).find(t => t.id === templateId);
  if (!tpl) return;

  AppState.selectedWorkflowTemplate = tpl;

  // Update card selected class
  const grid = document.getElementById(isInd ? 'individualWorkflowGrid' : 'groupWorkflowGrid');
  if (grid) {
    grid.querySelectorAll('.sp-workflow-card').forEach(c => c.classList.remove('selected'));
  }
  const cardEl = document.getElementById(`wfCard-${templateId}`);
  if (cardEl) cardEl.classList.add('selected');

  // Reveal View button & Enable Submit button (Exact Match with Live SharePoint Screenshot 1 & 3)
  const viewBtn = document.getElementById(isInd ? 'btnViewIndividualWorkflow' : 'btnViewGroupWorkflow');
  const submitBtn = document.getElementById(isInd ? 'btnSubmitIndividualWorkflow' : 'btnSubmitGroupWorkflow');

  if (viewBtn) viewBtn.style.display = 'inline-block';
  if (submitBtn) {
    submitBtn.disabled = false;
    submitBtn.style.opacity = '1';
    submitBtn.style.cursor = 'pointer';
  }
}

// Step 3 -> View Workflow (Matches Live SharePoint Screenshots: Individual vs Group Workflow)
function viewSelectedWorkflow() {
  const tpl = AppState.selectedWorkflowTemplate;
  if (!tpl) return;

  if (document.getElementById('uploadStep3IndividualContainer')) document.getElementById('uploadStep3IndividualContainer').style.display = 'none';
  if (document.getElementById('uploadStep3GroupContainer')) document.getElementById('uploadStep3GroupContainer').style.display = 'none';
  if (document.getElementById('uploadStep2Container')) document.getElementById('uploadStep2Container').style.display = 'none';
  if (document.getElementById('uploadStep4WorkflowFormContainer')) document.getElementById('uploadStep4WorkflowFormContainer').style.display = 'none';

  const viewContainer = document.getElementById('uploadViewWorkflowContainer');
  if (viewContainer) viewContainer.style.display = 'block';

  const mainTitle = document.getElementById('uploadModalMainTitle');
  if (mainTitle) mainTitle.innerText = '< View Workflow';

  const isGroup = (tpl.type === 'group') || (AppState.activeWorkflowType === 'group');

  const dialog = document.querySelector('#uploadModal .dms-modal-dialog');
  if (dialog) dialog.style.maxWidth = isGroup ? '580px' : '620px';

  if (isGroup) {
    renderViewGroupWorkflow(tpl);
  } else {
    renderViewIndividualWorkflow(tpl);
  }
}

// Render Group Workflow View (Exact Match with Live SharePoint Screenshots 2, 3, 4)
function renderViewGroupWorkflow(tpl) {
  const container = document.getElementById('viewWorkflowStepsList');
  if (!container) return;

  const groupName = tpl.groupName || tpl.name || 'Test-GM3';
  const members = tpl.groupMembers || ['Mohammed Wasifuddin'];
  const action = tpl.action || 'Review & Approve';
  const dueDate = tpl.dueDate || '22-10-2026';
  const addMembersChecked = !!tpl.addAdditionalMembers;

  let membersHtml = members.map(m => `
    <span class="sp-wf-member-pill">${m}</span>
  `).join('');

  container.innerHTML = `
    <div style="display: flex; flex-direction: column; gap: 14px; padding: 4px 0;">
      <!-- Group Name Row -->
      <div class="sp-wf-view-row">
        <span class="sp-wf-view-label">Group Name</span>
        <div class="sp-wf-view-field">${groupName}</div>
      </div>

      <!-- Group Members Row -->
      <div class="sp-wf-view-row" style="align-items: flex-start;">
        <span class="sp-wf-view-label" style="padding-top: 6px;">Group Members</span>
        <div class="sp-wf-members-box" id="wfGroupMembersBox">
          ${membersHtml}
        </div>
      </div>

      <!-- Add Additional members Toggle Row -->
      <div class="sp-wf-view-row">
        <span class="sp-wf-view-label">Add Additional members</span>
        <div style="flex: 1; display: flex; align-items: center;">
          <label class="sp-wf-toggle-switch">
            <input type="checkbox" id="wfAddMembersToggle" ${addMembersChecked ? 'checked' : ''} onchange="toggleGroupViewAddMembers(this.checked)">
            <span class="sp-wf-toggle-slider"></span>
          </label>
        </div>
      </div>

      <!-- Additional Members Selector Panel (Toggled by switch) -->
      <div id="wfAddMembersPickerRow" style="display: ${addMembersChecked ? 'block' : 'none'}; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 4px; padding: 12px; margin-left: 175px;">
        <div style="font-size: 12px; font-weight: 600; color: #475569; margin-bottom: 8px;">Select Additional Member:</div>
        <div style="display: flex; gap: 8px;">
          <select id="wfAddMemberSelect" class="sp-form-input" style="font-size: 12.5px; height: 32px; padding: 4px 8px;">
            <option value="Abdul Jabbar">Abdul Jabbar (Operations Head)</option>
            <option value="Sayed Asif Ahmed">Sayed Asif Ahmed (Systems Analyst)</option>
            <option value="Mohammed Wasifuddin">Mohammed Wasifuddin (VP Corporate Operations)</option>
            <option value="Aditya Rao">Aditya Rao (Systems Architect)</option>
            <option value="Neha Joshi">Neha Joshi (Operational Systems)</option>
            <option value="Mohammad Akram">Mohammad Akram (Quality Lead)</option>
            <option value="Mahesh Reddy">Mahesh Reddy (IT Manager)</option>
            <option value="Mohammed Zuberuddin">Mohammed Zuberuddin (Lead Systems)</option>
            <option value="Kavya Nair">Kavya Nair (Corporate Ops)</option>
          </select>
          <button type="button" onclick="handleAddMemberToGroupView()" style="background: #008272; color: #ffffff; border: none; border-radius: 4px; padding: 4px 14px; font-size: 12px; font-weight: 600; cursor: pointer;">Add</button>
        </div>
      </div>

      <!-- Action Row -->
      <div class="sp-wf-view-row">
        <span class="sp-wf-view-label">Action</span>
        <div class="sp-wf-view-field">${action}</div>
      </div>

      <!-- Due Date Row -->
      <div class="sp-wf-view-row">
        <span class="sp-wf-view-label">Due Date</span>
        <div class="sp-wf-view-field" style="justify-content: space-between;">
          <span>${dueDate}</span>
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2" style="flex-shrink:0;">
            <rect x="3" y="4" width="18" height="18" rx="2" ry="2"/>
            <line x1="16" y1="2" x2="16" y2="6"/>
            <line x1="8" y1="2" x2="8" y2="6"/>
            <line x1="3" y1="10" x2="21" y2="10"/>
          </svg>
        </div>
      </div>
    </div>
  `;
}

function toggleGroupViewAddMembers(isChecked) {
  if (AppState.selectedWorkflowTemplate) {
    AppState.selectedWorkflowTemplate.addAdditionalMembers = isChecked;
  }
  const panel = document.getElementById('wfAddMembersPickerRow');
  if (panel) {
    panel.style.display = isChecked ? 'block' : 'none';
  }
}

function handleAddMemberToGroupView() {
  const sel = document.getElementById('wfAddMemberSelect');
  if (!sel) return;
  const memberName = sel.value;
  const tpl = AppState.selectedWorkflowTemplate;
  if (!tpl) return;
  if (!tpl.groupMembers) tpl.groupMembers = [];
  if (!tpl.groupMembers.includes(memberName)) {
    tpl.groupMembers.push(memberName);
    const box = document.getElementById('wfGroupMembersBox');
    if (box) {
      const pill = document.createElement('span');
      pill.className = 'sp-wf-member-pill';
      pill.textContent = memberName;
      box.appendChild(pill);
    }
    showToast(`Added "${memberName}" to group members.`, 'info');
  } else {
    showToast(`"${memberName}" is already in group members.`, 'info');
  }
}

// Render Individual Workflow Steps in View Workflow Modal (Single vs Double Approval)
function renderViewIndividualWorkflow(tpl) {
  const container = document.getElementById('viewWorkflowStepsList');
  if (!container) return;

  const steps = tpl.steps || [{
    stepNumber: 1,
    stepTitle: 'Step 1: Asif',
    approver: 'Sayed Asif Ahmed',
    action: 'Review & Approve',
    dueDate: '22-10-2026',
    triggerMail: false
  }];
  const isSingle = steps.length === 1;

  let html = '';

  steps.forEach((step, idx) => {
    // Exact header title matching Screenshot 4 ("Approval 1", "Approval 2") & Screenshot 5 ("Step 1: Asif")
    const headerTitle = isSingle ? (step.stepTitle || 'Step 1: Asif') : `Approval ${step.stepNumber || (idx + 1)}`;

    html += `
      <div class="sp-wf-view-step-card">
        <div class="sp-wf-view-step-header">${headerTitle}</div>
        <div class="sp-wf-view-step-body">
          <!-- Approver -->
          <div class="sp-wf-view-row">
            <span class="sp-wf-view-label">Approver</span>
            <div class="sp-wf-view-field">${step.approver}</div>
          </div>
          <!-- Action -->
          <div class="sp-wf-view-row">
            <span class="sp-wf-view-label">Action</span>
            <div class="sp-wf-view-field">${step.action || 'Review & Approve'}</div>
          </div>
          <!-- Due Date -->
          <div class="sp-wf-view-row">
            <span class="sp-wf-view-label">Due Date</span>
            <div class="sp-wf-view-field" style="justify-content: space-between;">
              <span>${step.dueDate || '22-10-2026'}</span>
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2" style="flex-shrink:0;">
                <rect x="3" y="4" width="18" height="18" rx="2" ry="2"/>
                <line x1="16" y1="2" x2="16" y2="6"/>
                <line x1="8" y1="2" x2="8" y2="6"/>
                <line x1="3" y1="10" x2="21" y2="10"/>
              </svg>
            </div>
          </div>
          <!-- Trigger Mail to the Approver -->
          <div class="sp-wf-view-row">
            <span class="sp-wf-view-label">Trigger Mail to the Approver</span>
            <div style="flex: 1; display: flex; align-items: center;">
              <label class="sp-wf-toggle-switch">
                <input type="checkbox" ${step.triggerMail ? 'checked' : ''} onchange="toggleStepMail(${idx}, this.checked)">
                <span class="sp-wf-toggle-slider"></span>
              </label>
            </div>
          </div>
        </div>
      </div>
    `;

    // Down chevron connector between steps in double approval (Screenshot 4)
    if (idx < steps.length - 1) {
      html += `
        <div class="sp-wf-view-chevron-down">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#008272" stroke-width="2.5">
            <polyline points="6 9 12 15 18 9"></polyline>
          </svg>
        </div>
      `;
    }
  });

  container.innerHTML = html;
}

function toggleStepMail(stepIndex, isChecked) {
  if (AppState.selectedWorkflowTemplate && AppState.selectedWorkflowTemplate.steps && AppState.selectedWorkflowTemplate.steps[stepIndex]) {
    AppState.selectedWorkflowTemplate.steps[stepIndex].triggerMail = isChecked;
  }
}

// Back from View Workflow to Step 3 Selection
function backFromViewWorkflow() {
  const viewContainer = document.getElementById('uploadViewWorkflowContainer');
  if (viewContainer) viewContainer.style.display = 'none';

  const mainTitle = document.getElementById('uploadModalMainTitle');
  const dialog = document.querySelector('#uploadModal .dms-modal-dialog');

  if (AppState.activeWorkflowType === 'individual') {
    if (document.getElementById('uploadStep3IndividualContainer')) document.getElementById('uploadStep3IndividualContainer').style.display = 'block';
    if (mainTitle) mainTitle.innerText = '< Select Individual Workflow';
    if (dialog) dialog.style.maxWidth = '740px';
  } else {
    if (document.getElementById('uploadStep3GroupContainer')) document.getElementById('uploadStep3GroupContainer').style.display = 'block';
    if (mainTitle) mainTitle.innerText = '< Select Group Workflow';
    if (dialog) dialog.style.maxWidth = '640px';
  }
}

// Shared tail of every "submit for approval" path: list the document as In Review,
// raise one approval request, and reset filters so the new row is visible
function registerUploadForApproval(doc, approvalOpts) {
  doc.status = 'In Review';
  doc.location = AppState.currentUploadSource === 'my-organization' ? 'My Organization'
               : AppState.currentUploadSource === 'my-department' ? 'My Department'
               : 'My Files';

  // Gap 30 & 31: Documents pending approval do not appear in My Organization until approved.
  // Only add to My Files if uploaded specifically within My Files.
  if (AppState.currentUploadSource === 'my-files') {
    if (!InnomickData.myFiles) InnomickData.myFiles = [];
    InnomickData.myFiles.unshift(doc);
  }

  // Create approval request in Outbox and route to Approver(s)
  createApprovalRequest(doc, approvalOpts);
  resetAllColumnFilters();
  AppState.globalSearchQuery = '';

  closeUploadModal();
  updateApprovalsLeftBadges();

  // Gap 28: Automatically redirect to My Approvals -> Outbox after successful submission
  navigateTo('approvals');
  switchApprovalTab('outbox');
  showInfoSuccessModal('Successfully submitted !');
}

function resetAllColumnFilters() {
  Object.keys(AppState.columnFilters).forEach(k => { AppState.columnFilters[k] = ''; });
}

// Submit with Selected Workflow Template (Handles both Group and Individual flows)
function submitSelectedWorkflow() {
  const doc = AppState.pendingUploadData;
  const tpl = AppState.selectedWorkflowTemplate;
  if (!doc || !tpl) {
    showToast('Please select a workflow.', 'error');
    return;
  }

  if (tpl.type === 'group' || AppState.activeWorkflowType === 'group') {
    const groupName = tpl.groupName || tpl.name;
    registerUploadForApproval(doc, {
      isGroup: true,
      groupName,
      groupMembers: (tpl.groupMembers && tpl.groupMembers.length) ? tpl.groupMembers : (GROUP_MEMBERS_MAP[groupName] || [])
    });
  } else {
    const steps = (tpl.steps && tpl.steps.length ? tpl.steps : [{ approver: tpl.approver }])
      .filter(st => st.approver)
      .map(st => ({ stepName: st.stepName || (tpl.steps && tpl.steps.length > 1 ? '' : (st.stepTitle || '')), approver: st.approver }));
    if (!steps.length) {
      showToast('The selected workflow has no approver configured.', 'error');
      return;
    }
    registerUploadForApproval(doc, { isGroup: false, steps });
  }
}

// Back to Step 2 (< Select Workflow)
function backToUploadStep2() {
  if (document.getElementById('uploadStep3IndividualContainer')) document.getElementById('uploadStep3IndividualContainer').style.display = 'none';
  if (document.getElementById('uploadStep3GroupContainer')) document.getElementById('uploadStep3GroupContainer').style.display = 'none';
  if (document.getElementById('uploadViewWorkflowContainer')) document.getElementById('uploadViewWorkflowContainer').style.display = 'none';
  if (document.getElementById('uploadStep4WorkflowFormContainer')) document.getElementById('uploadStep4WorkflowFormContainer').style.display = 'none';
  document.getElementById('uploadStep2Container').style.display = 'block';
  updateStep2WorkflowFooter();

  const backBtn = document.getElementById('uploadModalBackBtn');
  if (backBtn) backBtn.style.display = 'none';

  const mainTitle = document.getElementById('uploadModalMainTitle');
  if (mainTitle) mainTitle.innerText = '< Select Workflow';

  const dialog = document.querySelector('#uploadModal .dms-modal-dialog');
  if (dialog) dialog.style.maxWidth = '520px';
}

// Step 3 -> Step 4: Open Workflow Form (+ Blank or Template)
function openWorkflowConfigForm(type, templateName) {
  AppState.activeWorkflowType = type;
  if (document.getElementById('uploadStep3IndividualContainer')) document.getElementById('uploadStep3IndividualContainer').style.display = 'none';
  if (document.getElementById('uploadStep3GroupContainer')) document.getElementById('uploadStep3GroupContainer').style.display = 'none';
  if (document.getElementById('uploadViewWorkflowContainer')) document.getElementById('uploadViewWorkflowContainer').style.display = 'none';
  document.getElementById('uploadStep4WorkflowFormContainer').style.display = 'block';

  const backBtn = document.getElementById('uploadModalBackBtn');
  if (backBtn) backBtn.style.display = 'none';

  const mainTitle = document.getElementById('uploadModalMainTitle');
  const labelEl = document.getElementById('wfLabelEntity');
  const selectEl = document.getElementById('wfSelectEntity');

  if (type === 'group') {
    if (mainTitle) mainTitle.innerText = '< Group Workflow';
    if (labelEl) labelEl.innerText = 'Group';
    // Exact options from live SharePoint screenshot group_workflow_form_dropdown_reference.png
    if (selectEl) {
      selectEl.innerHTML = `
        <option value="Test-GM3" ${templateName === 'Test-GM3' ? 'selected' : ''}>Test-GM3</option>
        <option value="Test-GM5">Test-GM5</option>
        <option value="CRM">CRM</option>
        <option value="HSE Committee" ${templateName === 'HSE Committee' ? 'selected' : ''}>HSE Committee</option>
        <option value="Hyderabad Staff">Hyderabad Staff</option>
        <option value="Finance_CanPrintReceipts">Finance_CanPrintReceipts</option>
        <option value="SG Senior Engineers" ${templateName === 'SG Senior Engineers' ? 'selected' : ''}>SG Senior Engineers</option>
        <option value="Compliance Meeting">Compliance Meeting</option>
      `;
    }
  } else {
    if (mainTitle) mainTitle.innerText = '< Individual Workflow';
    if (labelEl) labelEl.innerText = 'Approver';
    if (selectEl) {
      selectEl.innerHTML = `
        <option value="Mohammed Wasifuddin" selected>Mohammed Wasifuddin</option>
        <option value="Sarah Thomas">Sarah Thomas</option>
        <option value="Imran Hussain">Imran Hussain</option>
        <option value="Farhan Qureshi">Farhan Qureshi</option>
        <option value="Nandini Rao">Nandini Rao</option>
      `;
    }
  }

  const dialog = document.querySelector('#uploadModal .dms-modal-dialog');
  if (dialog) dialog.style.maxWidth = '520px';
}

// Back from Step 4 to Step 3
function backToWorkflowStep3() {
  document.getElementById('uploadStep4WorkflowFormContainer').style.display = 'none';
  if (document.getElementById('uploadViewWorkflowContainer')) document.getElementById('uploadViewWorkflowContainer').style.display = 'none';
  const backBtn = document.getElementById('uploadModalBackBtn');
  if (backBtn) backBtn.style.display = 'none';

  const mainTitle = document.getElementById('uploadModalMainTitle');
  const dialog = document.querySelector('#uploadModal .dms-modal-dialog');

  if (AppState.activeWorkflowType === 'individual') {
    document.getElementById('uploadStep3IndividualContainer').style.display = 'block';
    if (mainTitle) mainTitle.innerText = '< Select Individual Workflow';
  } else {
    document.getElementById('uploadStep3GroupContainer').style.display = 'block';
    if (mainTitle) mainTitle.innerText = '< Select Group Workflow';
  }

  if (dialog) dialog.style.maxWidth = '740px';
}

// Skip Workflow execution (Immediately publishes document directly into file list in My Files)
function executeSkipWorkflowUpload() {
  const doc = AppState.pendingUploadData;
  if (!doc) return;

  doc.status = 'Active';

  if (AppState.currentUploadSource === 'my-department') {
    closeUploadModal();
    addDocumentToDepartmentAndShowSuccess(doc, 'Successfully submitted !');
    return;
  }

  if (AppState.currentUploadSource === 'my-organization') {
    doc.location = 'My Organization';
    if (InnomickData.documents) InnomickData.documents.unshift(doc);
  } else {
    doc.location = 'My Files';
    if (!InnomickData.myFiles) InnomickData.myFiles = [];
    InnomickData.myFiles.unshift(doc);
  }

  if (InnomickData.recentDocuments) {
    InnomickData.recentDocuments.unshift({
      id: doc.id,
      name: doc.name,
      type: doc.type,
      uploadDate: doc.uploadDate || '2026-10-23',
      lastAccessed: 'Just now',
      location: doc.location || 'My Files',
      owner: doc.owner || InnomickData.currentUser.name || 'Mohammed Wasifuddin',
      category: doc.category || 'Accreditation Documents',
      subCategory: doc.subCategory || '-',
      department: doc.department || 'Information Technology',
      referenceNumber: doc.referenceNumber || 'AST-2026-902',
      stakeHolder: doc.stakeHolder || 'IT Asset Directorate',
      vendor: doc.vendor || 'Innomick Technologies Ltd.',
      description: doc.description || 'Uploaded document directly.',
      favorite: false
    });
  }

  // Clear any active column search filters so the user's new row is 100% visible!
  AppState.columnFilters = {
    name: '',
    category: '',
    status: '',
    department: '',
    referenceNumber: '',
    favorite: '',
    stakeHolder: '',
    owner: '',
    subCategory: '',
    uploadDate: '',
    id: '',
    vendor: '',
    expireDate: '',
    description: ''
  };
  AppState.globalSearchQuery = '';

  closeUploadModal();
  renderCurrentView();
  showInfoSuccessModal('Successfully submitted !');
}

function toggleUploadFilePermissions() {
  const row = document.getElementById('uploadFilePermissionRow');
  if (row) row.style.display = row.style.display === 'none' ? 'flex' : 'none';
}

// "Self Approve": the uploader approves their own document, so it is published straight away
function executeSelfApproveUpload() {
  const doc = AppState.pendingUploadData;
  if (!doc) return;
  doc.status = 'Active';
  closeUploadModal();
  if (AppState.currentUploadSource === 'my-department') {
    addDocumentToDepartmentAndShowSuccess(doc, 'Successfully submitted !');
    return;
  }
  if (AppState.currentUploadSource === 'my-organization') {
    doc.location = 'My Organization';
    if (InnomickData.documents) InnomickData.documents.unshift(doc);
  } else {
    doc.location = 'My Files';
    if (!InnomickData.myFiles) InnomickData.myFiles = [];
    InnomickData.myFiles.unshift(doc);
  }
  if (InnomickData.recentDocuments) InnomickData.recentDocuments.unshift({ ...doc, lastAccessed: 'Just now' });
  renderCurrentView();
  showInfoSuccessModal('Successfully submitted !');
}

// Submit with Workflow execution (Blank / custom workflow configured inside the upload dialog)
function executeWorkflowUploadSubmit() {
  const doc = AppState.pendingUploadData;
  if (!doc) return;

  const isGroup = AppState.activeWorkflowType === 'group';
  const entitySelect = document.getElementById('wfSelectEntity');
  const selectedEntity = entitySelect ? entitySelect.value : '';
  if (!selectedEntity) {
    showToast(isGroup ? 'Please select a group.' : 'Please select an approver.', 'error');
    return;
  }
  const actionSelect = document.getElementById('wfSelectAction');
  const actionName = actionSelect ? actionSelect.value : 'Review & Approve';
  const saveAsTemplateChecked = document.getElementById('wfSaveAsTemplate')?.checked;
  const groupMembers = isGroup ? (GROUP_MEMBERS_MAP[selectedEntity] || []) : [];

  if (saveAsTemplateChecked) {
    if (!InnomickData.workflowTemplates) InnomickData.workflowTemplates = [];
    InnomickData.workflowTemplates.unshift(isGroup ? {
      id: 'wt-' + Date.now(),
      name: `${selectedEntity} ${actionName}`,
      type: 'Group Workflow',
      accessLevel: 'User Only',
      createdBy: currentUserName(),
      group: selectedEntity,
      groupMembers,
      action: actionName
    } : {
      id: 'wt-' + Date.now(),
      name: `${selectedEntity} ${actionName}`,
      type: 'Individual Workflow',
      accessLevel: 'User Only',
      createdBy: currentUserName(),
      steps: [{ stepName: '', approver: selectedEntity, action: actionName, triggerMail: true }],
      stepName: '',
      approver: selectedEntity,
      action: actionName,
      triggerMail: true
    });
    persistWorkflowTemplates();
  }

  registerUploadForApproval(doc, isGroup
    ? { isGroup: true, groupName: selectedEntity, groupMembers }
    : { isGroup: false, steps: [{ stepName: '', approver: selectedEntity }] });
}

// =============================================================================
// MICROSOFT 365 / SHAREPOINT HEADER PANELS & PROFILE FLYOUT CONTROLLER
// (Matches Exact Right Side Panels from Live SharePoint Screenshots)
// =============================================================================

let activeHeaderPanelId = null;
let selectedFeedbackCategory = 'compliment';

function toggleHeaderPanel(panelKey) {
  const panelMap = {
    'copilot': { panel: 'sidePanelCopilot', btn: 'topbarCopilotBtn' },
    'nextsteps': { panel: 'sidePanelNextSteps', btn: 'topbarNextStepsBtn' },
    'feedback': { panel: 'sidePanelFeedback', btn: 'topbarFeedbackBtn' },
    'settings': { panel: 'sidePanelSettings', btn: 'topbarSettingsBtn' },
    'help': { panel: 'sidePanelHelp', btn: 'topbarHelpBtn' }
  };

  const target = panelMap[panelKey];
  if (!target) return;

  closeM365ProfileFlyout();

  const el = document.getElementById(target.panel);
  const btn = document.getElementById(target.btn);
  if (!el) return;

  const isCurrentlyOpen = (activeHeaderPanelId === panelKey && el.style.display !== 'none');

  // Close any open panel first
  closeHeaderPanel();

  if (!isCurrentlyOpen) {
    el.style.display = 'flex';
    if (btn) btn.classList.add('active');
    activeHeaderPanelId = panelKey;

    if (panelKey === 'copilot') {
      const greeting = document.getElementById('copilotGreetingText');
      if (greeting) {
        greeting.innerText = `Hello ${InnomickData.currentUser.name || 'Intranet'}!`;
      }
      setTimeout(() => {
        const inp = document.getElementById('copilotTextInput');
        if (inp) inp.focus();
      }, 100);
    }
  }
}

function closeHeaderPanel() {
  const panels = ['sidePanelCopilot', 'sidePanelNextSteps', 'sidePanelFeedback', 'sidePanelSettings', 'sidePanelHelp'];
  panels.forEach(id => {
    const p = document.getElementById(id);
    if (p) p.style.display = 'none';
  });

  const btns = ['topbarCopilotBtn', 'topbarNextStepsBtn', 'topbarFeedbackBtn', 'topbarSettingsBtn', 'topbarHelpBtn'];
  btns.forEach(id => {
    const b = document.getElementById(id);
    if (b) b.classList.remove('active');
  });

  activeHeaderPanelId = null;
}

// -----------------------------------------------------------------------------
// 1. Copilot Interactions
// -----------------------------------------------------------------------------
function handleCopilotPrompt(promptText) {
  const messagesList = document.getElementById('copilotMessagesList');
  if (!messagesList) return;

  // Append user message
  const userMsg = document.createElement('div');
  userMsg.className = 'copilot-msg-user';
  userMsg.innerText = promptText;
  messagesList.appendChild(userMsg);

  // Generate contextual AI response
  const botMsg = document.createElement('div');
  botMsg.className = 'copilot-msg-bot';
  botMsg.innerHTML = '<span style="color:#64748b; font-style:italic;">Copilot is reviewing site documents...</span>';
  messagesList.appendChild(botMsg);

  const scrollContainer = document.getElementById('copilotChatScroll');
  if (scrollContainer) scrollContainer.scrollTop = scrollContainer.scrollHeight;

  setTimeout(() => {
    let reply = '';
    if (promptText.includes('How can you help me')) {
      reply = `<strong>Here is what I can do on this Document Management System:</strong><br><br>
      &bull; <strong>Search & Summarize</strong>: Find policies, technical inventories, audit logs, or SOPs across all categories.<br>
      &bull; <strong>Approval Tracking</strong>: Check pending approval requests in your Outbox and active workflows.<br>
      &bull; <strong>Metadata Insights</strong>: Auto-detect missing fields (Target Audience, Location, Vendor, Due Dates) before file submission.<br>
      &bull; <strong>Permissions & Sharing</strong>: Check who has access to sensitive intranet documents.`;
    } else if (promptText.includes('recent updates')) {
      const pendingCount = (InnomickData.pendingApprovals || []).length;
      reply = `<strong>Recent updates on Innomick Technologies DMS:</strong><br><br>
      &bull; <strong>New Uploads</strong>: <em>Testing Jawad.docx</em>, <em>test.docx</em>, and <em>Developer-Fix-Retest-Results.xlsx</em> were recently indexed under Accreditation Documents and Asset Inventories.<br>
      &bull; <strong>Workflows</strong>: You have <strong>${pendingCount} pending approvals</strong> and active workflows in progress.<br>
      &bull; <strong>Library Cache</strong>: All 32 document categories are synchronized with SharePoint intranet lists.`;
    } else if (promptText.includes('relevant to me')) {
      const cur = InnomickData.currentUser;
      reply = `<strong>Relevant information for ${cur.name} (${cur.department}):</strong><br><br>
      &bull; <strong>Role</strong>: ${cur.role}<br>
      &bull; <strong>Assigned Library</strong>: <em>${cur.department}</em> and <em>My Organization</em> documents.<br>
      &bull; <strong>Pending Actions</strong>: Access your <a onclick="navigateTo('approvals'); closeHeaderPanel();" style="color:#0078d4; cursor:pointer; text-decoration:underline;">Approvals Dashboard</a> to review items awaiting your sign-off or track your submitted Outbox.`;
    } else {
      reply = `I reviewed the DMS library for <em>"${promptText}"</em>. You can locate corresponding files in <strong>My Files</strong> or under <strong>Categories</strong>. Let me know if you would like me to draft an approval summary or check document version history.`;
    }

    botMsg.innerHTML = reply;
    if (scrollContainer) scrollContainer.scrollTop = scrollContainer.scrollHeight;
  }, 450);
}

function sendCopilotMessage() {
  const input = document.getElementById('copilotTextInput');
  if (!input) return;
  const text = input.value.trim();
  if (!text) return;
  input.value = '';
  handleCopilotPrompt(text);
}

// -----------------------------------------------------------------------------
// 2. Next Steps Actions
// -----------------------------------------------------------------------------
function handleNextStepAction(action) {
  closeHeaderPanel();
  switch (action) {
    case 'upload':
      triggerDeviceUpload();
      break;
    case 'distribute':
      showToast('Opening Engage & Viva news distribution sync settings...', 'info');
      break;
    case 'customize':
      showToast('SharePoint site layout customization activated.', 'info');
      break;
    case 'post-news':
      showToast('Creating new intranet news & announcement post...', 'info');
      break;
    case 'change-look':
      showToast('Opening theme, branding & color palette configuration...', 'info');
      break;
    default:
      showToast(`Selected next step: ${action}`);
  }
}

// -----------------------------------------------------------------------------
// 3. Submit Feedback Interactions
// -----------------------------------------------------------------------------
function selectFeedbackType(type) {
  selectedFeedbackCategory = type;

  const btnC = document.getElementById('fbBtnCompliment');
  const btnP = document.getElementById('fbBtnProblem');
  const btnS = document.getElementById('fbBtnSuggestion');

  [btnC, btnP, btnS].forEach(b => {
    if (b) {
      b.classList.remove('selected');
      b.classList.add('outline');
    }
  });

  const detailArea = document.getElementById('feedbackDetailArea');
  const label = document.getElementById('feedbackPromptLabel');
  const textarea = document.getElementById('feedbackCommentInput');

  if (detailArea) detailArea.style.display = 'block';

  if (type === 'compliment' && btnC) {
    btnC.classList.add('selected');
    btnC.classList.remove('outline');
    if (label) label.innerText = 'Tell us what you liked:';
    if (textarea) textarea.placeholder = 'What features or parts of the document system worked best?';
  } else if (type === 'problem' && btnP) {
    btnP.classList.add('selected');
    btnP.classList.remove('outline');
    if (label) label.innerText = 'Describe the problem:';
    if (textarea) textarea.placeholder = 'What issue or unexpected behavior did you encounter?';
  } else if (type === 'suggestion' && btnS) {
    btnS.classList.add('selected');
    btnS.classList.remove('outline');
    if (label) label.innerText = 'Make a suggestion:';
    if (textarea) textarea.placeholder = 'What enhancement or new workflow would you like to see?';
  }

  updateFeedbackSubmitState();
}

function updateFeedbackSubmitState() {
  const textarea = document.getElementById('feedbackCommentInput');
  const submitBtn = document.getElementById('btnSubmitFeedback');
  if (submitBtn) {
    const hasText = textarea && textarea.value.trim().length > 0;
    submitBtn.disabled = !hasText;
  }
}

function submitUserFeedback() {
  const textarea = document.getElementById('feedbackCommentInput');
  const val = textarea ? textarea.value.trim() : '';
  showToast('Thank you! Your feedback has been submitted to Microsoft.', 'success');
  if (textarea) textarea.value = '';
  closeHeaderPanel();
}

// -----------------------------------------------------------------------------
// 4. Settings Actions
// -----------------------------------------------------------------------------
function handleSettingsAction(action) {
  closeHeaderPanel();
  switch (action) {
    case 'add-page':
      showToast('Opening SharePoint page creation wizard...', 'info');
      break;
    case 'add-app':
      showToast('Opening SharePoint app & web part catalog...', 'info');
      break;
    case 'site-contents':
      navigateTo('categories');
      break;
    case 'site-info':
      showToast('Site Information: Innomick Technologies DMS (Enterprise Environment)');
      break;
    case 'site-ai':
      toggleHeaderPanel('copilot');
      break;
    case 'site-permissions':
      openManageAccessModal('doc-1');
      break;
    case 'apply-template':
      showToast('Opening Enterprise DMS site templates library...', 'info');
      break;
    case 'site-usage':
      showToast('Site Analytics: 1,420 document views this week, 98% workflow completion SLA.', 'info');
      break;
    case 'site-reviews':
    case 'site-performance':
      showToast('Site Performance: Page load < 350ms, cache status optimal.', 'info');
      break;
    case 'change-look':
    case 'site-branding':
      showToast('Opening Site Branding & Logo configuration...', 'info');
      break;
    case 'manage-dashboard':
    case 'news-distribution':
    case 'view-all':
      showToast(`Opening ${action} settings in Microsoft 365 Admin Center...`, 'info');
      break;
    default:
      showToast(`Opening setting: ${action}`);
  }
}

// -----------------------------------------------------------------------------
// 5. Microsoft 365 Profile Flyout & Persona Switching
// -----------------------------------------------------------------------------
function toggleM365ProfileFlyout(event) {
  if (event) event.stopPropagation();
  closeHeaderPanel();

  const flyout = document.getElementById('m365ProfileFlyout');
  if (!flyout) return;

  const isVisible = flyout.style.display === 'block';
  if (isVisible) {
    flyout.style.display = 'none';
  } else {
    // Populate current persona details
    const cur = InnomickData.currentUser;
    const nameEl = document.getElementById('profileUserName');
    const emailEl = document.getElementById('profileUserEmail');
    const avatarLetter = document.getElementById('profileAvatarLetter');
    const orgName = document.getElementById('profileOrgName');

    if (nameEl) nameEl.innerText = cur.name;
    if (emailEl) emailEl.innerText = cur.email;
    if (avatarLetter) avatarLetter.innerText = cur.avatar || cur.name.charAt(0);
    if (orgName) orgName.innerText = 'Innomick Technologies';

    // Item 22: Immediately render grouped personas in Switch User (Demo) panel
    renderProfilePersonaList();

    flyout.style.display = 'block';
  }
}

function closeM365ProfileFlyout() {
  const flyout = document.getElementById('m365ProfileFlyout');
  if (flyout) flyout.style.display = 'none';
}

function togglePersonaPickerFromProfile() {
  const picker = document.getElementById('profilePersonaPicker');
  if (!picker) return;

  const isVisible = picker.style.display === 'block';
  if (isVisible) {
    picker.style.display = 'none';
  } else {
    renderProfilePersonaList();
    picker.style.display = 'block';
  }
}

function renderProfilePersonaList() {
  const container = document.getElementById('profilePersonaGroupedList') || document.getElementById('profilePersonaList');
  if (!container) return;

  const personas = InnomickData.availablePersonas || [];
  const currentId = (InnomickData.currentUser && InnomickData.currentUser.id) || 'wasif';

  // Group personas by roleGroup (Item 22: Switch User Demo panel matching ProjectHub)
  const groups = {};
  personas.forEach(p => {
    const grp = p.roleGroup || 'Other Roles';
    if (!groups[grp]) groups[grp] = [];
    groups[grp].push(p);
  });

  let html = '<div style="font-size: 11px; color: #64748b; margin-bottom: 8px;">Select a user to view from their perspective</div>';
  Object.keys(groups).forEach(groupName => {
    const count = groups[groupName].length;
    html += `
      <div style="font-size: 10px; font-weight: 700; color: #008272; text-transform: uppercase; margin: 10px 0 4px 2px; letter-spacing: 0.5px; display: flex; justify-content: space-between;">
        <span>${groupName.toUpperCase()} (${count})</span>
      </div>
    `;
    groups[groupName].forEach(p => {
      const isSelected = p.id === currentId;
      html += `
        <div onclick="switchUserPersona('${p.id}')" style="display: flex; align-items: center; justify-content: space-between; padding: 7px 10px; border-radius: 4px; cursor: pointer; transition: background 0.15s; background: ${isSelected ? '#f0fdfa' : '#ffffff'}; margin-bottom: 4px; border: 1px solid ${isSelected ? '#99f6e4' : '#e2e8f0'};" onmouseover="this.style.background='#f1f5f9'" onmouseout="this.style.background='${isSelected ? '#f0fdfa' : '#ffffff'}'">
          <div style="display: flex; align-items: center; gap: 8px;">
            <div style="width: 28px; height: 28px; border-radius: 50%; background: ${isSelected ? '#008272' : '#64748b'}; color: white; display: flex; align-items: center; justify-content: center; font-weight: 600; font-size: 11px;">
              ${p.avatar}
            </div>
            <div>
              <strong style="display: block; font-size: 12px; color: #1e293b;">${p.name}</strong>
              <span style="font-size: 10.5px; color: #64748b;">${p.department} &bull; <em style="color: #008272;">${p.role}</em></span>
            </div>
          </div>
          ${isSelected ? '<span style="color: #008272; font-weight: bold; font-size: 13px;">✓</span>' : ''}
        </div>
      `;
    });
  });

  container.innerHTML = html;
}

const PERSONA_STORAGE_KEY = 'innomick_dms_active_persona';

function applyPersonaToHeader(persona) {
  const avatarText = document.getElementById('topbarAvatarText');
  if (avatarText) avatarText.innerText = persona.avatar;
  const profileAvatarLetter = document.getElementById('profileAvatarLetter');
  if (profileAvatarLetter) profileAvatarLetter.innerText = persona.avatar;
  const profileUserName = document.getElementById('profileUserName');
  if (profileUserName) profileUserName.innerText = persona.name;
  const profileUserEmail = document.getElementById('profileUserEmail');
  if (profileUserEmail) profileUserEmail.innerText = persona.email;
}

function restorePersistedPersona() {
  let id = null;
  try { id = localStorage.getItem(PERSONA_STORAGE_KEY); } catch (e) {}
  const persona = id && (InnomickData.availablePersonas || []).find(p => p.id === id);
  if (persona) InnomickData.currentUser = { ...persona };
  applyPersonaToHeader(InnomickData.currentUser);
}

function switchUserPersona(personaId) {
  const persona = (InnomickData.availablePersonas || []).find(p => p.id === personaId);
  if (!persona) return;

  InnomickData.currentUser = { ...persona };
  try { localStorage.setItem(PERSONA_STORAGE_KEY, persona.id); } catch (e) {}
  AppState.approvalsInboxPage = 1;
  AppState.approvalsOutboxPage = 1;
  AppState.approvalsHistoryPage = 1;

  // Update topbar avatar
  const avatarText = document.getElementById('topbarAvatarText');
  if (avatarText) avatarText.innerText = persona.avatar;

  const profileAvatarLetter = document.getElementById('profileAvatarLetter');
  if (profileAvatarLetter) profileAvatarLetter.innerText = persona.avatar;

  const profileUserName = document.getElementById('profileUserName');
  if (profileUserName) profileUserName.innerText = persona.name;

  const profileUserEmail = document.getElementById('profileUserEmail');
  if (profileUserEmail) profileUserEmail.innerText = persona.email;

  // Close flyout
  closeM365ProfileFlyout();

  showToast(`Switched active profile to "${persona.name}" (${persona.role})`, 'info');
  // Gap 29: After successful user switch, automatically redirect to the Home page of the selected user
  navigateTo('home');
}

function handleProfileSignOut(event) {
  if (event) {
    event.preventDefault();
    event.stopPropagation();
  }
  closeM365ProfileFlyout();
  showToast('Signed out of Microsoft 365 Intranet session.', 'info');
}

// Global click handler to close panels/flyouts when clicking outside
window.addEventListener('click', (e) => {
  // Close M365 profile flyout if clicking outside
  const flyout = document.getElementById('m365ProfileFlyout');
  if (flyout && flyout.style.display === 'block') {
    if (!e.target.closest('#m365ProfileFlyout') && !e.target.closest('#userPersonaBtn')) {
      flyout.style.display = 'none';
    }
  }

  // Close side panels if clicking outside both the panel and the topbar trigger buttons
  if (activeHeaderPanelId) {
    const isInsidePanel = e.target.closest('.m365-side-panel');
    const isInsideTopbarBtn = e.target.closest('.topbar-icon-btn');
    const isFloatingCopilot = e.target.closest('#floatingCopilotBadge');
    if (!isInsidePanel && !isInsideTopbarBtn && !isFloatingCopilot) {
      closeHeaderPanel();
    }
  }
});

// Escape key closes open header panels and flyout
window.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    closeHeaderPanel();
    closeM365ProfileFlyout();
  }
});

// =============================================================================
// WORKFLOW TEMPLATE VIEW & EDIT MODALS (INDIVIDUAL & GROUP WORKFLOWS)
// =============================================================================
// =============================================================================
// WORKFLOW TEMPLATE VIEW & EDIT MODALS (INDIVIDUAL & GROUP WORKFLOWS)
// =============================================================================
let activeWorkflowTemplateId = null;
let activeApprovalDocId = null;

// Safe select value setter (creates option if missing so value is never lost)
function setSelectValueSafe(selectId, value) {
  const el = document.getElementById(selectId);
  if (!el) return;
  if (!value) {
    el.value = '';
    return;
  }
  const exists = Array.from(el.options).some(opt => opt.value === value);
  if (!exists) {
    const opt = document.createElement('option');
    opt.value = value;
    opt.textContent = value;
    el.appendChild(opt);
  }
  el.value = value;
}

function setFormFieldLocked(id, isLocked) {
  const el = document.getElementById(id);
  if (!el) return;
  el.disabled = isLocked;
  if (isLocked) {
    el.classList.add('is-readonly');
    el.setAttribute('readonly', 'true');
  } else {
    el.classList.remove('is-readonly');
    el.removeAttribute('readonly');
  }
}

function viewWorkflowTemplate(id) {
  activeWorkflowTemplateId = id;
  const template = (InnomickData.workflowTemplates || []).find(t => t.id === id);
  if (!template) return;

  if (template.type === 'Group Workflow') {
    openGroupWorkflowModal(template, true);
  } else {
    openIndividualWorkflowModal(template, true);
  }
}

// Workflow Multi-Step Wizard State
let currentWorkflowSteps = [];
let currentWorkflowStepIndex = 0;

// Departments per branch (the sidebar lists the signed-in user's own department under each)
const BRANCH_DEPARTMENTS = {
  'Innomick Hyderabad': ['Executive Director Office', 'CEO Office', 'Finance & Accounts', 'Human Resources & Talent', 'Customer Relations', 'Facilities'],
  'Innomick Delhi': ['Information Technology', 'Engineering & Cloud Architecture', 'QA & Automation Testing', 'Central Supply Chain Management', 'Client Solutions & Delivery']
};

function populateWorkflowDeptOptions(selectId, branch, selected) {
  const sel = document.getElementById(selectId);
  if (!sel) return;
  // Department stays disabled until a branch is chosen (as in the live DMS)
  const depts = BRANCH_DEPARTMENTS[branch] || [];
  sel.innerHTML = `<option value="">Please Select Department</option>` +
    depts.map(d => `<option value="${d}">${d}</option>`).join('');
  sel.disabled = !branch;
  sel.value = depts.includes(selected) ? selected : '';
}

function handleIndWfAccessLevelChange(val) {
  const deptGroup = document.getElementById('indWfDeptFieldsGroup');
  if (deptGroup) {
    deptGroup.style.display = val === 'Department' ? 'block' : 'none';
  }
}

function handleGroupWfAccessLevelChange(val) {
  const deptGroup = document.getElementById('groupWfDeptFieldsGroup');
  if (deptGroup) {
    deptGroup.style.display = val === 'Department' ? 'block' : 'none';
  }
}

function backFromIndividualWorkflow() {
  closeIndividualWorkflowModal();
  openCreateWorkflowModal();
}

function backFromGroupWorkflow() {
  closeGroupWorkflowModal();
  openCreateWorkflowModal();
}

function renderWorkflowStepView(index) {
  if (!currentWorkflowSteps || currentWorkflowSteps.length === 0) {
    currentWorkflowSteps = [{ stepName: '', approver: '', action: 'Review & Approve', triggerMail: true }];
  }
  if (index < 0) index = 0;
  if (index >= currentWorkflowSteps.length) index = currentWorkflowSteps.length - 1;
  currentWorkflowStepIndex = index;

  const step = currentWorkflowSteps[index];

  // Update step title and navigation buttons
  const stepTitle = document.getElementById('indWfStepTitle');
  if (stepTitle) {
    stepTitle.textContent = `Step ${index + 1}`;
  }

  const prevBtn = document.getElementById('indWfStepPrevBtn');
  if (prevBtn) {
    prevBtn.style.display = index > 0 ? 'inline-flex' : 'none';
  }

  const nextBtn = document.getElementById('indWfStepNextBtn');
  if (nextBtn) {
    nextBtn.style.display = index < currentWorkflowSteps.length - 1 ? 'inline-flex' : 'none';
  }

  // Populate step fields - allow custom step name without forced 'Step 1' prefill (Item 1)
  const stepNameInput = document.getElementById('indWfStepName');
  if (stepNameInput) {
    stepNameInput.value = (step.stepName !== undefined && step.stepName !== null) ? step.stepName : '';
    stepNameInput.placeholder = 'Enter Step Name';
  }

  setSelectValueSafe('indWfApprover', step.approver || '');

  const triggerMailInput = document.getElementById('indWfTriggerMail');
  if (triggerMailInput) {
    triggerMailInput.checked = step.triggerMail !== false;
  }
}

function saveCurrentStepData() {
  if (!currentWorkflowSteps || !currentWorkflowSteps[currentWorkflowStepIndex]) return;
  const stepNameInput = document.getElementById('indWfStepName');
  const approverInput = document.getElementById('indWfApprover');
  const triggerMailInput = document.getElementById('indWfTriggerMail');

  currentWorkflowSteps[currentWorkflowStepIndex] = {
    stepName: stepNameInput ? stepNameInput.value.trim() : '',
    approver: approverInput ? approverInput.value : '',
    action: 'Review & Approve',
    triggerMail: triggerMailInput ? triggerMailInput.checked : true
  };
}

function navigateWorkflowStep(delta) {
  saveCurrentStepData();
  const nextIdx = currentWorkflowStepIndex + delta;
  if (nextIdx >= 0 && nextIdx < currentWorkflowSteps.length) {
    renderWorkflowStepView(nextIdx);
  }
}

function addWorkflowStep() {
  saveCurrentStepData();
  const newStepNum = currentWorkflowSteps.length + 1;
  currentWorkflowSteps.push({
    stepName: '',
    approver: '',
    action: 'Review & Approve',
    triggerMail: true
  });
  renderWorkflowStepView(currentWorkflowSteps.length - 1);
  showToast(`Step ${newStepNum} added to workflow`);
}


// Open modal to create brand new Individual Workflow (editable)
function openCreateIndividualWorkflow() {
  closeCreateWorkflowModal();
  activeWorkflowTemplateId = null;

  const backBtn = document.getElementById('indWfBackBtn');
  if (backBtn) backBtn.style.display = 'inline-flex';

  const modalTitle = document.getElementById('indWfModalTitle');
  if (modalTitle) modalTitle.textContent = 'Individual Workflow';

  const addStepBtn = document.getElementById('indWfAddStepBtn');
  if (addStepBtn) addStepBtn.style.display = 'inline-flex';

  document.getElementById('indWfTemplateName').value = '';
  document.getElementById('indWfTemplateName').placeholder = 'Enter Template Name';
  setSelectValueSafe('indWfAccessLevel', '');
  handleIndWfAccessLevelChange('');

  const defaultBranch = AppState.activeBranch === 'Innomick Delhi' ? 'Innomick Delhi' : 'Innomick Hyderabad';
  setSelectValueSafe('indWfBranch', '');
  populateWorkflowDeptOptions('indWfDept', '', '');

  setFormFieldLocked('indWfTemplateName', false);
  setFormFieldLocked('indWfAccessLevel', false);
  setFormFieldLocked('indWfBranch', false);
  setFormFieldLocked('indWfDept', false);
  setFormFieldLocked('indWfStepName', false);
  setFormFieldLocked('indWfApprover', false);

  // Item 1: Step Name allows custom free-text input without prefilling 'Step 1'
  currentWorkflowSteps = [
    { stepName: '', approver: '', action: 'Review & Approve', triggerMail: true }
  ];
  currentWorkflowStepIndex = 0;
  renderWorkflowStepView(0);

  const modal = document.getElementById('individualWorkflowModal');
  if (modal) modal.classList.add('active');
}

// Open modal to create brand new Group Workflow (editable)
function openCreateGroupWorkflow() {
  closeCreateWorkflowModal();
  activeWorkflowTemplateId = null;

  const backBtn = document.getElementById('groupWfBackBtn');
  if (backBtn) backBtn.style.display = 'inline-flex';

  const modalTitle = document.getElementById('groupWfModalTitle');
  if (modalTitle) modalTitle.textContent = 'Group Workflow';

  document.getElementById('groupWfTemplateName').value = '';
  document.getElementById('groupWfTemplateName').placeholder = 'Enter Template Name';
  setSelectValueSafe('groupWfAccessLevel', '');
  handleGroupWfAccessLevelChange('');

  const defaultBranch = AppState.activeBranch === 'Innomick Delhi' ? 'Innomick Delhi' : 'Innomick Hyderabad';
  setSelectValueSafe('groupWfBranch', '');
  populateWorkflowDeptOptions('groupWfDept', '', '');

  setFormFieldLocked('groupWfTemplateName', false);
  setFormFieldLocked('groupWfAccessLevel', false);
  setFormFieldLocked('groupWfBranch', false);
  setFormFieldLocked('groupWfDept', false);
  setFormFieldLocked('groupWfGroup', false);

  // Item 5: Start with blank group selection and empty members chips
  setSelectValueSafe('groupWfGroup', '');
  renderGroupMembersChips([]);

  const addToggle = document.getElementById('groupWfAddAdditionalToggle');
  if (addToggle) {
    addToggle.checked = false;
    toggleAdditionalMembersField(false);
  }

  const actionSelect = document.getElementById('groupWfAction');
  if (actionSelect) {
    actionSelect.value = 'Review & Approve';
  }

  const modal = document.getElementById('groupWorkflowModal');
  if (modal) modal.classList.add('active');
}

function openIndividualWorkflowModal(template = {}, isExisting = true) {
  const modal = document.getElementById('individualWorkflowModal');
  if (!modal) return;

  const backBtn = document.getElementById('indWfBackBtn');
  if (backBtn) backBtn.style.display = isExisting ? 'none' : 'inline-flex';

  const modalTitle = document.getElementById('indWfModalTitle');
  if (modalTitle) modalTitle.textContent = 'Individual Workflow';

  const addStepBtn = document.getElementById('indWfAddStepBtn');
  if (addStepBtn) addStepBtn.style.display = isExisting ? 'none' : 'inline-flex';

  const defaultBranch = AppState.activeBranch === 'Innomick Delhi' ? 'Innomick Delhi' : 'Innomick Hyderabad';

  document.getElementById('indWfTemplateName').value = template.name || 'Testing UAT';
  setSelectValueSafe('indWfAccessLevel', template.accessLevel || 'Department');
  handleIndWfAccessLevelChange(template.accessLevel || 'Department');

  setSelectValueSafe('indWfBranch', template.branch || defaultBranch);
  populateWorkflowDeptOptions('indWfDept', document.getElementById('indWfBranch').value, template.department || 'Facilities');

  // Lock metadata fields for existing workflow templates (as seen in Live SharePoint Screenshot)
  setFormFieldLocked('indWfTemplateName', isExisting);
  setFormFieldLocked('indWfAccessLevel', isExisting);
  setFormFieldLocked('indWfBranch', isExisting);
  setFormFieldLocked('indWfDept', isExisting);
  setFormFieldLocked('indWfAction', isExisting);

  // Editable fields in Individual Workflow: Step Name, Approver, and Trigger Mail
  setFormFieldLocked('indWfStepName', false);
  setFormFieldLocked('indWfApprover', false);

  if (template.steps && Array.isArray(template.steps) && template.steps.length > 0) {
    currentWorkflowSteps = template.steps.map(s => ({ ...s }));
  } else {
    currentWorkflowSteps = [
      {
        stepName: template.stepName || '',
        approver: template.approver || '',
        action: template.action || 'Review & Approve',
        triggerMail: template.triggerMail !== false
      }
    ];
  }
  currentWorkflowStepIndex = 0;
  renderWorkflowStepView(0);

  modal.classList.add('active');
}

function closeIndividualWorkflowModal() {
  const modal = document.getElementById('individualWorkflowModal');
  if (modal) modal.classList.remove('active');
  activeWorkflowTemplateId = null;
}

function submitIndividualWorkflow() {
  saveCurrentStepData();
  const name = document.getElementById('indWfTemplateName').value.trim();
  const accessLevel = document.getElementById('indWfAccessLevel').value;
  if (!activeWorkflowTemplateId) {
    if (!name) { showToast('Please enter a template name.', 'error'); return; }
    if (!accessLevel) { showToast('Please select an access level.', 'error'); return; }
  }
  const missingIdx = currentWorkflowSteps.findIndex(st => !st.approver);
  if (missingIdx !== -1) {
    renderWorkflowStepView(missingIdx);
    showToast(`Please select an approver for Step ${missingIdx + 1}.`, 'error');
    return;
  }
  const defaultBranch = AppState.activeBranch === 'Innomick Delhi' ? 'Innomick Delhi' : 'Innomick Hyderabad';
  const branch = document.getElementById('indWfBranch').value || defaultBranch;
  const department = document.getElementById('indWfDept').value || 'Facilities';

  if (activeWorkflowTemplateId) {
    const templateId = activeWorkflowTemplateId;
    const template = (InnomickData.workflowTemplates || []).find(t => t.id === templateId);
    if (template) {
      template.name = name;
      template.accessLevel = accessLevel;
      template.branch = branch;
      template.department = department;
      template.steps = [...currentWorkflowSteps];
      if (currentWorkflowSteps[0]) {
        template.stepName = currentWorkflowSteps[0].stepName || '';
        template.approver = currentWorkflowSteps[0].approver;
        template.triggerMail = currentWorkflowSteps[0].triggerMail;
      }
      persistWorkflowTemplates();
    }
    closeIndividualWorkflowModal();
    openTemplateUpdatedModal(templateId);
    return;
  }

  // Add newly created template
  InnomickData.workflowTemplates.unshift({
    id: 'wt-' + Date.now(),
    name: name,
    type: 'Individual Workflow',
    accessLevel: accessLevel,
    branch: branch,
    department: department,
    steps: [...currentWorkflowSteps],
    stepName: currentWorkflowSteps[0] ? currentWorkflowSteps[0].stepName : '',
    approver: currentWorkflowSteps[0] ? currentWorkflowSteps[0].approver : 'MD Faisal Ahmed',
    action: 'Review & Approve',
    triggerMail: currentWorkflowSteps[0] ? currentWorkflowSteps[0].triggerMail : true,
    createdBy: currentUserName()
  });

  persistWorkflowTemplates();
  closeIndividualWorkflowModal();
  showToast(`Individual Workflow "${name}" saved successfully!`, 'success');
  renderCurrentView();
}

// Members currently shown as chips (the group's members minus any the user removed)
function getGroupMembersFromChips() {
  return [...document.querySelectorAll('#groupWfMembersContainer .sp-group-member-pill > span:first-child')].map(el => el.textContent.trim());
}

function renderGroupMembersChips(members = []) {
  const container = document.getElementById('groupWfMembersContainer');
  if (!container) return;
  const hasGroup = Boolean((document.getElementById('groupWfGroup') || {}).value);
  ['groupWfMembersRow', 'groupWfAddAdditionalRow'].forEach(id => {
    const row = document.getElementById(id);
    if (row) row.style.display = hasGroup ? 'flex' : 'none';
  });
  if (!hasGroup) {
    const addToggle = document.getElementById('groupWfAddAdditionalToggle');
    if (addToggle) addToggle.checked = false;
    toggleAdditionalMembersField(false);
  }
  if (!members || members.length === 0) {
    container.innerHTML = '<span style="color: #605e5c; font-size: 13px;">No members in group</span>';
    return;
  }
  container.innerHTML = members.map(m => `
    <span class="sp-group-member-pill">
      <span>${m}</span>
      <span class="sp-group-member-minus-btn" title="Remove member" onclick="removeGroupMemberChip('${m}')">&minus;</span>
    </span>
  `).join('');
}

function removeGroupMemberChip(memberName) {
  const template = (InnomickData.workflowTemplates || []).find(t => t.id === activeWorkflowTemplateId);
  if (template && template.groupMembers) {
    template.groupMembers = template.groupMembers.filter(m => m !== memberName);
    renderGroupMembersChips(template.groupMembers);
  } else {
    const container = document.getElementById('groupWfMembersContainer');
    if (container) {
      const pills = container.querySelectorAll('.sp-group-member-pill');
      pills.forEach(p => {
        if (p.textContent.includes(memberName)) p.remove();
      });
    }
  }
}

function handleGroupSelectChange(groupName) {
  if (!groupName) {
    renderGroupMembersChips([]);
    const template = (InnomickData.workflowTemplates || []).find(t => t.id === activeWorkflowTemplateId);
    if (template) {
      template.group = '';
      template.groupMembers = [];
    }
    return;
  }
  const members = GROUP_MEMBERS_MAP[groupName] || ['Mohammed Wasifuddin'];
  const template = (InnomickData.workflowTemplates || []).find(t => t.id === activeWorkflowTemplateId);
  if (template) {
    template.group = groupName;
    template.groupMembers = [...members];
  }
  renderGroupMembersChips(members);
}

function openGroupWorkflowModal(template = {}, isExisting = true) {
  const modal = document.getElementById('groupWorkflowModal');
  if (!modal) return;

  const backBtn = document.getElementById('groupWfBackBtn');
  if (backBtn) backBtn.style.display = isExisting ? 'none' : 'inline-flex';

  const modalTitle = document.getElementById('groupWfModalTitle');
  if (modalTitle) modalTitle.textContent = 'Group Workflow';

  const defaultBranch = AppState.activeBranch === 'Innomick Delhi' ? 'Innomick Delhi' : 'Innomick Hyderabad';

  document.getElementById('groupWfTemplateName').value = template.name || 'Testing GRP Latest';
  setSelectValueSafe('groupWfAccessLevel', template.accessLevel || 'Department');
  handleGroupWfAccessLevelChange(template.accessLevel || 'Department');

  setSelectValueSafe('groupWfBranch', template.branch || defaultBranch);
  populateWorkflowDeptOptions('groupWfDept', document.getElementById('groupWfBranch').value, template.department || 'CEO Office');
  setSelectValueSafe('groupWfGroup', template.group || '');

  // Lock metadata fields for existing workflow templates (as seen in Live SharePoint Screenshots 4 & 5)
  setFormFieldLocked('groupWfTemplateName', isExisting);
  setFormFieldLocked('groupWfAccessLevel', isExisting);
  setFormFieldLocked('groupWfBranch', isExisting);
  setFormFieldLocked('groupWfDept', isExisting);
  setFormFieldLocked('groupWfAction', isExisting);

  // Group remains editable
  setFormFieldLocked('groupWfGroup', false);

  const groupName = template.group || '';
  const members = template.groupMembers || (groupName ? (GROUP_MEMBERS_MAP[groupName] || ['Mohammed Wasifuddin']) : []);
  renderGroupMembersChips(members);

  const addToggle = document.getElementById('groupWfAddAdditionalToggle');
  if (addToggle) {
    addToggle.checked = template.addAdditional === true;
    toggleAdditionalMembersField(addToggle.checked);
  }

  renderAdditionalMembersChips(template.additionalMembers || []);
  document.getElementById('groupWfAction').value = template.action || 'Review & Approve';

  modal.classList.add('active');
}

function closeGroupWorkflowModal() {
  const modal = document.getElementById('groupWorkflowModal');
  if (modal) modal.classList.remove('active');
  activeWorkflowTemplateId = null;
}

function submitGroupWorkflow() {
  const name = document.getElementById('groupWfTemplateName').value.trim();
  const accessLevel = document.getElementById('groupWfAccessLevel').value;
  if (!activeWorkflowTemplateId) {
    if (!name) { showToast('Please enter a template name.', 'error'); return; }
    if (!accessLevel) { showToast('Please select an access level.', 'error'); return; }
  }
  if (!document.getElementById('groupWfGroup').value) { showToast('Please select a group.', 'error'); return; }
  const defaultBranch = AppState.activeBranch === 'Innomick Delhi' ? 'Innomick Delhi' : 'Innomick Hyderabad';
  const branch = document.getElementById('groupWfBranch').value || defaultBranch;
  const department = document.getElementById('groupWfDept').value || 'CEO Office';
  const group = document.getElementById('groupWfGroup').value || '';
  const addAdditional = document.getElementById('groupWfAddAdditionalToggle').checked;
  const action = document.getElementById('groupWfAction').value || 'Review & Approve';

  if (activeWorkflowTemplateId) {
    const templateId = activeWorkflowTemplateId;
    const template = (InnomickData.workflowTemplates || []).find(t => t.id === templateId);
    if (template) {
      template.name = name;
      template.accessLevel = accessLevel;
      template.branch = branch;
      template.department = department;
      template.group = group;
      template.groupMembers = getGroupMembersFromChips();
      template.addAdditional = addAdditional;
      template.additionalMembers = addAdditional ? [...currentGroupAdditionalMembers] : [];
      template.action = action;
      persistWorkflowTemplates();
    }
    closeGroupWorkflowModal();
    openTemplateUpdatedModal(templateId);
    return;
  }

  // Add newly created group template
  InnomickData.workflowTemplates.unshift({
    id: 'wt-' + Date.now(),
    name: name,
    type: 'Group Workflow',
    accessLevel: accessLevel,
    branch: branch,
    department: department,
    group: group,
    groupMembers: getGroupMembersFromChips(),
    addAdditional: addAdditional,
    additionalMembers: addAdditional ? [...currentGroupAdditionalMembers] : [],
    accessLevelUsers: accessLevel === 'User Only' ? [currentUserName()] : [],
    createdBy: currentUserName(),
    action: action
  });

  persistWorkflowTemplates();
  closeGroupWorkflowModal();
  showToast(`Group Workflow "${name}" saved successfully!`, 'success');
  renderCurrentView();
}

// =============================================================================
// TEMPLATE UPDATED CONFIRMATION MODAL (Matching Live SharePoint Screenshot)
// =============================================================================
let pendingUpdatedTemplateId = null;

function openTemplateUpdatedModal(templateId) {
  pendingUpdatedTemplateId = templateId;
  const modal = document.getElementById('templateUpdatedModal');
  if (modal) {
    modal.classList.add('active');
  }
}

function closeTemplateUpdatedModal() {
  const modal = document.getElementById('templateUpdatedModal');
  if (modal) {
    modal.classList.remove('active');
  }
  pendingUpdatedTemplateId = null;
}

function handleTemplateUpdatedDecision(updateOngoing) {
  const template = (InnomickData.workflowTemplates || []).find(t => t.id === pendingUpdatedTemplateId);
  const templateName = template ? template.name : 'Template';

  if (updateOngoing) {
    if (InnomickData.approvalsInbox && getMyInbox().length > 0) {
      InnomickData.approvalsInbox.forEach(item => {
        if (!item.workflowTemplate || item.workflowTemplate === templateName || item.workflowTemplate === pendingUpdatedTemplateId) {
          if (template && template.approver) {
            item.assignedTo = template.approver;
          }
          if (template && template.stepName) {
            item.stage = template.stepName;
          }
        }
      });
    }
    showToast(`Ongoing approvals updated with new "${templateName}" workflow!`, 'success');
  } else {
    showToast(`"${templateName}" updated for future approvals.`, 'info');
  }

  closeTemplateUpdatedModal();
  renderCurrentView();
}

const ADDITIONAL_MEMBERS_OPTIONS = [
  "Ananya Reddy",
  "Mohammed Wasifuddin",
  "Sayed Asif Ahmed",
  "MD Faisal Ahmed",
  "Mohammed Zuberuddin",
  "Imran Hussain",
  "Abdul Jabbar",
  "Arjun Mehta",
  "Neha Joshi",
  "Fatima Shaikh",
  "Nelluri Venkatarao",
  "Navin Patel",
  "Priya Sharma",
  "Rahul Verma",
  "Kavya Nair",
  "Yash Dhanraj",
  "Jawad Shaikh"
];

let currentGroupAdditionalMembers = [];

function toggleAdditionalMembersField(show) {
  const row = document.getElementById('additionalMembersRow');
  if (row) {
    row.style.display = show ? 'flex' : 'none';
  }
  if (!show) {
    closeAdditionalMembersDropdown();
  }
}

function toggleAdditionalMembersDropdown(e) {
  if (e) {
    e.stopPropagation();
  }
  const dropdown = document.getElementById('additionalMembersDropdown');
  const box = document.querySelector('.sp-additional-members-box');
  if (!dropdown) return;

  const isShown = dropdown.style.display === 'block';
  if (isShown) {
    dropdown.style.display = 'none';
    if (box) box.classList.remove('active');
  } else {
    renderAdditionalMembersDropdown();
    dropdown.style.display = 'block';
    if (box) box.classList.add('active');
  }
}

function closeAdditionalMembersDropdown() {
  const dropdown = document.getElementById('additionalMembersDropdown');
  const box = document.querySelector('.sp-additional-members-box');
  if (dropdown) dropdown.style.display = 'none';
  if (box) box.classList.remove('active');
}

function renderAdditionalMembersDropdown() {
  const dropdown = document.getElementById('additionalMembersDropdown');
  if (!dropdown) return;

  dropdown.innerHTML = ADDITIONAL_MEMBERS_OPTIONS.map(opt => {
    const isSelected = currentGroupAdditionalMembers.includes(opt);
    return `
      <div class="sp-additional-member-option ${isSelected ? 'selected' : ''}" onclick="selectAdditionalMember('${opt}', event)">
        <span>${opt}</span>
        ${isSelected ? '<span class="check-mark">✓</span>' : ''}
      </div>
    `;
  }).join('');
}

function selectAdditionalMember(name, e) {
  if (e) e.stopPropagation();

  if (currentGroupAdditionalMembers.includes(name)) {
    // If clicked again, remove selection
    currentGroupAdditionalMembers = currentGroupAdditionalMembers.filter(m => m !== name);
  } else {
    // Add selection
    currentGroupAdditionalMembers.push(name);
  }

  // Sync with template if editing existing
  const template = (InnomickData.workflowTemplates || []).find(t => t.id === activeWorkflowTemplateId);
  if (template) {
    template.additionalMembers = [...currentGroupAdditionalMembers];
  }

  renderAdditionalMembersChips(currentGroupAdditionalMembers);
  renderAdditionalMembersDropdown();
}

function renderAdditionalMembersChips(members = []) {
  currentGroupAdditionalMembers = [...members];
  const container = document.getElementById('additionalMembersChips');
  if (!container) return;

  if (!members || members.length === 0) {
    container.innerHTML = '<span class="sp-chips-placeholder" style="color: #64748b; font-size: 13px;">Select Members</span>';
    return;
  }

  container.innerHTML = members.map(m => `
    <span class="sp-member-chip">
      <span>${m}</span>
      <span class="chip-remove-x" onclick="removeAdditionalMemberTag('${m}', event)">✕</span>
    </span>
  `).join('');
}

function removeAdditionalMemberTag(name, e) {
  if (e) e.stopPropagation();
  currentGroupAdditionalMembers = currentGroupAdditionalMembers.filter(m => m !== name);

  const template = (InnomickData.workflowTemplates || []).find(t => t.id === activeWorkflowTemplateId);
  if (template) {
    template.additionalMembers = [...currentGroupAdditionalMembers];
  }

  renderAdditionalMembersChips(currentGroupAdditionalMembers);
  renderAdditionalMembersDropdown();
  showToast(`Removed "${name}" from additional members.`, 'info');
}

function clearAllAdditionalMembers(e) {
  if (e) e.stopPropagation();
  currentGroupAdditionalMembers = [];

  const template = (InnomickData.workflowTemplates || []).find(t => t.id === activeWorkflowTemplateId);
  if (template) {
    template.additionalMembers = [];
  }

  renderAdditionalMembersChips([]);
  renderAdditionalMembersDropdown();
  showToast('Cleared all additional members.', 'info');
}

// Create Workflow Modal
function openCreateWorkflowModal() {
  const modal = document.getElementById('createWorkflowModal');
  if (modal) modal.classList.add('active');
}

function closeCreateWorkflowModal() {
  const modal = document.getElementById('createWorkflowModal');
  if (modal) modal.classList.remove('active');
}

// =============================================================================
// APPROVAL REVIEW & DECISION MODAL
// =============================================================================
function openApprovalModal(id) {
  activeApprovalDocId = id;
  const item = InnomickData.approvalsInbox.find(i => i.id === id);
  if (!item) {
    showToast('Approval item not found.', 'error');
    return;
  }

  const nameEl = document.getElementById('appReviewDocName');
  if (nameEl) nameEl.textContent = item.documentName;
  
  const catEl = document.getElementById('appReviewCategory');
  if (catEl) catEl.textContent = item.category || 'Architecture Blueprints';

  const subEl = document.getElementById('appReviewSubmittedBy');
  if (subEl) subEl.textContent = item.submittedBy || 'MD Faisal Ahmed';

  const deptEl = document.getElementById('appReviewDept');
  if (deptEl) deptEl.textContent = item.department || 'Engineering';

  const dateEl = document.getElementById('appReviewDate');
  if (dateEl) dateEl.textContent = item.submittedDate || '2026-10-14 11:30';

  const stageEl = document.getElementById('appReviewStage');
  if (stageEl) stageEl.textContent = item.currentStage || 'Technical Lead';

  const badgeEl = document.getElementById('appReviewPriorityBadge');
  if (badgeEl) {
    badgeEl.textContent = item.priority || 'Normal';
    badgeEl.className = `status-pill ${item.priority === 'Urgent' ? 'rejected' : 'active'}`;
  }

  const iconEl = document.getElementById('appReviewDocIcon');
  if (iconEl) {
    const ext = item.documentName.split('.').pop().toLowerCase();
    iconEl.innerHTML = getFileIconSvg(ext);
  }

  const commentsEl = document.getElementById('appReviewComments');
  if (commentsEl) commentsEl.value = '';

  const approveRadio = document.querySelector('input[name="appDecisionRadio"][value="Approved"]');
  if (approveRadio) approveRadio.checked = true;

  const modal = document.getElementById('approvalReviewModal');
  if (modal) modal.classList.add('active');
}

function closeApprovalModal() {
  const modal = document.getElementById('approvalReviewModal');
  if (modal) modal.classList.remove('active');
  activeApprovalDocId = null;
}

function previewApprovalDocument() {
  if (!activeApprovalDocId) return;
  const item = InnomickData.approvalsInbox.find(i => i.id === activeApprovalDocId);
  if (item) {
    openPreviewModalWithDoc({
      name: item.documentName,
      category: item.category,
      department: item.department,
      owner: item.submittedBy,
      version: '1.0',
      status: item.currentStage
    });
  }
}

function submitApprovalDecision() {
  if (!activeApprovalDocId) return;
  const selectedRadio = document.querySelector('input[name="appDecisionRadio"]:checked');
  const decision = selectedRadio ? selectedRadio.value : 'Approved';
  const commentsInput = document.getElementById('appReviewComments');
  const comments = (commentsInput && commentsInput.value.trim()) || 
    (decision === 'Approved' ? 'Approved in formal management review.' : 'Returned with comments.');

  processApprovalDecision(activeApprovalDocId, decision, comments);
  closeApprovalModal();
}

// Approval Actions
function quickApprove(id) {
  processApprovalDecision(id, 'Approved', 'Approved in quick review.');
}

function quickReject(id) {
  processApprovalDecision(id, 'Rejected', 'Document does not meet criteria.');
}

function processApprovalDecision(id, decision, comments) {
  const idx = InnomickData.approvalsInbox.findIndex(i => i.id === id);
  if (idx !== -1) {
    const item = InnomickData.approvalsInbox.splice(idx, 1)[0];
    const docName = item.name || item.documentName || 'Document';

    if (decision === 'Approved') {
      if (!InnomickData.recentDocuments) InnomickData.recentDocuments = [];
      const existingIdx = InnomickData.recentDocuments.findIndex(d => d.id === item.id || d.name === docName);
      if (existingIdx !== -1) InnomickData.recentDocuments.splice(existingIdx, 1);

      InnomickData.recentDocuments.unshift({
        id: item.id || 'recent-' + Date.now(),
        name: docName,
        type: docName.endsWith('.pdf') ? 'pdf' : (docName.endsWith('.docx') ? 'docx' : 'pdf'),
        uploadDate: '4 minutes ago',
        lastAccessed: '4 minutes ago',
        modified: '4 minutes ago',
        location: item.location || 'My Department',
        owner: item.owner || item.requestedBy || item.submittedBy || 'Mohammed Wasifuddin',
        category: item.category || 'Accreditation Documents',
        subCategory: item.tagSubCategory || item.subCategory || '-',
        department: item.department || 'Information Technology',
        favorite: false,
        isCheckedOut: false,
        status: 'Approved'
      });
    }

    InnomickData.approvalsHistory.unshift({
      id: 'hist-' + Date.now(),
      documentName: docName,
      name: docName,
      actionBy: (InnomickData.currentUser && InnomickData.currentUser.name) || 'Mohammed Wasifuddin',
      response: decision,
      fromStage: item.currentStage,
      toStage: decision === 'Approved' ? 'Completed & Published' : 'Returned to Author',
      transitionKey: decision === 'Approved' ? 'STAGE-APPROVE' : 'STAGE-REJECT',
      completedDate: new Date().toISOString().replace('T', ' ').substring(0, 16),
      comments: comments
    });

    updateApprovalsLeftBadges();
    showToast(`Document "${docName}" has been ${decision.toLowerCase()}.`);
    renderCurrentView();
  }
}

// Document Preview Modal
function previewDocument(id) {
  const doc = findAnyDoc(id) || InnomickData.documents[0];
  openPreviewModalWithDoc(doc);
}

function previewMyFile(id) {
  const doc = InnomickData.myFiles.find(d => d.id === id) || InnomickData.myFiles[0];
  openPreviewModalWithDoc(doc);
}

function previewSharedFile(id) {
  const doc = InnomickData.sharedFiles.find(d => d.id === id) || InnomickData.sharedFiles[0];
  openPreviewModalWithDoc({
    ...doc,
    owner: doc.sharedBy,
    uploadDate: doc.dateShared,
    version: '1.0',
    category: 'Shared Resource'
  });
}

function previewExpiringFile(id) {
  const doc = InnomickData.expiringFiles.find(d => d.id === id) || InnomickData.expiringFiles[0];
  openPreviewModalWithDoc({
    ...doc,
    uploadDate: '2026-01-15',
    version: '1.0'
  });
}

// =============================================================================
// RICH DOCUMENT & PDF VIEWER CONTROLLER
// =============================================================================
let currentViewerDoc = null;
let currentViewerPage = 1;
let currentViewerZoom = 1.0;

function openPreviewModalWithDoc(doc) {
  if (!doc) {
    doc = InnomickData.documents[0] || {
      id: 'DOC-101',
      name: 'My Dept PDF.pdf',
      type: 'pdf',
      category: 'Budget & Financial Plans',
      department: 'Central Supply Chain Management',
      status: 'Active',
      owner: 'Sayed Asif Ahmed',
      referenceNumber: '454522',
      uploadDate: '11 October 2026',
      version: '1.0'
    };
  }

  currentViewerDoc = doc;
  currentViewerPage = 1;
  currentViewerZoom = 1.0;

  const modal = document.getElementById('previewModal');
  if (!modal) return;

  const titleEl = document.getElementById('previewModalTitle');
  if (titleEl) titleEl.innerText = doc.name;

  const iconEl = document.getElementById('previewModalIcon');
  if (iconEl) iconEl.innerHTML = getFileIconSvg(doc.type || 'pdf');

  const catEl = document.getElementById('previewDocCategory');
  if (catEl) catEl.innerText = doc.category || 'Standard Documentation';

  const deptEl = document.getElementById('previewDocDept');
  if (deptEl) deptEl.innerText = doc.department || 'Central Supply Chain Management';

  const verEl = document.getElementById('previewDocVersion');
  if (verEl) verEl.innerText = `v${doc.version || '1.0'}`;

  const statusEl = document.getElementById('previewDocStatus');
  if (statusEl) statusEl.innerText = doc.status || 'Active';

  const ownerEl = document.getElementById('previewDocOwner');
  if (ownerEl) ownerEl.innerText = doc.owner || InnomickData.currentUser.name;

  const pageIndicator = document.getElementById('viewerPageIndicator');
  if (pageIndicator) pageIndicator.innerText = `Page 1 of 2`;

  const zoomIndicator = document.getElementById('viewerZoomIndicator');
  if (zoomIndicator) zoomIndicator.innerText = `100%`;

  renderDocViewerSheet();

  modal.classList.add('active');
}

function closePreviewModal() {
  const modal = document.getElementById('previewModal');
  if (modal) {
    modal.classList.remove('active');
  }
}

function changeViewerPage(delta) {
  const newPage = currentViewerPage + delta;
  if (newPage < 1 || newPage > 2) return;
  currentViewerPage = newPage;
  const pageIndicator = document.getElementById('viewerPageIndicator');
  if (pageIndicator) pageIndicator.innerText = `Page ${currentViewerPage} of 2`;
  renderDocViewerSheet();
}

function adjustViewerZoom(delta) {
  currentViewerZoom = Math.min(1.5, Math.max(0.6, Math.round((currentViewerZoom + delta) * 10) / 10));
  const zoomIndicator = document.getElementById('viewerZoomIndicator');
  if (zoomIndicator) zoomIndicator.innerText = `${Math.round(currentViewerZoom * 100)}%`;
  const container = document.getElementById('viewerSheetContainer');
  if (container) {
    container.style.transform = `scale(${currentViewerZoom})`;
    container.style.transformOrigin = 'top center';
  }
}

function renderDocViewerSheet() {
  const body = document.getElementById('previewModalViewerBody');
  if (!body) return;

  const doc = currentViewerDoc || {
    name: 'My Dept PDF.pdf',
    type: 'pdf',
    category: 'Budget & Financial Plans',
    department: 'Central Supply Chain Management',
    status: 'Active',
    owner: 'Sayed Asif Ahmed',
    referenceNumber: '454522',
    uploadDate: '11 October 2026'
  };

  const isPage1 = (currentViewerPage === 1);
  const isXlsx = (doc.type === 'xlsx' || doc.name.endsWith('.xlsx'));
  const isPptx = (doc.type === 'pptx' || doc.name.endsWith('.pptx'));

  let sheetContent = '';

  if (isXlsx) {
    sheetContent = `
      <div style="font-family: 'Segoe UI', Tahoma, sans-serif; font-size: 13px; color: #1e293b;">
        <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #107c41; padding-bottom: 12px; margin-bottom: 18px;">
          <div style="display: flex; align-items: center; gap: 10px;">
            <div style="background: #107c41; color: white; padding: 6px 10px; border-radius: 4px; font-weight: bold; font-size: 14px;">XLSX</div>
            <div>
              <div style="font-weight: 700; font-size: 16px; color: #0f172a;">${doc.name}</div>
              <div style="color: #64748b; font-size: 11.5px;">Worksheet: <strong>Department Allocations & Expenditure 2026</strong></div>
            </div>
          </div>
          <div style="text-align: right; font-size: 11.5px; color: #64748b;">
            <div>Ref: <strong>${doc.referenceNumber || 'INM-9914'}</strong></div>
            <div>Owner: <strong>${doc.owner || 'Mohammed Wasifuddin'}</strong></div>
          </div>
        </div>

        <table style="width: 100%; border-collapse: collapse; font-size: 12.5px; margin-bottom: 20px;">
          <thead>
            <tr style="background: #f1f5f9; color: #334155; border: 1px solid #cbd5e1;">
              <th style="padding: 8px 10px; border: 1px solid #cbd5e1; text-align: left; width: 40px;">#</th>
              <th style="padding: 8px 10px; border: 1px solid #cbd5e1; text-align: left;">Budget Line / Cost Center</th>
              <th style="padding: 8px 10px; border: 1px solid #cbd5e1; text-align: left;">Department</th>
              <th style="padding: 8px 10px; border: 1px solid #cbd5e1; text-align: right;">Q1 Allocated</th>
              <th style="padding: 8px 10px; border: 1px solid #cbd5e1; text-align: right;">Q2 Committed</th>
              <th style="padding: 8px 10px; border: 1px solid #cbd5e1; text-align: right;">Variance (%)</th>
              <th style="padding: 8px 10px; border: 1px solid #cbd5e1; text-align: center;">Status</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; font-weight: bold; color: #64748b;">1</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; font-weight: 600;">Enterprise Cloud Storage Infrastructure</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0;">Information Technology</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: right;">$450,000.00</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: right;">$428,500.00</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: right; color: #16a34a;">+4.7%</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: center;"><span style="background: #dcfce7; color: #15803d; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600;">Cleared</span></td>
            </tr>
            <tr style="background: #f8fafc;">
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; font-weight: bold; color: #64748b;">2</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; font-weight: 600;">Strategic Sourcing & Vendor Logistics</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0;">Central Supply Chain</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: right;">$820,000.00</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: right;">$790,000.00</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: right; color: #16a34a;">+3.6%</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: center;"><span style="background: #dcfce7; color: #15803d; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600;">Cleared</span></td>
            </tr>
            <tr>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; font-weight: bold; color: #64748b;">3</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; font-weight: 600;">Contractor SOW Compliance Audits</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0;">Central Supply Chain</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: right;">$175,000.00</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: right;">$164,200.00</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: right; color: #16a34a;">+6.1%</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: center;"><span style="background: #fef3c7; color: #b45309; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600;">Auditing</span></td>
            </tr>
            <tr style="background: #f8fafc;">
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; font-weight: bold; color: #64748b;">4</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; font-weight: 600;">Information Security & SOC-2 Certification</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0;">Information Technology</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: right;">$310,000.00</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: right;">$298,000.00</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: right; color: #16a34a;">+3.8%</td>
              <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: center;"><span style="background: #dcfce7; color: #15803d; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600;">Cleared</span></td>
            </tr>
            <tr style="background: #f1f5f9; font-weight: 700;">
              <td colspan="3" style="padding: 8px 10px; border: 1px solid #cbd5e1; text-align: right;">Consolidated Total:</td>
              <td style="padding: 8px 10px; border: 1px solid #cbd5e1; text-align: right; color: #008272;">$1,755,000.00</td>
              <td style="padding: 8px 10px; border: 1px solid #cbd5e1; text-align: right; color: #008272;">$1,680,700.00</td>
              <td style="padding: 8px 10px; border: 1px solid #cbd5e1; text-align: right; color: #16a34a;">+4.2%</td>
              <td style="padding: 8px 10px; border: 1px solid #cbd5e1; text-align: center;">Active</td>
            </tr>
          </tbody>
        </table>

        <div style="background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 4px; padding: 12px; font-size: 12px; color: #166534; display: flex; align-items: center; justify-content: space-between;">
          <span>&check; Audit Verification: Formula validation completed with zero recalculation exceptions.</span>
          <span style="font-weight: 600;">Sheet: Sheet1 (Allocations)</span>
        </div>
      </div>
    `;
  } else if (isPptx) {
    sheetContent = `
      <div style="font-family: 'Segoe UI', Tahoma, sans-serif; color: #1e293b;">
        <div style="background: linear-gradient(135deg, #008272 0%, #004d40 100%); color: white; padding: 28px; border-radius: 6px; margin-bottom: 24px;">
          <div style="font-size: 11px; text-transform: uppercase; letter-spacing: 1.5px; opacity: 0.85; margin-bottom: 6px;">Executive Briefing Presentation</div>
          <h2 style="margin: 0 0 10px 0; font-size: 22px; font-weight: 700; color: #ffffff;">${doc.name.replace('.pptx', '')}</h2>
          <div style="font-size: 13px; opacity: 0.9;">Innomick Technologies DMS &bull; ${doc.department || 'Corporate Management'}</div>
        </div>

        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 24px;">
          <div style="border: 1px solid #e2e8f0; border-radius: 6px; padding: 16px; background: #f8fafc;">
            <div style="font-weight: 700; font-size: 14px; color: #0f172a; margin-bottom: 8px; display: flex; align-items: center; gap: 6px;">
              <span style="color: #008272;">●</span> Key Objectives & Milestones
            </div>
            <ul style="margin: 0; padding-left: 18px; font-size: 12.5px; color: #475569; line-height: 1.6;">
              <li>Centralize cross-department enterprise assets in secure DMS vault.</li>
              <li>Enforce single-click governance, check-in locking, and version tracking.</li>
              <li>Automate corporate group workflows and multi-stage managerial approvals.</li>
            </ul>
          </div>

          <div style="border: 1px solid #e2e8f0; border-radius: 6px; padding: 16px; background: #f8fafc;">
            <div style="font-weight: 700; font-size: 14px; color: #0f172a; margin-bottom: 8px; display: flex; align-items: center; gap: 6px;">
              <span style="color: #008272;">●</span> Target Deliverables
            </div>
            <ul style="margin: 0; padding-left: 18px; font-size: 12.5px; color: #475569; line-height: 1.6;">
              <li>Full SharePoint-level 13-column metadata table with live search filters.</li>
              <li>Automated document retention & expiring file notifications.</li>
              <li>Role-based access controls and electronic audit signatures.</li>
            </ul>
          </div>
        </div>

        <div style="border-top: 1px solid #cbd5e1; padding-top: 14px; display: flex; justify-content: space-between; font-size: 11.5px; color: #64748b;">
          <span>Presenter: ${doc.owner || 'Mohammed Wasifuddin'}</span>
          <span>Slide ${currentViewerPage} of 2</span>
        </div>
      </div>
    `;
  } else {
    // Standard Authentic Multi-page PDF Document
    if (isPage1) {
      sheetContent = `
        <div style="font-family: 'Segoe UI', Tahoma, sans-serif; color: #0f172a; line-height: 1.55;">
          <!-- Top Corporate Header -->
          <div style="display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 2px solid #008272; padding-bottom: 16px; margin-bottom: 20px;">
            <div>
              <div style="display: flex; align-items: center; gap: 8px;">
                <div style="width: 24px; height: 24px; background: #008272; border-radius: 4px; display: flex; align-items: center; justify-content: center; color: #ffffff; font-weight: bold; font-size: 13px;">I</div>
                <span style="font-weight: 800; font-size: 16px; color: #008272; letter-spacing: 0.5px;">INNOMICK TECHNOLOGIES</span>
              </div>
              <div style="font-size: 11px; color: #64748b; margin-top: 4px; letter-spacing: 0.3px;">ENTERPRISE DOCUMENT MANAGEMENT &bull; OFFICIAL VAULT RECORD</div>
            </div>
            <div style="text-align: right;">
              <div style="background: #e6f3f1; color: #008272; font-size: 11px; font-weight: 700; padding: 4px 10px; border-radius: 3px; display: inline-block;">
                OFFICIAL RECORD &bull; CLASSIFIED
              </div>
              <div style="font-size: 11.5px; color: #64748b; margin-top: 5px;">REF NO: <strong>${doc.referenceNumber || '454522'}</strong></div>
            </div>
          </div>

          <!-- Document Title & Header Meta -->
          <div style="margin-bottom: 22px;">
            <h1 style="font-size: 21px; font-weight: 700; color: #0f172a; margin: 0 0 8px 0; line-height: 1.3;">${doc.name}</h1>
            <div style="font-size: 12.5px; color: #475569;">
              Department: <strong>${doc.department || 'Central Supply Chain Management'}</strong> &bull; 
              Category: <strong>${doc.category || 'Budget & Financial Plans'}</strong>
            </div>
          </div>

          <!-- Summary Meta Box -->
          <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 14px 18px; margin-bottom: 24px; display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; font-size: 12px;">
            <div>
              <div style="color: #64748b; font-size: 11px; text-transform: uppercase;">Owner / Author</div>
              <div style="font-weight: 600; color: #0f172a; margin-top: 2px;">${doc.owner || 'Sayed Asif Ahmed'}</div>
            </div>
            <div>
              <div style="color: #64748b; font-size: 11px; text-transform: uppercase;">Stakeholder</div>
              <div style="font-weight: 600; color: #0f172a; margin-top: 2px;">${doc.stakeHolder || 'Navin Patel'}</div>
            </div>
            <div>
              <div style="color: #64748b; font-size: 11px; text-transform: uppercase;">Upload / Effective</div>
              <div style="font-weight: 600; color: #0f172a; margin-top: 2px;">${doc.uploadDate || '11 October 2026'}</div>
            </div>
            <div>
              <div style="color: #64748b; font-size: 11px; text-transform: uppercase;">Governance Status</div>
              <div style="font-weight: 600; color: #16a34a; margin-top: 2px;">● ${doc.status || 'Active'} (Verified)</div>
            </div>
          </div>

          <!-- Section 1 -->
          <div style="margin-bottom: 20px;">
            <div style="font-size: 13.5px; font-weight: 700; color: #008272; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px;">1. EXECUTIVE SUMMARY & OBJECTIVES</div>
            <p style="font-size: 12.5px; color: #334155; margin: 0 0 8px 0; text-align: justify;">
              This official document sets forth the comprehensive governance framework, procurement milestones, and budgetary allocations overseen by the Central Supply Chain Management and Corporate Operations steering committee. All activities outlined herein have undergone compliance review in accordance with Innomick Information Security and ISO 27001 standards.
            </p>
          </div>

          <!-- Section 2: Financial Allocation Table -->
          <div style="margin-bottom: 22px;">
            <div style="font-size: 13.5px; font-weight: 700; color: #008272; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px;">2. ALLOCATIONS & FISCAL DISBURSEMENTS (FY2026)</div>
            <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
              <thead>
                <tr style="background: #f1f5f9; color: #334155; border: 1px solid #cbd5e1;">
                  <th style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: left;">Cost Center Category</th>
                  <th style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: right;">Approved Allocation</th>
                  <th style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: right;">Utilized to Date</th>
                  <th style="padding: 7px 10px; border: 1px solid #cbd5e1; text-align: center;">Audit Clearance</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td style="padding: 7px 10px; border: 1px solid #e2e8f0; font-weight: 600;">Vendor Contracts & Strategic Procurement</td>
                  <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: right;">$840,000.00</td>
                  <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: right;">$792,500.00</td>
                  <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: center; color: #16a34a; font-weight: 600;">&check; Cleared</td>
                </tr>
                <tr style="background: #f8fafc;">
                  <td style="padding: 7px 10px; border: 1px solid #e2e8f0; font-weight: 600;">Cloud Infrastructure & DMS License Subscriptions</td>
                  <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: right;">$420,000.00</td>
                  <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: right;">$395,000.00</td>
                  <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: center; color: #16a34a; font-weight: 600;">&check; Cleared</td>
                </tr>
                <tr>
                  <td style="padding: 7px 10px; border: 1px solid #e2e8f0; font-weight: 600;">Technical Support & Service Level Guarantees</td>
                  <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: right;">$190,000.00</td>
                  <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: right;">$178,200.00</td>
                  <td style="padding: 7px 10px; border: 1px solid #e2e8f0; text-align: center; color: #16a34a; font-weight: 600;">&check; Cleared</td>
                </tr>
              </tbody>
            </table>
          </div>

          <!-- Document Watermark & Footer -->
          <div style="border-top: 1px solid #cbd5e1; padding-top: 12px; margin-top: 24px; display: flex; justify-content: space-between; align-items: center; font-size: 11px; color: #64748b;">
            <span>Innomick Technologies DMS v2.4 &bull; Official Digital Record</span>
            <span>Page 1 of 2</span>
          </div>
        </div>
      `;
    } else {
      // Page 2
      sheetContent = `
        <div style="font-family: 'Segoe UI', Tahoma, sans-serif; color: #0f172a; line-height: 1.55;">
          <!-- Top Header -->
          <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #008272; padding-bottom: 12px; margin-bottom: 20px;">
            <div style="font-weight: 700; font-size: 14px; color: #008272;">INNOMICK TECHNOLOGIES DMS &bull; AUDIT & DIGITAL SIGNATURES</div>
            <div style="font-size: 11px; color: #64748b;">REF: <strong>${doc.referenceNumber || '454522'}</strong></div>
          </div>

          <!-- Section 3: Terms & Approvals -->
          <div style="margin-bottom: 20px;">
            <div style="font-size: 13.5px; font-weight: 700; color: #008272; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px;">3. COMPLIANCE & LEGAL COVENANTS</div>
            <p style="font-size: 12px; color: #334155; margin: 0 0 10px 0; text-align: justify;">
              All provisions described in this schedule are legally binding upon execution. Any amendments or version revisions require formal re-routing through the Innomick Document Management Corporate Approval workflow. Unauthorized alteration constitutes a breach of enterprise policy.
            </p>
          </div>

          <!-- Section 4: Audit Verification Trail -->
          <div style="margin-bottom: 22px;">
            <div style="font-size: 13.5px; font-weight: 700; color: #008272; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px;">4. AUDIT & REVISION LOG</div>
            <table style="width: 100%; border-collapse: collapse; font-size: 11.5px;">
              <thead>
                <tr style="background: #f1f5f9; color: #334155; border: 1px solid #cbd5e1;">
                  <th style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: left;">Version</th>
                  <th style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: left;">Action Taken</th>
                  <th style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: left;">Modified By</th>
                  <th style="padding: 6px 10px; border: 1px solid #cbd5e1; text-align: left;">Timestamp</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td style="padding: 6px 10px; border: 1px solid #e2e8f0; font-weight: bold; color: #008272;">v1.0</td>
                  <td style="padding: 6px 10px; border: 1px solid #e2e8f0;">Initial Draft Creation & Submission</td>
                  <td style="padding: 6px 10px; border: 1px solid #e2e8f0;">Sayed Asif Ahmed</td>
                  <td style="padding: 6px 10px; border: 1px solid #e2e8f0; color: #64748b;">11 Oct 2026 09:15 AM</td>
                </tr>
                <tr style="background: #f8fafc;">
                  <td style="padding: 6px 10px; border: 1px solid #e2e8f0; font-weight: bold; color: #008272;">v1.1</td>
                  <td style="padding: 6px 10px; border: 1px solid #e2e8f0;">Stakeholder Review & Clearance</td>
                  <td style="padding: 6px 10px; border: 1px solid #e2e8f0;">Navin Patel</td>
                  <td style="padding: 6px 10px; border: 1px solid #e2e8f0; color: #64748b;">11 Oct 2026 14:30 PM</td>
                </tr>
                <tr>
                  <td style="padding: 6px 10px; border: 1px solid #e2e8f0; font-weight: bold; color: #008272;">v1.2</td>
                  <td style="padding: 6px 10px; border: 1px solid #e2e8f0;">Final Corporate Approval Sign-off</td>
                  <td style="padding: 6px 10px; border: 1px solid #e2e8f0;">Mohammed Wasifuddin</td>
                  <td style="padding: 6px 10px; border: 1px solid #e2e8f0; color: #64748b;">12 Oct 2026 11:00 AM</td>
                </tr>
              </tbody>
            </table>
          </div>

          <!-- Section 5: Electronic Signatures -->
          <div style="margin-bottom: 24px;">
            <div style="font-size: 13.5px; font-weight: 700; color: #008272; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 12px;">5. AUTHORIZED ELECTRONIC SIGNATURES</div>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px;">
              <div style="border: 1px dashed #008272; border-radius: 6px; padding: 14px; background: #f0fdf4;">
                <div style="font-size: 11px; text-transform: uppercase; color: #166534; font-weight: 700; letter-spacing: 0.5px; margin-bottom: 6px;">DIGITALLY VERIFIED SIGNATURE</div>
                <div style="font-family: 'Brush Script MT', cursive, sans-serif; font-size: 24px; color: #0f172a; margin-bottom: 4px;">Sayed Asif Ahmed</div>
                <div style="font-size: 11.5px; font-weight: 600; color: #334155;">Sayed Asif Ahmed</div>
                <div style="font-size: 11px; color: #64748b;">Head of Supply Chain Management</div>
                <div style="font-size: 10.5px; color: #15803d; margin-top: 4px;">&check; SHA-256 Certificate: INM-SEC-8829-OK</div>
              </div>

              <div style="border: 1px dashed #008272; border-radius: 6px; padding: 14px; background: #f0fdf4;">
                <div style="font-size: 11px; text-transform: uppercase; color: #166534; font-weight: 700; letter-spacing: 0.5px; margin-bottom: 6px;">DIGITALLY VERIFIED SIGNATURE</div>
                <div style="font-family: 'Brush Script MT', cursive, sans-serif; font-size: 24px; color: #0f172a; margin-bottom: 4px;">Mohammed Wasifuddin</div>
                <div style="font-size: 11.5px; font-weight: 600; color: #334155;">Mohammed Wasifuddin</div>
                <div style="font-size: 11px; color: #64748b;">Operations & IT Director</div>
                <div style="font-size: 10.5px; color: #15803d; margin-top: 4px;">&check; SHA-256 Certificate: INM-SEC-9941-OK</div>
              </div>
            </div>
          </div>

          <!-- Bottom Footer -->
          <div style="border-top: 1px solid #cbd5e1; padding-top: 12px; margin-top: 24px; display: flex; justify-content: space-between; align-items: center; font-size: 11px; color: #64748b;">
            <span>Innomick Technologies DMS v2.4 &bull; Official Digital Record</span>
            <span>Page 2 of 2 (End of Document)</span>
          </div>
        </div>
      `;
    }
  }

  body.innerHTML = `
    <div id="viewerSheetContainer" style="width: 720px; min-height: 860px; background: #ffffff; box-shadow: 0 4px 24px rgba(0,0,0,0.3); padding: 44px; border-radius: 2px; box-sizing: border-box; transform: scale(${currentViewerZoom}); transform-origin: top center; transition: transform 0.12s ease; margin-bottom: 24px;">
      ${sheetContent}
    </div>
  `;
}

function triggerFileDownload(fileName, doc) {
  let docObj = doc;
  if (!docObj && fileName) {
    docObj = findAnyDoc(fileName) || currentViewerDoc;
  }
  const name = (docObj && docObj.name) ? docObj.name : (fileName || 'Document.docx');
  const ext = (name.split('.').pop() || 'docx').toLowerCase();

  let mimeType = 'application/octet-stream';
  let content = '';

  if (ext === 'docx' || ext === 'doc') {
    mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    content = `INNOMICK TECHNOLOGIES • ENTERPRISE DOCUMENT MANAGEMENT SYSTEM
================================================================================
DOCUMENT NAME: ${name}
CATEGORY: ${docObj?.category || 'Accreditation Documents'}
SUB CATEGORY: ${docObj?.subCategory || 'ISO Operational Standards'}
DEPARTMENT: ${docObj?.department || 'Finance & Accounts'}
LOCATION: ${docObj?.location || 'My Organization'}
OWNER / CREATED BY: ${docObj?.owner || 'MD Faisal Ahmed'}
REFERENCE NUMBER: ${docObj?.referenceNumber || 'ACC-DOC-101'}
STAKE HOLDER: ${docObj?.stakeHolder || 'Quality Accreditation Board'}
VENDOR: ${docObj?.vendor || 'Innomick Technologies Ltd.'}
UPLOAD DATE: ${docObj?.uploadDate || '2026-10-11'}
GOVERNANCE STATUS: Active (Verified)
DESCRIPTION: ${docObj?.description || 'Operational accreditation testing protocols and operational verification benchmarks.'}
================================================================================

1. EXECUTIVE SUMMARY & OBJECTIVES
This official enterprise document sets forth the comprehensive governance framework, procedural milestones, and budgetary allocations overseen by the Central Steering Committee and Operations Council. All activities outlined herein have undergone compliance review in accordance with ISO 27001, SOC 2 Type II, and corporate security guidelines.

2. DEPARTMENTAL RESPONSIBILITY
Primary Department: ${docObj?.department || 'Finance & Accounts'}.
The designated document owner and administrative custodian is ${docObj?.owner || 'MD Faisal Ahmed'}. Any revisions or amendments must be submitted through the standard DMS Approval Workflow prior to publication.

3. OPERATIONAL AUDIT SPECIFICATIONS
- Security Certification: ISO 27001 / SOC 2 Type II
- Document Classification: Confidential / Proprietary
- Retention Policy: 7 Years (Statutory Digital Vault)
- Cryptographic Verification: SHA-256 Checksum Verified

4. AUTHORIZED SIGN-OFF
Prepared by: ${docObj?.owner || 'MD Faisal Ahmed'}
Department: ${docObj?.department || 'Finance & Accounts'}
Status: Electronically Approved & Distributed
`;
  } else if (ext === 'xlsx' || ext === 'xls' || ext === 'csv') {
    mimeType = ext === 'csv' ? 'text/csv;charset=utf-8;' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    content = `Location,Owner,Category,Sub Category,Department,Reference Number,Stake Holder,Vendor,Upload Date,Description\n` +
      `"${docObj?.location || 'My Organization'}","${docObj?.owner || 'MD Faisal Ahmed'}","${docObj?.category || 'Accreditation Documents'}","${docObj?.subCategory || 'QA Testing'}","${docObj?.department || 'Information Technology'}","${docObj?.referenceNumber || 'QA-FIX-123'}","${docObj?.stakeHolder || 'Steering Committee'}","${docObj?.vendor || 'Innomick Technologies Ltd.'}","${docObj?.uploadDate || '2026-10-11'}","${docObj?.description || 'Enterprise spreadsheet record'}"\n` +
      `"My Organization","Mohammed Wasifuddin","Policies and Procedures","QA Testing Results","Information Technology","QA-FIX-123","Software Quality Assurance Council","Innomick Technologies Ltd.","2026-09-30","Validation matrix for developer sprint bug fixes and retest verification test runs"\n` +
      `"My Organization","Vikram Singh","Accreditation Documents","Supply Chain Reconciliation","Central Supply Chain Management","SCN-VER-405","Procurement Committee","Oracle Supply Cloud","2026-09-23","Operational inventory reconciliation workbook and vendor voucher audit test suite"\n`;
  } else if (ext === 'pdf') {
    mimeType = 'application/pdf';
    content = `%PDF-1.4\n%âãÏÓ\n1 0 obj\n<< /Title (${name}) /Author (${docObj?.owner || 'Innomick Technologies'}) /Subject (${docObj?.category || 'DMS Document'}) >>\nendobj\n` +
      `INNOMICK TECHNOLOGIES • ENTERPRISE DMS PDF EXPORT\n` +
      `Document: ${name}\n` +
      `Category: ${docObj?.category || 'Accreditation Documents'}\n` +
      `Sub Category: ${docObj?.subCategory || 'Operational Standards'}\n` +
      `Department: ${docObj?.department || 'Corporate Management'}\n` +
      `Owner: ${docObj?.owner || 'Mohammed Wasifuddin'}\n` +
      `Reference: ${docObj?.referenceNumber || 'REF-2026-089'}\n` +
      `Stakeholder: ${docObj?.stakeHolder || 'Enterprise Steering Committee'}\n` +
      `Vendor: ${docObj?.vendor || 'Innomick Technologies Ltd.'}\n` +
      `Date: ${docObj?.uploadDate || '2026-10-18'}\n` +
      `Description: ${docObj?.description || 'Official enterprise compliance documentation.'}\n` +
      `Status: Verified\n%%EOF`;
  } else {
    mimeType = 'text/plain;charset=utf-8';
    content = `INNOMICK TECHNOLOGIES DMS\nDocument: ${name}\nCategory: ${docObj?.category || 'Accreditation Documents'}\nOwner: ${docObj?.owner || 'Admin'}\nDate: ${docObj?.uploadDate || '2026-10-18'}\nDescription: ${docObj?.description || 'Exported document'}`;
  }

  const blob = new Blob([content], { type: mimeType });
  const downloadUrl = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = downloadUrl;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(downloadUrl), 2000);

  showToast(`Downloaded "${name}" successfully.`, 'success');
}

function downloadDocument(id) {
  closeContextMenu();
  let doc = null;
  if (id === 'current') {
    doc = currentViewerDoc;
  } else if (id) {
    doc = findAnyDoc(id);
  }
  if (!doc && currentViewerDoc) {
    doc = currentViewerDoc;
  }
  const fileName = doc ? doc.name : (typeof id === 'string' && id.includes('.') ? id : 'Document.docx');
  triggerFileDownload(fileName, doc);
}

function initiateRenewal(id) {
  const doc = InnomickData.expiringFiles.find(d => d.id === id);
  if (doc) {
    doc.daysLeft += 365;
    doc.status = 'Renewed';
    showToast(`Document renewal requested for "${doc.name}". Expiry extended by 1 year!`);
    renderCurrentView();
  }
}

// Toast System
function showToast(message, type = 'success') {
  const container = document.getElementById('toastContainer');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = `dms-toast ${type}`;
  toast.innerHTML = `
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
      ${type === 'success' 
        ? '<polyline points="20 6 9 17 4 12"/>' 
        : '<circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>'}
    </svg>
    <span>${message}</span>
  `;

  container.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transition = 'opacity 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}

// Initialize on DOM Ready
document.addEventListener('DOMContentLoaded', () => {
  loadPersistedWorkflowTemplates();
  initApprovalEngine();
  restorePersistedPersona();
  initRouter();

  // Test Harness / Screenshot Automation Helper
  const params = new URLSearchParams(window.location.search);
  const testModal = params.get('testModal');
  if (testModal) {
    const modalEl = document.getElementById('uploadModal');
    if (modalEl) modalEl.style.transition = 'none';
    const dialogEl = document.querySelector('#uploadModal .dms-modal-dialog');
    if (dialogEl) dialogEl.style.transition = 'none';

    setTimeout(() => {
      if (testModal === 'assessment') {
        openUploadModal('Assessment Results');
      } else if (testModal === 'asset') {
        openUploadModal('Asset Inventories');
      } else if (testModal === 'audit') {
        openUploadModal('Audit Documents');
      } else if (testModal === 'step2') {
        openUploadModal('Assessment Results');
        handleUploadStep1Next(new Event('submit'));
      } else if (testModal === 'step3_individual') {
        openUploadModal('Assessment Results');
        handleUploadStep1Next(new Event('submit'));
        showSelectIndividualWorkflow();
      } else if (testModal === 'step3_group') {
        openUploadModal('Assessment Results');
        handleUploadStep1Next(new Event('submit'));
        showSelectGroupWorkflow();
      } else if (testModal === 'step4_group') {
        openUploadModal('Assessment Results');
        handleUploadStep1Next(new Event('submit'));
        showSelectGroupWorkflow();
        openWorkflowConfigForm('group');
      } else if (testModal === 'uploaded_row') {
        openUploadModal('Accreditation Documents');
        handleUploadStep1Next(new Event('submit'));
        executeSkipWorkflowUpload();
      }
    }, 200);
  }
});
