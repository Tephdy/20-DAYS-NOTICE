const ROOT_FOLDER_ID = '16qLfbrPeU-H_UH0QUl2lIqUoqf5_nVLQ';
const MAIN_SPREADSHEET_ID = '16dehNzBJq2h_6uF4pPj4OhTJ_9Ja2QYYjt__K82Nz6M';
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbxj5XbaBTofwthjYumNdlMOw0azJVsfGNqmbRD-Uykn4qnSKYdAZJd8vX4nGKFGTiDu/exec'; // <-- Paste your Apps Script Web App /exec URL here
const MAIN_SPREADSHEET_TABS = ['ECO888', 'GREEN', 'ADI168', 'KALAYAAN888', 'PLEASANT', 'HOMEY', 'PENTHAUZ', 'DREAM', 'B&B'];
const LOCAL_STORAGE_KEY = `lease_renewals_direct_cache_${MAIN_SPREADSHEET_ID}`;

let appData = {};
let activeTab = null;
const selectedTenants = new Set();
let workbookCache = null;
let currentEditingTenant = null;
let isAddMode = false;

const tabs = document.querySelectorAll('.office-tab');
const contents = document.querySelectorAll('.ribbon-content');

tabs.forEach(tab => {
  tab.addEventListener('click', () => {
    tabs.forEach(t => t.classList.remove('active'));
    contents.forEach(c => c.classList.add('hidden'));
    tab.classList.add('active');
    document.getElementById(`ribbon-${tab.dataset.target}`).classList.remove('hidden');
  });
});

function persistLocalCache() {
  try {
    localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(appData));
  } catch (e) {
    console.warn('Could not cache spreadsheet data locally:', e);
  }
}

async function postToAppsScript(payload) {
  if (!WEB_APP_URL || WEB_APP_URL === 'YOUR_APPS_SCRIPT_WEB_APP_URL_HERE') {
    console.warn('WEB_APP_URL is not configured; changes were saved locally only.');
    return { success: false, localOnly: true };
  }

  const response = await fetch(WEB_APP_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(payload)
  });

  const text = await response.text();
  try {
    const result = JSON.parse(text);
    if (result && result.success === false) {
      throw new Error(result.error || 'Apps Script returned an error.');
    }
    return result;
  } catch (err) {
    if (!response.ok) throw new Error(`HTTP ${response.status}: ${text}`);
    return { success: true };
  }
}

function buildTenantSheetPayload(action, tabName, tenant, extra = {}) {
  const lockedTermLabel = String(tenant.termLabel || '').trim();
  return {
    action,
    tabName,
    tenantName: tenant.tenantName || '',
    unit: tenant.unit || '',
    building: tenant.building || tabName,
    computed20th: tenant.computed20th || '',
    computedDueDate: tenant.computed20th || '',
    rent: Number(tenant.rent) || 0,
    status: tenant.status || 'Pending',
    dateSent: tenant.dateSent || '',
    pdfUrl: tenant.pdfUrl || '',
    term: lockedTermLabel,
    termLabel: lockedTermLabel,
    reason: tenant.reason || '—',
    startDate: tenant.startDate || '',
    endContract: tenant.endContract || '',
    ...extra
  };
}

async function syncToDatabase(options = {}) {
  persistLocalCache();

  const { action, tabName, tenant, originalName, originalUnit } = options;
  if (!action || !tabName) {
    return { success: true };
  }

  if (action === 'clearTab') {
    return await postToAppsScript({ action: 'clearTab', tabName });
  }

  if ((action === 'save' || action === 'update' || action === 'delete') && tenant) {
    const payload = buildTenantSheetPayload(action, tabName, tenant, {
      originalName: originalName || tenant.tenantName,
      originalUnit: originalUnit || tenant.unit
    });
    return await postToAppsScript(payload);
  }

  return { success: true };
}

/**
 * Short Term: Monthly, Trial, Extended Trial, 2nd Trial, M-M, 10-Day, or < 3 months.
 * Long Term: 1 year (12 months), 6 months, 4 months, 3 months, etc.
 */
function isShortTerm(term) {
  const termString = String(term || '').toUpperCase().trim();
  if (!termString) return false;

  if (
    termString === 'M-M' ||
    termString === 'MM' ||
    termString.includes('MONTHLY') ||
    termString.includes('MONTH-MONTH') ||
    termString.includes('TRIAL') ||
    termString.includes('10-DAY') ||
    termString.includes('10 DAY')
  ) {
    return true;
  }

  if (termString.includes('YEAR')) return false;

  const numericMatch = termString.match(/(\d+)/);
  if (numericMatch) {
    const months = Number(numericMatch[1]);
    if (months > 0 && months < 3) return true;
    if (months >= 3) return false;
  }

  return false;
}

function isLongTerm(term) {
  return !isShortTerm(term);
}

function isMonthlyTerm(term) {
  const s = String(term || '').toUpperCase().trim();
  if (!s) return false;
  return (
    s === 'M-M' ||
    s === 'MM' ||
    s.includes('MONTHLY') ||
    s.includes('MONTH-MONTH') ||
    /^1\s*MONTHS?$/.test(s)
  );
}

function isEcoProperty(buildingOrTab) {
  const normalized = String(buildingOrTab || '')
    .toLowerCase()
    .replace(/[()]/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
  return normalized.includes('eco') || normalized === 'eco888';
}

function computeCalculatedRent(tenant, fallbackTab = '') {
  const rawAmount = normalizeMainSheetRent(tenant && tenant.rent);
  const amount = Number.isFinite(Number(rawAmount)) && rawAmount !== '' ? Number(rawAmount) : 0;
  const propertyLabel = (tenant && tenant.building) || fallbackTab || activeTab || '';
  const termLabel = (tenant && tenant.termLabel) || '';

  if (isEcoProperty(propertyLabel) && isMonthlyTerm(termLabel)) {
    if (amount === 0 || amount === 6888) return 6590;
    if (amount === 6590) return 6590;
  }

  const lastDigit = Math.abs(Math.round(amount)) % 10;
  if (lastDigit === 7) return amount - 298;
  if (lastDigit === 9) return amount;

  return amount;
}

function formatTermLabel(term) {
  const raw = String(term || '').trim();
  const termString = raw.toUpperCase();
  if (!termString) return '1 year (12 months)';
  if (termString.includes('2ND') && termString.includes('TRIAL')) return '2nd Trial';
  if (termString.includes('EXT') && termString.includes('TRIAL')) return 'Extended Trial';
  if (termString.includes('TRIAL')) return 'Trial';
  if (termString === 'M-M' || termString === 'MM' || termString.includes('MONTHLY') || termString === 'MONTH-MONTH') {
    return 'Monthly (M-M)';
  }
  if (termString.includes('YEAR')) {
    if (termString.includes('24') || termString.includes('2')) return '2 years (24 months)';
    return '1 year (12 months)';
  }
  const numericMatch = termString.match(/(\d+)/);
  if (numericMatch) {
    const months = Number(numericMatch[1]);
    if (months === 3) return '3 months';
    if (months === 4) return '4 months';
    if (months === 6) return '6 months';
    if (months === 12) return '1 year (12 months)';
    if (months === 24) return '2 years (24 months)';
    return `${months} months`;
  }
  return raw;
}

function getDailyRateByProperty(buildingOrTab) {
  const normalized = String(buildingOrTab || '')
    .toLowerCase()
    .replace(/[()]/g, ' ')
    .replace(/&/g, 'and')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();

  if (!normalized) return 700;
  if (normalized.includes('bed and bath') || normalized === 'b and b' || normalized === 'bnb') return 300;
  if (normalized.includes('homey')) return 500;
  if (normalized.includes('adi 168') || normalized === 'adi168') return 500;
  if (
    normalized.includes('dream') ||
    normalized.includes('penthauz') ||
    normalized.includes('penthouse') ||
    normalized.includes('eco') ||
    normalized.includes('888') ||
    normalized.includes('green') ||
    normalized.includes('kalayaan') ||
    normalized.includes('pleasant')
  ) {
    return 700;
  }

  return 700;
}

function hydrateTenantRecord(source, tabName) {
  const sourceEndDate = normalizeMainSheetDate(source.endDate || source.endContract);
  const sourceStartDate = normalizeMainSheetDate(source.startDate);
  const sourceComputedDueDate = normalizeMainSheetDate(source.computedDueDate || source.computed20th);
  const rawTerm = String(source.term || source.termLabel || '').trim();
  const hasExplicitTerm = rawTerm !== '';
  const effectiveTerm = hasExplicitTerm ? rawTerm : '12';
  const preservedTermLabel = formatTermLabel(effectiveTerm);

  const fixLegacyYear = isoStr => {
    if (!isoStr) return '';
    const m = isoStr.match(/^(191\d)-(\d{2}-\d{2})$/);
    if (m) {
      const correctedYear = Number(m[1]) + 109;
      return `${correctedYear}-${m[2]}`;
    }
    return isoStr;
  };

  const cleanEndDate = fixLegacyYear(sourceEndDate);
  const cleanStartDate = fixLegacyYear(sourceStartDate);
  const cleanComputedDue = fixLegacyYear(sourceComputedDueDate);

  let dates;
  if (cleanComputedDue) {
    dates = {
      computed20th: cleanComputedDue,
      ...deriveFromComputed20th(cleanComputedDue, effectiveTerm)
    };
  } else if (cleanEndDate) {
    dates = deriveFromContractEndDate(cleanEndDate, effectiveTerm);
  } else if (cleanStartDate) {
    dates = parseDate(cleanStartDate, effectiveTerm);
  } else {
    dates = {
      computed20th: '',
      sixthDay: '_______________',
      folderMonth: 'Unknown Month'
    };
  }

  return {
    tenantName: String(source.tenantName || '').trim(),
    unit: String(source.unit || '').trim(),
    building: String(source.building || tabName).trim(),
    rent: normalizeMainSheetRent(source.rent) || 0,
    computed20th: dates.computed20th,
    sixthDay: dates.sixthDay,
    folderMonth: dates.folderMonth,
    termLabel: preservedTermLabel,
    hasExplicitTerm,
    status: source.status ? String(source.status).trim() : 'Pending',
    pdfUrl: source.pdfUrl ? String(source.pdfUrl).trim() : '',
    reason: source.reason ? String(source.reason).trim() : '—',
    dateSent: normalizeMainSheetDate(source.dateSent),
    startDate: cleanStartDate,
    endContract: cleanEndDate
  };
}

async function initDatabase(preserveTab = false, forceRemote = true) {
  const loadingOverlay = document.getElementById('loadingOverlay');
  loadingOverlay.classList.remove('hidden');
  loadingOverlay.classList.add('flex');
  document.getElementById('loadingText').textContent = "Fetching Data Directly from Google Sheets...";

  try {
    const nextData = {};
    const failedTabs = [];

    for (let i = 0; i < MAIN_SPREADSHEET_TABS.length; i++) {
      const tabName = MAIN_SPREADSHEET_TABS[i];
      document.getElementById('loadingText').textContent = `Fetching ${tabName} (${i + 1}/${MAIN_SPREADSHEET_TABS.length})...`;
      try {
        const rows = await loadMainSpreadsheetTab(tabName);
        const extracted = extractMainTenantRows(rows);
        nextData[tabName] = extracted.map(row => hydrateTenantRecord(row, tabName));
      } catch (err) {
        console.warn(`Tab "${tabName}" skipped:`, err.message);
        failedTabs.push(`${tabName}: ${err.message}`);
        nextData[tabName] = [];
      }
    }

    const totalLoaded = Object.values(nextData).reduce((sum, list) => sum + list.length, 0);

    if (totalLoaded === 0 && !forceRemote) {
      const cached = localStorage.getItem(LOCAL_STORAGE_KEY);
      if (cached) appData = JSON.parse(cached);
    } else if (totalLoaded === 0 && failedTabs.length === MAIN_SPREADSHEET_TABS.length) {
      const cached = localStorage.getItem(LOCAL_STORAGE_KEY);
      if (cached) {
        appData = JSON.parse(cached);
        alert("Could not reach the Google Sheet (showing cached copy):\n" + failedTabs[0]);
      } else {
        throw new Error(failedTabs[0] || "Could not read any tabs from the Google Sheet. Ensure link sharing is set to 'Anyone with the link can view'.");
      }
    } else {
      appData = nextData;
      persistLocalCache();
    }

    const tabNames = Object.keys(appData);
    if (tabNames.length > 0) {
      if (!preserveTab || !activeTab || !tabNames.includes(activeTab)) {
        activeTab = tabNames.find(name => appData[name].length > 0) || tabNames[0];
      }
      renderTabs();
      renderSpreadsheet();
    } else {
      document.getElementById('spreadsheetBody').innerHTML = `<tr><td colspan="9" class="px-4 py-12 text-center text-slate-400">No tabs found in the Google Sheet.</td></tr>`;
    }
  } catch (error) {
    console.error(error);
    alert("Failed to fetch Google Sheet directly: " + error.message);
  } finally {
    loadingOverlay.classList.add('hidden');
    loadingOverlay.classList.remove('flex');
  }
}

function findTenantByNameAndUnit(tabName, tenantName, unit) {
  if (!tabName || !appData[tabName]) return null;
  if (unit !== undefined && unit !== null && String(unit).trim() !== '') {
    const exact = appData[tabName].find(t => t.tenantName === tenantName && String(t.unit) === String(unit));
    if (exact) return exact;
  }
  return appData[tabName].find(t => t.tenantName === tenantName) || null;
}

function normalizeSyncValue(value) {
  return String(value == null ? '' : value).trim().toLowerCase().replace(/\s+/g, ' ');
}

function normalizeSyncHeader(value) {
  return String(value == null ? '' : value).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function parseMainSheetValue(value, formattedValue) {
  if (value == null || value === '') return formattedValue || '';
  if (typeof value !== 'string') return value;
  const dateMatch = value.match(/^Date\((\d{4}),(\d{1,2}),(\d{1,2})(?:,.*)?\)$/);
  if (dateMatch) {
    return formatDate(new Date(Number(dateMatch[1]), Number(dateMatch[2]), Number(dateMatch[3])));
  }
  return value;
}

function normalizeMainSheetDate(value) {
  const dateValue = toDateInputValue(value);
  return dateValue && !isNaN(new Date(dateValue).getTime()) ? dateValue : '';
}

function deriveFromContractEndDate(endDate, term) {
  const noticeDate = new Date(endDate);
  noticeDate.setDate(noticeDate.getDate() - (isShortTerm(term) ? 10 : 20));
  const computed20th = formatDate(noticeDate);
  return { computed20th, ...deriveFromComputed20th(computed20th, term) };
}

function normalizeMainSheetRent(value) {
  if (typeof value === 'number') return value;
  const normalized = String(value == null ? '' : value).replace(/[₱,$\s]/g, '');
  if (!normalized) return '';
  const amount = Number(normalized);
  return Number.isFinite(amount) ? amount : '';
}

function loadMainSpreadsheetTab(tabName) {
  return new Promise((resolve, reject) => {
    const callbackName = `mainSheetCallback_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const script = document.createElement('script');
    const timeout = setTimeout(() => finish(new Error(`Timed out loading the "${tabName}" tab.`)), 20000);
    let settled = false;

    function finish(error, result) {
      if (settled) return;
      settled = true;
      clearTimeout(timeout);
      script.remove();
      delete window[callbackName];
      if (error) reject(error);
      else resolve(result);
    }

    window[callbackName] = response => {
      if (response.status === 'error') {
        finish(new Error(response.errors?.[0]?.detailed_message || response.errors?.[0]?.message || `Could not read the "${tabName}" tab.`));
        return;
      }
      if (!response.table || !Array.isArray(response.table.rows)) {
        finish(new Error(`The "${tabName}" tab returned no readable data.`));
        return;
      }

      const colLabels = (response.table.cols || []).map(col => (col && col.label ? col.label.trim() : ''));
      const dataRows = response.table.rows.map(row =>
        (row.c || []).map(cell => parseMainSheetValue(cell && cell.v, cell && cell.f))
      );

      if (colLabels.some(Boolean)) {
        finish(null, [colLabels, ...dataRows]);
      } else {
        finish(null, dataRows);
      }
    };
    script.onerror = () => finish(new Error(`Could not access the "${tabName}" tab. Check that the source spreadsheet is shared for viewing.`));

    const endpoint = `https://docs.google.com/spreadsheets/d/${MAIN_SPREADSHEET_ID}/gviz/tq`;
    const params = new URLSearchParams({
      tqx: `out:json;responseHandler:${callbackName}`,
      sheet: tabName,
      headers: '1'
    });
    script.src = `${endpoint}?${params.toString()}`;
    document.head.appendChild(script);
  });
}

function extractMainTenantRows(rows) {
  const headerAliases = {
    tenantName: ['tenant name', 'name', 'tenant'],
    unit: ['unit', 'unit no', 'unit number'],
    building: ['building', 'property', 'property name'],
    computedDueDate: ['computed due date', 'computed due', 'computed notice deadline', 'notice deadline', 'computed 20th', 'computed20th'],
    rent: ['monthly rent', 'rent', 'rate', 'monthly rate'],
    status: ['status'],
    dateSent: ['date sent'],
    pdfUrl: ['pdf link', 'pdf url'],
    term: ['term', 'lease term', 'contract term', 'contract'],
    reason: ['reason'],
    startDate: ['start date', 'start of contract', 'contract start', 'lease start', 'renewal/movein', 'renewal', 'movein', 'move in date'],
    endDate: ['end of contract', 'end contract', 'end contract date', 'contract end', 'contract end date', 'lease end', 'lease end date', 'end date', 'expiration date', 'contract expiration', 'lease expiration', 'expiry date', 'contract expiry', 'contract expiration date']
  };
  const aliasSets = Object.fromEntries(
    Object.entries(headerAliases).map(([field, aliases]) => [field, new Set(aliases.map(normalizeSyncHeader))])
  );

  const headerRowIndex = rows.findIndex(row => {
    const headers = row.map(normalizeSyncHeader);
    return headers.some(header => aliasSets.tenantName.has(header)) &&
      headers.some(header => aliasSets.unit.has(header));
  });
  if (headerRowIndex === -1) {
    throw new Error('Could not find a header row containing both Tenant Name and Unit.');
  }

  const headers = rows[headerRowIndex].map(normalizeSyncHeader);
  const columnIndexes = {};
  Object.entries(aliasSets).forEach(([field, aliases]) => {
    columnIndexes[field] = headers.findIndex(header => aliases.has(header));
  });

  const positionalDefaults = {
    tenantName: 0,
    unit: 1,
    building: 2,
    computedDueDate: 3,
    rent: 4,
    status: 5,
    dateSent: 6,
    pdfUrl: 7,
    term: 8,
    reason: 9,
    startDate: 10
  };
  Object.entries(positionalDefaults).forEach(([field, defaultIdx]) => {
    if (columnIndexes[field] === -1 && headers.length > defaultIdx) {
      columnIndexes[field] = defaultIdx;
    }
  });

  if (columnIndexes.endDate === -1) {
    columnIndexes.endDate = headers.findIndex(header =>
      header.includes('contract') && (header.includes('end') || header.includes('expir') || header.includes('expiry'))
    );
  }
  if (columnIndexes.endDate === -1 && headers.length > 11) {
    columnIndexes.endDate = headers.length - 1;
  }

  return rows.slice(headerRowIndex + 1).map(row => {
    const record = {};
    Object.entries(columnIndexes).forEach(([field, index]) => {
      if (index >= 0) record[field] = row[index] == null ? '' : row[index];
    });
    return record;
  }).filter(record => {
    const name = normalizeSyncValue(record.tenantName);
    const unit = normalizeSyncValue(record.unit);
    if (!name || name === 'tenant name' || name === 'vacant' || name === '-') return false;
    return Boolean(unit && unit !== 'unit');
  });
}

function countByKey(records, keyFunction) {
  const counts = new Map();
  records.forEach(record => {
    const key = keyFunction(record);
    counts.set(key, (counts.get(key) || 0) + 1);
  });
  return counts;
}

function mainTenantIdentity(tenant) {
  return `${normalizeSyncValue(tenant.tenantName)}\u0000${normalizeSyncValue(tenant.unit)}`;
}

async function syncMainSpreadsheet(activeTabOnly = false, selectedOnly = false) {
  const buttonId = selectedOnly ? 'syncSelectedTenantsBtn' : activeTabOnly ? 'syncActiveTabBtn' : 'syncMainSpreadsheetBtn';
  const button = document.getElementById(buttonId);
  const loadingOverlay = document.getElementById('loadingOverlay');
  const summary = { imported: 0, updated: 0, unchanged: 0, unmatched: 0, ambiguous: 0, endContractFetched: 0, endContractUpdated: 0, failedTabs: [] };

  if (selectedOnly && (!activeTab || !appData[activeTab] || !appData[activeTab].some(tenant => selectedTenants.has(tenant)))) {
    return alert('Select at least one tenant in the active tab to synchronize.');
  }

  if (button) button.disabled = true;
  loadingOverlay.classList.remove('hidden');
  loadingOverlay.classList.add('flex');

  try {
    const tabNames = activeTabOnly || selectedOnly
      ? (activeTab ? [activeTab] : [])
      : Array.from(new Set([...Object.keys(appData), ...MAIN_SPREADSHEET_TABS]));

    if (tabNames.length === 0) {
      throw new Error(activeTabOnly ? 'Select a tab before synchronizing.' : 'There are no property tabs to synchronize.');
    }

    for (let tabIndex = 0; tabIndex < tabNames.length; tabIndex++) {
      const tabName = tabNames[tabIndex];
      document.getElementById('loadingText').textContent = `Refreshing ${tabName} (${tabIndex + 1}/${tabNames.length})...`;
      let sourceRecords;
      try {
        sourceRecords = extractMainTenantRows(await loadMainSpreadsheetTab(tabName));
      } catch (error) {
        summary.failedTabs.push(`${tabName}: ${error.message}`);
        continue;
      }

      if (!Array.isArray(appData[tabName])) appData[tabName] = [];
      const localRecords = appData[tabName];
      const localByPair = new Map();
      localRecords.forEach(tenant => {
        const key = mainTenantIdentity(tenant);
        if (!localByPair.has(key)) localByPair.set(key, []);
        localByPair.get(key).push(tenant);
      });
      const sourcePairCounts = countByKey(sourceRecords, mainTenantIdentity);
      const localNameCounts = countByKey(localRecords, tenant => normalizeSyncValue(tenant.tenantName));
      const sourceNameCounts = countByKey(sourceRecords, tenant => normalizeSyncValue(tenant.tenantName));
      const localUnitCounts = countByKey(localRecords, tenant => normalizeSyncValue(tenant.unit));
      const sourceUnitCounts = countByKey(sourceRecords, tenant => normalizeSyncValue(tenant.unit));
      const processedTenants = new Set();

      for (const source of sourceRecords) {
        const pairKey = mainTenantIdentity(source);
        let target = null;
        const exactMatches = localByPair.get(pairKey) || [];
        if (sourcePairCounts.get(pairKey) === 1 && exactMatches.length === 1) {
          target = exactMatches[0];
        } else {
          const unitKey = normalizeSyncValue(source.unit);
          const unitMatches = localRecords.filter(tenant => normalizeSyncValue(tenant.unit) === unitKey);
          if (sourceUnitCounts.get(unitKey) === 1 && localUnitCounts.get(unitKey) === 1 && unitMatches.length === 1) {
            target = unitMatches[0];
          } else {
            const nameKey = normalizeSyncValue(source.tenantName);
            const nameMatches = localRecords.filter(tenant => normalizeSyncValue(tenant.tenantName) === nameKey);
            if (sourceNameCounts.get(nameKey) === 1 && localNameCounts.get(nameKey) === 1 && nameMatches.length === 1) {
              target = nameMatches[0];
            }
          }
        }

        if (selectedOnly && (!target || !selectedTenants.has(target))) continue;

        if (!target || processedTenants.has(target)) {
          const canImportMissingTenant = !target &&
            !selectedOnly &&
            sourcePairCounts.get(pairKey) === 1 &&
            exactMatches.length === 0;
          if (canImportMissingTenant) {
            const importedTenant = hydrateTenantRecord(source, tabName);
            localRecords.push(importedTenant);
            processedTenants.add(importedTenant);
            await syncToDatabase({ action: 'save', tabName, tenant: importedTenant });
            summary.imported++;
            if (importedTenant.endContract) {
              summary.endContractFetched++;
              summary.endContractUpdated++;
            }
            continue;
          }
          if (target || exactMatches.length > 1 || sourcePairCounts.get(pairKey) > 1) summary.ambiguous++;
          else summary.unmatched++;
          continue;
        }
        processedTenants.add(target);

        const hydrated = hydrateTenantRecord(source, tabName);
        const updated = { ...target };
        const setIfPresent = (field, value) => {
          if (value !== '' && value != null) updated[field] = value;
        };
        setIfPresent('tenantName', hydrated.tenantName);
        setIfPresent('unit', hydrated.unit);
        setIfPresent('building', hydrated.building);
        setIfPresent('rent', hydrated.rent);
        if (source.reason && String(source.reason).trim() !== '' && String(source.reason).trim() !== '—') {
          updated.reason = hydrated.reason;
        }
        if (source.status && String(source.status).trim() !== '') {
          updated.status = hydrated.status;
        }
        if (source.pdfUrl && String(source.pdfUrl).trim() !== '') {
          updated.pdfUrl = hydrated.pdfUrl;
        }
        if (hydrated.dateSent) updated.dateSent = hydrated.dateSent;
        if (hydrated.startDate) updated.startDate = hydrated.startDate;
        if (hydrated.endContract) {
          summary.endContractFetched++;
          if (hydrated.endContract !== normalizeMainSheetDate(target.endContract)) {
            summary.endContractUpdated++;
          }
          updated.endContract = hydrated.endContract;
        }
        if (hydrated.hasExplicitTerm && hydrated.termLabel && !target.hasExplicitTerm) {
          updated.termLabel = hydrated.termLabel;
          updated.hasExplicitTerm = true;
        }
        if (hydrated.computed20th) {
          updated.computed20th = hydrated.computed20th;
          updated.sixthDay = hydrated.sixthDay;
          updated.folderMonth = hydrated.folderMonth;
        }

        const changed = ['tenantName', 'unit', 'building', 'rent', 'reason', 'status', 'dateSent', 'pdfUrl', 'startDate', 'endContract', 'termLabel', 'computed20th']
          .some(field => normalizeSyncValue(updated[field]) !== normalizeSyncValue(target[field]));
        if (!changed) {
          summary.unchanged++;
          continue;
        }

        Object.assign(target, updated);
        await syncToDatabase({ action: 'update', tabName, tenant: target });
        summary.updated++;
      }

      if (selectedOnly) {
        summary.unmatched += localRecords.filter(tenant =>
          selectedTenants.has(tenant) && !processedTenants.has(tenant)
        ).length;
      }
    }

    persistLocalCache();
    renderTabs();
    renderSpreadsheet();
    const details = [
      `Imported new tenants: ${summary.imported}`,
      `Updated: ${summary.updated}`,
      `End of Contract dates fetched: ${summary.endContractFetched}`,
      `End of Contract dates updated: ${summary.endContractUpdated}`,
      `Already current: ${summary.unchanged}`,
      `No match: ${summary.unmatched}`,
      `Ambiguous matches skipped: ${summary.ambiguous}`
    ];
    if (summary.failedTabs.length) details.push(`Tabs not synchronized:\n${summary.failedTabs.join('\n')}`);
    alert(`Google Sheet refresh complete.\n\n${details.join('\n')}`);
  } catch (error) {
    console.error(error);
    renderTabs();
    renderSpreadsheet();
    alert(`Google Sheet refresh stopped: ${error.message}`);
  } finally {
    if (button) button.disabled = false;
    loadingOverlay.classList.add('hidden');
    loadingOverlay.classList.remove('flex');
  }
}

document.getElementById('syncMainSpreadsheetBtn')?.addEventListener('click', () => syncMainSpreadsheet(false));
document.getElementById('syncActiveTabBtn')?.addEventListener('click', () => syncMainSpreadsheet(true));
document.getElementById('syncSelectedTenantsBtn')?.addEventListener('click', () => syncMainSpreadsheet(true, true));

function formatPeso(amount) {
  return '₱' + (Number(amount) || 0).toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatDate(dateObj) {
  const yyyy = dateObj.getFullYear();
  const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
  const dd = String(dateObj.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

function formatDisplayDate(value) {
  const isoDate = toDateInputValue(value);
  if (!isoDate) return '';
  const [year, month, day] = isoDate.split('-');
  return `${Number(month)}/${Number(day)}/${year}`;
}

function formatLongDate(value) {
  const isoDate = toDateInputValue(value);
  if (!isoDate) return value || '';
  const [year, month, day] = isoDate.split('-').map(Number);
  const d = new Date(year, month - 1, day);
  return isNaN(d) ? String(value) : d.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
}

function toDateInputValue(val) {
  if (!val) return '';
  if (val instanceof Date) return isNaN(val) ? '' : formatDate(val);
  if (typeof val === 'number') {
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return isNaN(d) ? '' : formatDate(d);
  }
  const s = String(val).trim();
  if (!s) return '';
  if (s.includes('T')) return toDateInputValue(s.split('T')[0]);

  const isoDateMatch = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (isoDateMatch) {
    const year = Number(isoDateMatch[1]);
    const month = Number(isoDateMatch[2]);
    const day = Number(isoDateMatch[3]);
    const date = new Date(year, month - 1, day);
    return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day
      ? formatDate(date)
      : '';
  }

  const usDateMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (usDateMatch) {
    const month = Number(usDateMatch[1]);
    const day = Number(usDateMatch[2]);
    const year = Number(usDateMatch[3]);
    const date = new Date(year, month - 1, day);
    return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day
      ? formatDate(date)
      : '';
  }

  const dmyTextMatch = s.match(/^(\d{1,2})[-\s]([A-Za-z]{3,9})[-\s](\d{2,4})$/);
  if (dmyTextMatch) {
    const months = {
      jan: 0, january: 0, feb: 1, february: 1, mar: 2, march: 2, apr: 3, april: 3,
      may: 4, jun: 5, june: 5, jul: 6, july: 6, aug: 7, august: 7,
      sep: 8, sept: 8, september: 8, oct: 9, october: 9, nov: 10, november: 10, dec: 11, december: 11
    };
    const day = Number(dmyTextMatch[1]);
    const monthIdx = months[dmyTextMatch[2].toLowerCase()];
    let year = Number(dmyTextMatch[3]);
    if (year < 100) year += 2000;
    if (monthIdx !== undefined) {
      const date = new Date(year, monthIdx, day);
      if (date.getFullYear() === year && date.getMonth() === monthIdx && date.getDate() === day) {
        return formatDate(date);
      }
    }
  }

  const d = new Date(s);
  return isNaN(d) ? '' : formatDate(d);
}

function deriveFromComputed20th(dateStr, term) {
  const out = { sixthDay: '_______________', folderMonth: 'Unknown Month' };
  if (!dateStr) return out;
  const d = new Date(dateStr);
  if (isNaN(d)) return out;
  const penalty = new Date(d);
  penalty.setDate(penalty.getDate() + 6);
  out.sixthDay = penalty.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  out.folderMonth = d.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  return out;
}

function parseDate(startContract, term) {
  let computed20th = '';
  let sixthDay = '';
  const termLabel = formatTermLabel(term || '12');
  let termMonths = 12;
  let folderMonth = 'Unknown Month';

  let startD;
  if (typeof startContract === 'number') {
    startD = new Date(Math.round((startContract - 25569) * 86400 * 1000));
  } else if (startContract instanceof Date) {
    startD = startContract;
  } else if (typeof startContract === 'string') {
    const iso = toDateInputValue(startContract);
    startD = iso ? new Date(iso) : new Date(startContract.trim());
  } else {
    startD = new Date(startContract);
  }

  if (!isNaN(startD.getTime())) {
    const shortTerm = isShortTerm(term);

    if (shortTerm) {
      termMonths = 1;
    } else {
      const numericMatch = String(term || '').match(/(\d+)/);
      termMonths = numericMatch ? Number(numericMatch[1]) : 12;
    }

    let actualEndD = new Date(startD);
    actualEndD.setMonth(actualEndD.getMonth() + termMonths);

    let noticeDate = new Date(actualEndD);
    if (shortTerm) {
      noticeDate.setDate(noticeDate.getDate() - 10);
    } else {
      noticeDate.setDate(noticeDate.getDate() - 20);
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    while (noticeDate.getTime() < today.getTime() && noticeDate.getTime() < actualEndD.getTime()) {
      noticeDate.setMonth(noticeDate.getMonth() + 1);
    }

    if (noticeDate.getTime() > actualEndD.getTime()) {
      noticeDate = new Date(actualEndD);
    }

    startD = noticeDate;

    computed20th = formatDate(startD);
    folderMonth = startD.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });

    let penaltyD = new Date(startD);
    penaltyD.setDate(penaltyD.getDate() + 6);
    sixthDay = penaltyD.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  } else {
    computed20th = startContract ? String(startContract) : 'No Date';
    sixthDay = '_______________';
    folderMonth = 'Unknown Month';
  }

  return { computed20th, sixthDay, termMonths, termLabel, folderMonth };
}

function renderTabs() {
  const tabsContainer = document.getElementById('tabsContainer');
  tabsContainer.innerHTML = '';
  const tabNames = Object.keys(appData);

  if (tabNames.length === 0) {
    tabsContainer.innerHTML = `<span class="text-sm text-slate-400 italic">No tabs found.</span>`;
    return;
  }

  tabNames.forEach(tabName => {
    const isActive = tabName === activeTab;
    const tabBtn = document.createElement('button');
    tabBtn.className = `px-4 py-2 rounded-t-lg font-bold text-sm transition-colors border-b-2 whitespace-nowrap ${isActive ? 'bg-white border-blue-600 text-blue-700 shadow-sm' : 'bg-slate-200 border-transparent text-slate-600 hover:bg-slate-300'}`;
    tabBtn.textContent = tabName;

    tabBtn.addEventListener('click', () => {
      selectedTenants.clear();
      activeTab = tabName;
      document.getElementById('searchInput').value = '';
      document.getElementById('statusFilter').value = 'ALL';
      renderTabs();
      renderSpreadsheet();
    });

    tabsContainer.appendChild(tabBtn);
  });
}

function populateMonthFilter() {
  const monthFilter = document.getElementById('monthFilter');
  monthFilter.innerHTML = '<option value="ALL">All Months</option>';

  if (!activeTab || !appData[activeTab]) return;

  const tenants = appData[activeTab];
  const uniqueMonths = new Set();

  tenants.forEach(t => {
    if (t.computed20th && t.computed20th !== 'No Date' && t.computed20th !== 'Unknown Month') {
      try {
        const d = new Date(t.computed20th);
        if (!isNaN(d)) {
          uniqueMonths.add(d.toLocaleDateString('en-US', { month: 'long', year: 'numeric' }));
        }
      } catch (e) {}
    }
  });

  Array.from(uniqueMonths).sort((a, b) => new Date(a) - new Date(b)).forEach(monthStr => {
    const option = document.createElement('option');
    option.value = monthStr;
    option.textContent = monthStr;
    monthFilter.appendChild(option);
  });

  monthFilter.value = 'ALL';
}

function renderSpreadsheet() {
  const spreadsheetBody = document.getElementById('spreadsheetBody');
  spreadsheetBody.innerHTML = '';

  let pending = 0, generated = 0, sent = 0;

  if (!activeTab || !appData[activeTab] || appData[activeTab].length === 0) {
    spreadsheetBody.innerHTML = `<tr><td colspan="9" class="px-4 py-12 text-center text-slate-400">No data in this tab.</td></tr>`;
    document.getElementById('countPending').textContent = `0 Pending`;
    document.getElementById('countGenerated').textContent = `0 Generated`;
    document.getElementById('countSent').textContent = `0 Sent`;
    updateSelectionControls();
    return;
  }

  if (document.getElementById('monthFilter').value === 'ALL') populateMonthFilter();

  const tenants = appData[activeTab];
  const searchTerm = document.getElementById('searchInput').value.toLowerCase();
  const statusValue = document.getElementById('statusFilter').value;
  const monthValue = document.getElementById('monthFilter').value;

  tenants.forEach(t => {
    if (t.status === 'Pending') pending++;
    else if (t.status === 'Generated') generated++;
    else if (t.status === 'Sent') sent++;
  });

  const filteredTenants = tenants.filter(t => {
    const matchesSearch = (t.tenantName || '').toLowerCase().includes(searchTerm) || (t.unit || '').toLowerCase().includes(searchTerm);
    const matchesStatus = (statusValue === 'ALL') || (t.status === statusValue);

    let matchesMonth = true;
    if (monthValue !== 'ALL') {
      try {
        const d = new Date(t.computed20th);
        matchesMonth = !isNaN(d) && (d.toLocaleDateString('en-US', { month: 'long', year: 'numeric' }) === monthValue);
      } catch (e) { matchesMonth = false; }
    }

    return matchesSearch && matchesStatus && matchesMonth;
  });

  if (filteredTenants.length === 0) {
    spreadsheetBody.innerHTML = `<tr><td colspan="9" class="px-4 py-12 text-center text-slate-400">No tenants found matching your search.</td></tr>`;
  } else {
    filteredTenants.forEach((tenant) => {
      let badgeClass = 'bg-yellow-100 text-yellow-800';
      if (tenant.status === 'Generated') badgeClass = 'bg-blue-100 text-blue-800';
      if (tenant.status === 'Sent') badgeClass = 'bg-green-100 text-green-800';

      const hasPdf = tenant.pdfUrl && String(tenant.pdfUrl).trim() !== '';
      let pdfCell = hasPdf ? `<a href="${tenant.pdfUrl}" target="_blank" rel="noopener" class="text-blue-600 hover:underline font-bold text-xs flex justify-center items-center gap-1">View PDF</a>` : `<span class="text-slate-300 text-xs">—</span>`;

      const isDisabled = tenant.status !== 'Sent';
      const disabledAttr = isDisabled ? 'disabled' : '';

      const safeName = String(tenant.tenantName || '').replace(/'/g, "\\'");
      const safeUnit = String(tenant.unit || '').replace(/'/g, "\\'");

      const dateSentHtml = `<input type="text" inputmode="numeric" placeholder="M/D/YYYY" class="table-date-input" value="${formatDisplayDate(tenant.dateSent)}" ${disabledAttr} onchange="updateDateSent('${safeName}', this.value, '${safeUnit}')" aria-label="Date sent (M/D/YYYY)">`;

      const computed20thValue = formatDisplayDate(tenant.computed20th);
      const computed20thHtml = `<input type="text" inputmode="numeric" placeholder="M/D/YYYY" class="table-date-input" value="${computed20thValue}" onchange="updateComputed20th('${safeName}', this.value, '${safeUnit}')" title="Edit Notice Deadline (M/D/YYYY)" aria-label="Notice deadline (M/D/YYYY)">`;
      const endContractDisplay = formatDisplayDate(tenant.endContract) || '—';

      let actionBtn = '';
      if (tenant.status === 'Pending') {
        actionBtn = `
          <div class="flex justify-end items-center gap-1">
            <button onclick="generateSingle('${safeName}', '${safeUnit}')" class="bg-slate-800 text-white px-2 py-1 rounded text-xs font-bold hover:bg-slate-900 transition-colors" title="Generate PDF">Generate</button>
            <button onclick="openEditModal('${safeName}', '${safeUnit}')" class="bg-blue-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-blue-700 transition-colors" title="Edit tenant">Edit</button>
            <button onclick="deleteTenant('${safeName}', '${safeUnit}')" class="bg-red-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-red-700 transition-colors" title="Delete tenant">Del</button>
          </div>`;
      } else if (tenant.status === 'Generated') {
        actionBtn = `
          <div class="flex justify-end items-center gap-1 flex-wrap">
            <button onclick="regenerateSingle('${safeName}', '${safeUnit}')" class="bg-slate-500 text-white px-2 py-1 rounded text-xs font-bold hover:bg-slate-600 transition-colors" title="Re-generate">Regenerate</button>
            <button onclick="markAsSent('${safeName}', '${safeUnit}')" class="bg-green-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-green-700 transition-colors">Sent</button>
            <button onclick="openEditModal('${safeName}', '${safeUnit}')" class="bg-blue-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-blue-700 transition-colors" title="Edit tenant">Edit</button>
            <button onclick="deleteTenant('${safeName}', '${safeUnit}')" class="bg-red-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-red-700 transition-colors" title="Delete tenant">Del</button>
          </div>`;
      } else {
        actionBtn = `
          <div class="flex justify-end items-center gap-1">
            <button onclick="openEditModal('${safeName}', '${safeUnit}')" class="bg-blue-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-blue-700 transition-colors" title="Edit tenant">Edit</button>
            <button onclick="deleteTenant('${safeName}', '${safeUnit}')" class="bg-red-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-red-700 transition-colors" title="Delete tenant">Del</button>
          </div>`;
      }

      const tr = document.createElement('tr');
      tr.className = "hover:bg-slate-50 transition-colors";

      tr.innerHTML = `
        <td class="px-4 py-3 font-medium text-slate-800">${tenant.tenantName}</td>
        <td class="px-4 py-3 text-slate-600">${tenant.unit}</td>
        <td class="px-4 py-2 text-slate-600 font-medium">${computed20thHtml}</td>
        <td class="px-4 py-3 text-slate-600">${endContractDisplay}</td>
        <td class="px-4 py-3"><span class="px-2 py-1 text-[10px] uppercase tracking-wide rounded-full font-bold ${badgeClass}">${tenant.status}</span></td>
        <td class="px-4 py-3 text-center">${pdfCell}</td>
        <td class="px-4 py-3 text-right">${actionBtn}</td>
        <td class="px-4 py-3 text-center">${dateSentHtml}</td>
      `;
      const selectionCell = document.createElement('td');
      selectionCell.className = 'px-4 py-3 text-center';
      const selectionCheckbox = document.createElement('input');
      selectionCheckbox.type = 'checkbox';
      selectionCheckbox.className = 'tenant-select-checkbox h-4 w-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500';
      selectionCheckbox.checked = selectedTenants.has(tenant);
      selectionCheckbox.setAttribute('aria-label', `Select ${tenant.tenantName}`);
      selectionCheckbox.tenantRecord = tenant;
      selectionCheckbox.addEventListener('change', () => {
        if (selectionCheckbox.checked) selectedTenants.add(tenant);
        else selectedTenants.delete(tenant);
        updateSelectionControls();
      });
      selectionCell.appendChild(selectionCheckbox);
      tr.prepend(selectionCell);
      spreadsheetBody.appendChild(tr);
    });
  }

  document.getElementById('countPending').textContent = `${pending} Pending`;
  document.getElementById('countGenerated').textContent = `${generated} Generated`;
  document.getElementById('countSent').textContent = `${sent} Sent`;
  updateSelectionControls();
}

function updateSelectionControls() {
  const selectedTenantCount = activeTab && appData[activeTab]
    ? appData[activeTab].filter(tenant => selectedTenants.has(tenant)).length
    : 0;
  const selectedCount = activeTab && appData[activeTab]
    ? appData[activeTab].filter(tenant => selectedTenants.has(tenant)).length
    : 0;
  const selectedPendingCount = activeTab && appData[activeTab]
    ? appData[activeTab].filter(tenant => selectedTenants.has(tenant) && tenant.status === 'Pending').length
    : 0;
  const generateSelectedButton = document.getElementById('batchGenerateSelectedBtn');
  if (generateSelectedButton) {
    document.getElementById('batchGenerateSelectedLabel').textContent = `Generate Selected (${selectedPendingCount})`;
    generateSelectedButton.disabled = selectedPendingCount === 0;
    generateSelectedButton.classList.toggle('opacity-50', selectedPendingCount === 0);
    generateSelectedButton.classList.toggle('cursor-not-allowed', selectedPendingCount === 0);
  }

  const downloadButton = document.getElementById('batchDownloadZipBtn');
  if (downloadButton) {
    document.getElementById('batchDownloadZipLabel').textContent = `Download Selected ZIP (${selectedCount})`;
    downloadButton.disabled = selectedCount === 0;
    downloadButton.classList.toggle('opacity-50', selectedCount === 0);
    downloadButton.classList.toggle('cursor-not-allowed', selectedCount === 0);
  }

  const syncSelectedButton = document.getElementById('syncSelectedTenantsBtn');
  if (syncSelectedButton) {
    document.getElementById('syncSelectedTenantsLabel').textContent = `Sync Selected (${selectedTenantCount})`;
    syncSelectedButton.disabled = selectedTenantCount === 0;
    syncSelectedButton.classList.toggle('opacity-50', selectedTenantCount === 0);
    syncSelectedButton.classList.toggle('cursor-not-allowed', selectedTenantCount === 0);
  }

  const selectAllCheckbox = document.getElementById('selectAllTenantsCheckbox');
  if (selectAllCheckbox) {
    const visibleCheckboxes = Array.from(document.querySelectorAll('#spreadsheetBody .tenant-select-checkbox'));
    const visibleSelectedCount = visibleCheckboxes.filter(checkbox => checkbox.checked).length;
    selectAllCheckbox.disabled = visibleCheckboxes.length === 0;
    selectAllCheckbox.checked = visibleCheckboxes.length > 0 && visibleSelectedCount === visibleCheckboxes.length;
    selectAllCheckbox.indeterminate = visibleSelectedCount > 0 && visibleSelectedCount < visibleCheckboxes.length;
  }
}

document.getElementById('selectAllTenantsCheckbox')?.addEventListener('change', event => {
  const isSelected = event.target.checked;
  document.querySelectorAll('#spreadsheetBody .tenant-select-checkbox').forEach(checkbox => {
    checkbox.checked = isSelected;
    const tenant = checkbox.tenantRecord;
    if (tenant) {
      if (isSelected) selectedTenants.add(tenant);
      else selectedTenants.delete(tenant);
    }
  });
  updateSelectionControls();
});

window.openAddModal = function() {
  if (!activeTab) return alert("Select a tab first.");
  isAddMode = true;
  currentEditingTenant = null;
  document.getElementById('editTenantName').value = '';
  document.getElementById('editTenantName').readOnly = false;
  document.getElementById('editUnit').value = '';
  document.getElementById('editBuilding').value = activeTab || "Eve's Residences";
  document.getElementById('editStartDate').value = '';
  document.getElementById('editEndContract').value = '';
  document.getElementById('editComputed20th').value = '';
  document.getElementById('editTerm').value = '12';
  document.getElementById('editRent').value = '6000';
  document.getElementById('editStatus').value = 'Pending';
  document.getElementById('editDateSent').value = '';
  document.getElementById('editReason').value = '';
  document.getElementById('editTenantModal').querySelector('h3').lastChild.textContent = ' Add Tenant';
  document.getElementById('editTenantModal').classList.remove('hidden');
  document.getElementById('editTenantModal').classList.add('flex');
};

window.openEditModal = function(tenantName, unit) {
  if (!activeTab) return;
  const tenant = findTenantByNameAndUnit(activeTab, tenantName, unit);
  if (!tenant) return alert("Tenant not found.");
  isAddMode = false;
  currentEditingTenant = tenant;

  document.getElementById('editTenantName').value = tenant.tenantName || '';
  document.getElementById('editTenantName').readOnly = false;
  document.getElementById('editUnit').value = tenant.unit || '';
  document.getElementById('editBuilding').value = tenant.building || '';

  document.getElementById('editStartDate').value = toDateInputValue(tenant.startDate);
  document.getElementById('editEndContract').value = toDateInputValue(tenant.endContract);
  document.getElementById('editComputed20th').value = toDateInputValue(tenant.computed20th);

  let termVal = '12';
  const tl = String(tenant.termLabel || '').toLowerCase();
  if (tl.includes('monthly') || tl.includes('m-m') || tl.includes('month-month')) termVal = 'M-M';
  else if (tl.includes('2nd') && tl.includes('trial')) termVal = '2nd Trial';
  else if (tl.includes('ext') && tl.includes('trial')) termVal = 'Ext. Trial';
  else if (tl.includes('trial')) termVal = 'Trial';
  else if (tl.includes('24') || tl.includes('2 years')) termVal = '24';
  else if (tl.includes('12') || tl.includes('1 year')) termVal = '12';
  else if (/\b6\b/.test(tl)) termVal = '6';
  else if (/\b4\b/.test(tl)) termVal = '4';
  else if (/\b3\b/.test(tl)) termVal = '3';
  document.getElementById('editTerm').value = termVal;

  document.getElementById('editRent').value = tenant.rent || 0;
  document.getElementById('editStatus').value = tenant.status || 'Pending';
  document.getElementById('editDateSent').value = toDateInputValue(tenant.dateSent);
  document.getElementById('editReason').value = tenant.reason && tenant.reason !== '—' ? tenant.reason : '';

  const modal = document.getElementById('editTenantModal');
  modal.querySelector('h3').lastChild.textContent = ' Edit Tenant';
  modal.classList.remove('hidden');
  modal.classList.add('flex');
};

window.closeEditModal = function() {
  const modal = document.getElementById('editTenantModal');
  modal.classList.add('hidden');
  modal.classList.remove('flex');
  currentEditingTenant = null;
  isAddMode = false;
};

window.saveTenantEdit = async function() {
  const name = document.getElementById('editTenantName').value.trim();
  if (!name) return alert("Tenant name is required.");
  if (!activeTab) return alert("No active tab.");

  const unit = document.getElementById('editUnit').value.trim();
  const building = document.getElementById('editBuilding').value.trim();
  const startDate = document.getElementById('editStartDate').value;
  const endContract = document.getElementById('editEndContract').value;
  const manualComputed20th = document.getElementById('editComputed20th').value;
  const term = document.getElementById('editTerm').value;
  const rent = Number(document.getElementById('editRent').value) || 0;
  const status = document.getElementById('editStatus').value;
  const dateSent = document.getElementById('editDateSent').value || '';
  const reason = document.getElementById('editReason').value.trim() || '—';

  const autoDates = parseDate(startDate, term);

  let finalComputed20th;
  let derived;
  if (manualComputed20th) {
    finalComputed20th = manualComputed20th;
    derived = deriveFromComputed20th(manualComputed20th, term);
  } else if (endContract && !startDate) {
    const endDerived = deriveFromContractEndDate(endContract, term);
    finalComputed20th = endDerived.computed20th;
    derived = { sixthDay: endDerived.sixthDay, folderMonth: endDerived.folderMonth };
  } else {
    finalComputed20th = autoDates.computed20th;
    derived = { sixthDay: autoDates.sixthDay, folderMonth: autoDates.folderMonth };
  }

  try {
    if (isAddMode) {
      if (appData[activeTab].some(t => t.tenantName === name && String(t.unit) === String(unit))) {
        throw new Error("A tenant with this name and unit already exists in this tab.");
      }
      const newTenant = {
        tenantName: name,
        unit, building, rent,
        computed20th: finalComputed20th,
        sixthDay: derived.sixthDay,
        folderMonth: derived.folderMonth,
        termLabel: formatTermLabel(term),
        hasExplicitTerm: true,
        status,
        pdfUrl: '',
        reason,
        dateSent,
        startDate,
        endContract
      };
      appData[activeTab].push(newTenant);
      await syncToDatabase({ action: 'save', tabName: activeTab, tenant: newTenant });
    } else {
      if (!currentEditingTenant) throw new Error("No tenant selected.");
      const originalName = currentEditingTenant.tenantName;
      const originalUnit = currentEditingTenant.unit;

      if ((name !== originalName || unit !== originalUnit) && appData[activeTab].some(t => t !== currentEditingTenant && t.tenantName === name && String(t.unit) === String(unit))) {
        throw new Error("Another tenant with this name and unit already exists.");
      }

      const updated = {
        ...currentEditingTenant,
        tenantName: name,
        unit, building, rent,
        computed20th: finalComputed20th,
        sixthDay: derived.sixthDay,
        folderMonth: derived.folderMonth,
        termLabel: formatTermLabel(term),
        hasExplicitTerm: true,
        status,
        reason,
        dateSent,
        startDate,
        endContract
      };

      const idx = appData[activeTab].indexOf(currentEditingTenant);
      if (idx !== -1) appData[activeTab][idx] = updated;
      await syncToDatabase({
        action: 'update',
        tabName: activeTab,
        tenant: updated,
        originalName,
        originalUnit
      });
    }

    closeEditModal();
    renderSpreadsheet();
  } catch (err) {
    console.error(err);
    alert("Error: " + err.message);
  }
};

window.deleteTenant = async function(tenantName, unit) {
  if (!activeTab) return;
  const tenant = findTenantByNameAndUnit(activeTab, tenantName, unit);
  if (!tenant) return alert("Tenant not found.");
  const idx = appData[activeTab].indexOf(tenant);

  const confirmDelete = confirm(`Remove tenant "${tenantName}" (Unit ${tenant.unit}) from the current view?`);
  if (!confirmDelete) return;

  selectedTenants.delete(tenant);
  appData[activeTab].splice(idx, 1);
  await syncToDatabase({ action: 'delete', tabName: activeTab, tenant });
  renderSpreadsheet();
};

document.getElementById('searchInput').addEventListener('input', renderSpreadsheet);
document.getElementById('statusFilter').addEventListener('change', renderSpreadsheet);
document.getElementById('monthFilter').addEventListener('change', renderSpreadsheet);

window.updateDateSent = async function(tenantName, newDateStr, unit) {
  const tenant = findTenantByNameAndUnit(activeTab, tenantName, unit);
  if (!tenant) return;

  const normalizedDate = newDateStr.trim() ? toDateInputValue(newDateStr) : '';
  if (newDateStr.trim() && !normalizedDate) {
    alert("Enter the date as M/D/YYYY, for example 7/22/2026.");
    renderSpreadsheet();
    return;
  }
  tenant.dateSent = normalizedDate;
  await syncToDatabase({ action: 'update', tabName: activeTab, tenant });
  renderSpreadsheet();
};

window.updateComputed20th = async function(tenantName, newDateStr, unit) {
  const tenant = findTenantByNameAndUnit(activeTab, tenantName, unit);
  if (!tenant) return;

  const normalizedDate = newDateStr.trim() ? toDateInputValue(newDateStr) : '';
  if (newDateStr.trim() && !normalizedDate) {
    alert("Enter the date as M/D/YYYY, for example 7/22/2026.");
    renderSpreadsheet();
    return;
  }
  if (!normalizedDate) {
    renderSpreadsheet();
    return;
  }

  const derived = deriveFromComputed20th(normalizedDate, tenant.termLabel);
  tenant.computed20th = normalizedDate;
  tenant.sixthDay = derived.sixthDay;
  tenant.folderMonth = derived.folderMonth;

  await syncToDatabase({ action: 'update', tabName: activeTab, tenant });
  renderSpreadsheet();
};

function applyContractDistinctions(rootEl, longTerm) {
  if (!rootEl) return;

  const fromWindow = longTerm ? 'last 10 days to last 5 days' : 'last 20 days to last 15 days';
  const toWindow = longTerm ? 'last 20 days to last 15 days' : 'last 10 days to last 5 days';
  const fromPenalty = longTerm ? 'last 4 days' : 'last 14 days';
  const toPenalty = longTerm ? 'last 14 days' : 'last 4 days';
  const fromPenaltyCap = longTerm ? 'Last 4 days' : 'Last 14 days';
  const toPenaltyCap = longTerm ? 'Last 14 days' : 'Last 4 days';
  const fromDays = longTerm ? '10' : '20';
  const toDays = longTerm ? '20' : '10';
  const fromOverstay = longTerm ? '1 day' : '7 days';
  const toOverstay = longTerm ? '7 days' : '1 day';

  const patterns = [
    [
      new RegExp(`Before the ${fromWindow.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')} via official Messenger`, 'gi'),
      `Before the ${toWindow} via official Messenger`
    ],
    [
      new RegExp(`From the ${fromWindow.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')} before contract ends`, 'gi'),
      `From the ${toWindow} before contract ends`
    ],
    [
      new RegExp(`Penalty Start date \\(${fromPenaltyCap}\\)`, 'gi'),
      `Penalty Start date (${toPenaltyCap})`
    ],
    [
      new RegExp(`${fromDays}-day notice period`, 'gi'),
      `${toDays}-day notice period`
    ],
    [
      new RegExp(`${fromDays}-DAY NOTICE`, 'g'),
      `${toDays}-DAY NOTICE`
    ],
    [
      new RegExp(fromWindow.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'),
      toWindow
    ],
    [
      new RegExp(fromPenaltyCap.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'),
      toPenaltyCap
    ],
    [
      new RegExp(fromPenalty.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'),
      toPenalty
    ],
    [
      new RegExp(`daily rate for ${fromOverstay}(?:\\s+only)?`, 'gi'),
      match => match.toLowerCase().includes('only') ? `daily rate for ${toOverstay} only` : `daily rate for ${toOverstay}`
    ]
  ];

  const walker = document.createTreeWalker(rootEl, NodeFilter.SHOW_TEXT, null);
  const textNodes = [];
  let node;
  while ((node = walker.nextNode())) {
    if (node.parentElement && (node.parentElement.id === 'outTermLabel' || node.parentElement.classList.contains('outTermLabel'))) continue;
    if (node.nodeValue && node.nodeValue.trim() !== '') {
      textNodes.push(node);
    }
  }

  textNodes.forEach(textNode => {
    let val = textNode.nodeValue;
    patterns.forEach(([regex, replacement]) => {
      val = val.replace(regex, replacement);
    });
    textNode.nodeValue = val;
  });
}

function applyDailyRateToTemplate(rootEl, formattedRate) {
  if (!rootEl) return;

  const rateEl = rootEl.querySelector('#outDailyRate');
  if (rateEl) rateEl.textContent = formattedRate;

  rootEl.querySelectorAll('.outDailyRate').forEach(el => {
    el.textContent = formattedRate;
  });

  const walker = document.createTreeWalker(rootEl, NodeFilter.SHOW_TEXT, null);
  const textNodes = [];
  let node;
  while ((node = walker.nextNode())) {
    if (node.parentElement && (node.parentElement.id === 'outTermLabel' || node.parentElement.classList.contains('outTermLabel'))) continue;
    if (node.nodeValue && node.nodeValue.trim() !== '') {
      textNodes.push(node);
    }
  }

  const dailyRateBlankRegex = /(daily rate of\s*)_{3,}/gi;
  textNodes.forEach(textNode => {
    if (dailyRateBlankRegex.test(textNode.nodeValue)) {
      textNode.nodeValue = textNode.nodeValue.replace(dailyRateBlankRegex, `$1${formattedRate}`);
    }
  });
}

async function generatePDFDoc(tenant) {
  const safeSetText = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.textContent = val;
  };

  const exactTermText = String(tenant.termLabel || '').trim();
  const longTerm = isLongTerm(exactTermText);
  const propertySource = tenant.building || activeTab || '';
  const dailyRateFormatted = formatPeso(getDailyRateByProperty(propertySource));
  const calculatedRent = computeCalculatedRent(tenant, activeTab);

  const noticeDaysStr = longTerm ? '20' : '10';
  const titleText = `${noticeDaysStr}-DAY NOTICE & ADVANCE AGREEMENT FOR RENEWAL OF CONTRACT`;
  const dayOrdinalStr = longTerm ? '20th' : '10th';
  const noticeWindowStr = longTerm ? 'last 20 days to last 15 days' : 'last 10 days to last 5 days';
  const nonRenewalNoticeStr = `Before the ${noticeWindowStr} via official Messenger`;
  const advancePaymentWindowStr = `From the ${noticeWindowStr} before contract ends`;
  const overstayDaysStr = longTerm ? '7 days' : '1 day';
  const penaltyWindowLower = longTerm ? 'last 14 days' : 'last 4 days';
  const penaltyWindowTitle = longTerm ? 'Last 14 days' : 'Last 4 days';
  const penaltyStartLabelStr = `Penalty Start date (${penaltyWindowTitle})`;
  const reminderClauseDaysStr = `${noticeDaysStr}-day notice period`;

  const page1 = document.getElementById('page-1');
  const page2 = document.getElementById('page-2');

  try {
    applyContractDistinctions(page1, longTerm);
    applyContractDistinctions(page2, longTerm);
    applyDailyRateToTemplate(page1, dailyRateFormatted);
    applyDailyRateToTemplate(page2, dailyRateFormatted);

    safeSetText('pdfTitleHeader', titleText);
    safeSetText('outNoticeDaysLabel', noticeDaysStr);
    safeSetText('outSigningWindow', noticeWindowStr);
    safeSetText('outNonRenewalNotice', nonRenewalNoticeStr);
    safeSetText('outAdvancePaymentWindow', advancePaymentWindowStr);
    safeSetText('outNoticeDaysClause7', noticeDaysStr);
    safeSetText('outReminderClausePeriod', reminderClauseDaysStr);
    safeSetText('outDayLabelOrdinal', dayOrdinalStr);
    safeSetText('outOverstayDays', overstayDaysStr);
    safeSetText('outPenaltyWindowLabel', penaltyWindowLower);
    safeSetText('outPenaltyStartWindowTitle', penaltyWindowTitle);
    safeSetText('outPenaltyStartLabel', penaltyStartLabelStr);
    safeSetText('outDailyRate', dailyRateFormatted);

    document.querySelectorAll('.outNoticeWindow').forEach(el => {
      el.textContent = noticeWindowStr;
    });
    document.querySelectorAll('.outPenaltyWindow').forEach(el => {
      el.textContent = penaltyWindowLower;
    });
    document.querySelectorAll('.outDailyRate').forEach(el => {
      el.textContent = dailyRateFormatted;
    });

    const todayDisplay = formatLongDate(tenant.date) || formatLongDate(new Date());
    safeSetText('outDate', todayDisplay);
    safeSetText('outTenant', tenant.tenantName || '');
    safeSetText('outBuilding', tenant.building || '');
    safeSetText('outUnit', tenant.unit || '');

    const noticeDeadlineDisplay = formatLongDate(tenant.computed20th) || tenant.computed20th || '_________________________';
    const contractEndDisplay = formatLongDate(tenant.endContract) || tenant.endContract || '________________________';

    let penaltyStartDisplay = '';
    const endIso = toDateInputValue(tenant.endContract);
    const noticeIso = toDateInputValue(tenant.computed20th);
    if (endIso) {
      const [y, m, d] = endIso.split('-').map(Number);
      const pDate = new Date(y, m - 1, d);
      pDate.setDate(pDate.getDate() - (longTerm ? 14 : 4));
      penaltyStartDisplay = formatLongDate(pDate);
    } else if (noticeIso) {
      const [y, m, d] = noticeIso.split('-').map(Number);
      const pDate = new Date(y, m - 1, d);
      pDate.setDate(pDate.getDate() + 6);
      penaltyStartDisplay = formatLongDate(pDate);
    } else {
      penaltyStartDisplay = formatLongDate(tenant.sixthDay) || tenant.sixthDay || '_________________________';
    }

    safeSetText('outEndContract', contractEndDisplay);
    safeSetText('outNoticeDeadline', noticeDeadlineDisplay);
    safeSetText('outPenaltyStartDate', penaltyStartDisplay);

    safeSetText('outCheckYes', (tenant.choice && tenant.choice.toUpperCase() === 'YES') ? '☑' : '☐');
    safeSetText('outCheckNo', (tenant.choice && tenant.choice.toUpperCase() === 'NO') ? '☑' : '☐');
    safeSetText('outReason', tenant.reason && tenant.reason !== '—' ? tenant.reason : '________________________________________');

    safeSetText('outTenant2', tenant.tenantName || '');
    safeSetText('outRent', formatPeso(calculatedRent));
    safeSetText('outEndContract2', contractEndDisplay);
    safeSetText('outSignature2', tenant.tenantName || '');

    safeSetText('outTermLabel', exactTermText);
    document.querySelectorAll('.outTermLabel').forEach(el => {
      el.textContent = exactTermText;
    });

    const propertyName = String(tenant.building || activeTab || '').toLowerCase().trim().replace(/\s+/g, ' ');
    const isBedAndBath = propertyName === 'bed and bath' || propertyName === 'b&b';
    safeSetText('outOccupancySupportFee', isBedAndBath ? '' : ', plus ₱298.00 monthly Occupancy Support Fee (except BED AND BATH property)');
  } catch (err) {
    throw new Error("Failed to populate template: " + err.message);
  }

  try {
    if (typeof html2canvas === 'undefined' || typeof window.jspdf === 'undefined') {
      throw new Error("Required libraries not loaded.");
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4');

    const canvas1 = await html2canvas(page1, { scale: 2, useCORS: true, windowWidth: 800 });
    doc.addImage(canvas1.toDataURL('image/jpeg', 0.98), 'JPEG', 0, 0, 210, 297);

    doc.addPage();
    const canvas2 = await html2canvas(page2, { scale: 2, useCORS: true, windowWidth: 800 });
    doc.addImage(canvas2.toDataURL('image/jpeg', 0.98), 'JPEG', 0, 0, 210, 297);

    return doc;
  } catch (err) {
    throw new Error("Canvas/PDF rendering failed: " + err.message);
  }
}

async function uploadPdfToDrive(doc, tenant, fileName) {
  if (!WEB_APP_URL || WEB_APP_URL === 'YOUR_APPS_SCRIPT_WEB_APP_URL_HERE') return '';
  const dataUri = doc.output('datauristring');
  const base64 = dataUri.split(',')[1];
  const res = await postToAppsScript({
    fileName,
    base64,
    rootFolderId: ROOT_FOLDER_ID,
    propertyName: tenant.building || activeTab || 'Unknown Property',
    monthName: tenant.folderMonth || 'Unknown Month',
    tenantName: tenant.tenantName || 'Unknown Tenant'
  });
  return res && res.fileUrl ? res.fileUrl : '';
}

window.generateSingle = async function(tenantName, unit) {
  const tenant = findTenantByNameAndUnit(activeTab, tenantName, unit);
  if (!tenant) return;

  const preservedTermLabel = String(tenant.termLabel || '').trim();
  document.getElementById('loadingOverlay').classList.remove('hidden');
  document.getElementById('loadingOverlay').classList.add('flex');
  document.getElementById('loadingText').textContent = `Generating PDF for ${tenant.tenantName}...`;

  try {
    const doc = await generatePDFDoc(tenant);
    const noticePrefix = isShortTerm(preservedTermLabel) ? '10-Day Notice' : '20-Day Notice';
    const fileName = `${noticePrefix} - ${tenant.tenantName}.pdf`;

    const uploadedUrl = await uploadPdfToDrive(doc, tenant, fileName);
    if (uploadedUrl) tenant.pdfUrl = uploadedUrl;

    tenant.termLabel = preservedTermLabel;
    tenant.hasExplicitTerm = true;
    tenant.status = 'Generated';
    await syncToDatabase({ action: 'update', tabName: activeTab, tenant });
    renderSpreadsheet();
  } catch (error) {
    console.error(error);
    alert("Error generating PDF: " + error.message);
  } finally {
    document.getElementById('loadingOverlay').classList.add('hidden');
    document.getElementById('loadingOverlay').classList.remove('flex');
  }
};

window.regenerateSingle = async function(tenantName, unit) {
  const tenant = findTenantByNameAndUnit(activeTab, tenantName, unit);
  if (!tenant) return;

  const preservedTermLabel = String(tenant.termLabel || '').trim();
  document.getElementById('loadingOverlay').classList.remove('hidden');
  document.getElementById('loadingOverlay').classList.add('flex');
  document.getElementById('loadingText').textContent = `Regenerating PDF for ${tenant.tenantName}...`;

  try {
    const doc = await generatePDFDoc(tenant);
    const noticePrefix = isShortTerm(preservedTermLabel) ? '10-Day Notice' : '20-Day Notice';
    const fileName = `${noticePrefix} - ${tenant.tenantName}.pdf`;

    const uploadedUrl = await uploadPdfToDrive(doc, tenant, fileName);
    if (uploadedUrl) tenant.pdfUrl = uploadedUrl;

    tenant.termLabel = preservedTermLabel;
    tenant.hasExplicitTerm = true;
    await syncToDatabase({ action: 'update', tabName: activeTab, tenant });
    renderSpreadsheet();
  } catch (error) {
    console.error(error);
    alert("Error regenerating PDF: " + error.message);
  } finally {
    document.getElementById('loadingOverlay').classList.add('hidden');
    document.getElementById('loadingOverlay').classList.remove('flex');
  }
};

async function batchGenerateTenants(pendingTenants, selectionLabel) {
  if (!activeTab) return alert("No active tab selected.");
  if (pendingTenants.length === 0) {
    return alert(selectionLabel === 'selected'
      ? "There are no selected Pending tenants to generate."
      : "There are no Pending tenants in this tab to generate.");
  }

  const scopeText = selectionLabel === 'selected' ? 'selected Pending' : 'Pending';
  const confirmGenerate = confirm(`Generate PDFs for ${selectionLabel === 'selected' ? '' : 'all '}${pendingTenants.length} ${scopeText} tenants in "${activeTab}"?`);
  if (!confirmGenerate) return;

  document.getElementById('loadingOverlay').classList.remove('hidden');
  document.getElementById('loadingOverlay').classList.add('flex');

  try {
    for (let i = 0; i < pendingTenants.length; i++) {
      const tenant = pendingTenants[i];
      const preservedTermLabel = String(tenant.termLabel || '').trim();
      document.getElementById('loadingText').textContent = `Generating (${i + 1}/${pendingTenants.length}): ${tenant.tenantName}...`;

      const doc = await generatePDFDoc(tenant);
      const noticePrefix = isShortTerm(preservedTermLabel) ? '10-Day Notice' : '20-Day Notice';
      const fileName = `${noticePrefix} - ${tenant.tenantName}.pdf`;

      const uploadedUrl = await uploadPdfToDrive(doc, tenant, fileName);
      if (uploadedUrl) tenant.pdfUrl = uploadedUrl;

      tenant.termLabel = preservedTermLabel;
      tenant.hasExplicitTerm = true;
      tenant.status = 'Generated';
      await syncToDatabase({ action: 'update', tabName: activeTab, tenant });
      renderSpreadsheet();
    }

    alert(`Successfully generated ${pendingTenants.length} ${selectionLabel === 'selected' ? 'selected ' : ''}notices.`);
  } catch (error) {
    console.error(error);
    alert("An error occurred during batch generation: " + error.message);
  } finally {
    document.getElementById('loadingOverlay').classList.add('hidden');
    document.getElementById('loadingOverlay').classList.remove('flex');
  }
}

document.getElementById('batchGenerateBtn')?.addEventListener('click', () => {
  if (!activeTab) return alert("No active tab selected.");
  const pendingTenants = appData[activeTab].filter(tenant => tenant.status === 'Pending');
  batchGenerateTenants(pendingTenants, 'all');
});

document.getElementById('batchGenerateSelectedBtn')?.addEventListener('click', () => {
  if (!activeTab) return alert("No active tab selected.");
  const pendingTenants = appData[activeTab].filter(tenant => selectedTenants.has(tenant) && tenant.status === 'Pending');
  batchGenerateTenants(pendingTenants, 'selected');
});

document.getElementById('batchDownloadZipBtn')?.addEventListener('click', async () => {
  if (!activeTab) return alert("No active tab selected.");

  const generatedTenants = appData[activeTab].filter(t => selectedTenants.has(t));

  if (generatedTenants.length === 0) {
    return alert("Select one or more tenants to download as a ZIP.");
  }

  const confirmZip = confirm(`Download ${generatedTenants.length} selected PDFs as a ZIP file?`);
  if (!confirmZip) return;

  document.getElementById('loadingOverlay').classList.remove('hidden');
  document.getElementById('loadingOverlay').classList.add('flex');
  document.getElementById('loadingText').textContent = `Preparing ZIP...`;

  try {
    const zip = new JSZip();
    let count = 0;
    const failedGenerations = [];

    for (let i = 0; i < generatedTenants.length; i++) {
      const tenant = generatedTenants[i];

      try {
        document.getElementById('loadingText').textContent = `Generating PDF (${i + 1}/${generatedTenants.length}): ${tenant.tenantName}...`;
        const doc = await generatePDFDoc(tenant);
        const blob = doc.output('blob');
        const noticePrefix = isShortTerm(tenant.termLabel) ? '10-Day Notice' : '20-Day Notice';
        const fileName = `${noticePrefix} - ${tenant.tenantName}.pdf`;

        const propertyName = tenant.building || "Unknown Property";

        let monthName = tenant.folderMonth;
        if (!monthName || monthName === 'Unknown Month' || monthName.trim() === '') {
          if (tenant.computed20th && tenant.computed20th.includes('-')) {
            monthName = new Date(tenant.computed20th).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
          } else {
            monthName = new Date().toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
          }
        }

        const tenantFolder = tenant.tenantName || "Unknown Tenant";
        const folderPath = `${propertyName}/${monthName}/${tenantFolder}/`;

        zip.file(folderPath + fileName, blob);
        count++;
      } catch (err) {
        console.error(`Could not generate PDF for ${tenant.tenantName}:`, err);
        failedGenerations.push(tenant.tenantName);
      }
    }

    if (count === 0) throw new Error("Could not generate any PDF files.");

    document.getElementById('loadingText').textContent = `Zipping files...`;
    const content = await zip.generateAsync({ type: "blob" });

    const zipFileName = `Lease Renewals - ${activeTab} - ${new Date().toISOString().split('T')[0]}.zip`;

    const a = document.createElement("a");
    a.href = URL.createObjectURL(content);
    a.download = zipFileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(a.href);
    if (failedGenerations.length > 0) {
      alert(`Added ${count} of ${generatedTenants.length} selected PDFs to the ZIP. Could not generate: ${failedGenerations.join(', ')}`);
    }
  } catch (error) {
    console.error(error);
    alert("Error creating ZIP: " + error.message);
  } finally {
    document.getElementById('loadingOverlay').classList.add('hidden');
    document.getElementById('loadingOverlay').classList.remove('flex');
  }
});

window.markAsSent = async function(tenantName, unit) {
  const tenant = findTenantByNameAndUnit(activeTab, tenantName, unit);
  if (!tenant) return;
  tenant.status = 'Sent';
  if (!tenant.dateSent) {
    tenant.dateSent = formatDate(new Date());
  }
  await syncToDatabase({ action: 'update', tabName: activeTab, tenant });
  renderSpreadsheet();
};

document.getElementById('clearCurrentTabBtn')?.addEventListener('click', async () => {
  if (!activeTab) return alert("No active tab selected.");

  const confirmClear = confirm(`Clear all rows from "${activeTab}" in the current view? (Click Sync Main Tenants to re-fetch from Google Sheets.)`);
  if (confirmClear) {
    appData[activeTab] = [];
    selectedTenants.clear();
    await syncToDatabase({ action: 'clearTab', tabName: activeTab });
    renderSpreadsheet();
    alert(`Tab "${activeTab}" cleared.`);
  }
});

document.getElementById('addTenantFormBtn')?.addEventListener('click', async () => {
  const tenantName = document.getElementById('inputTenant').value;
  if (!tenantName) return alert("Enter Tenant Name");
  if (!activeTab) return alert("Select a tab first.");

  const startContract = document.getElementById('inputStartDate').value;
  const endContract = document.getElementById('inputEndContract').value;
  const term = document.getElementById('inputTerm').value;
  const dates = endContract && !startContract
    ? deriveFromContractEndDate(endContract, term)
    : parseDate(startContract, term);

  const newTenant = {
    tenantName: tenantName,
    building: document.getElementById('inputBuilding').value,
    unit: document.getElementById('inputUnit').value,
    rent: document.getElementById('inputRent').value,
    computed20th: dates.computed20th,
    sixthDay: dates.sixthDay,
    folderMonth: dates.folderMonth,
    termLabel: formatTermLabel(term),
    hasExplicitTerm: true,
    status: 'Pending',
    pdfUrl: '',
    reason: '—',
    dateSent: '',
    startDate: startContract,
    endContract
  };

  appData[activeTab].push(newTenant);
  await syncToDatabase({ action: 'save', tabName: activeTab, tenant: newTenant });
  renderSpreadsheet();
  document.getElementById('inputTenant').value = '';
  document.getElementById('inputEndContract').value = '';
});

document.getElementById('recalculateDatesBtn')?.addEventListener('click', async () => {
  if (!activeTab) return alert("No active tab selected.");

  const confirmRecalc = confirm(`Recalculate ALL dates in the "${activeTab}" tab?`);
  if (!confirmRecalc) return;

  const tenants = appData[activeTab];
  for (let i = 0; i < tenants.length; i++) {
    const tenant = tenants[i];
    if (tenant.endContract) {
      const dates = deriveFromContractEndDate(tenant.endContract, tenant.termLabel);
      tenant.computed20th = dates.computed20th;
      tenant.sixthDay = dates.sixthDay;
      tenant.folderMonth = dates.folderMonth;
    } else if (tenant.startDate && tenant.termLabel) {
      const dates = parseDate(tenant.startDate, tenant.termLabel);
      tenant.computed20th = dates.computed20th;
      tenant.sixthDay = dates.sixthDay;
      tenant.folderMonth = dates.folderMonth;
    }
    await syncToDatabase({ action: 'update', tabName: activeTab, tenant });
  }

  renderSpreadsheet();
  alert("Dates recalculated successfully!");
});

document.getElementById('excelInput')?.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  document.getElementById('loadingOverlay').classList.remove('hidden');
  document.getElementById('loadingOverlay').classList.add('flex');
  document.getElementById('loadingText').textContent = "Reading Excel File...";

  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const data = new Uint8Array(evt.target.result);
      workbookCache = XLSX.read(data, { type: 'array', cellDates: true });

      const sheetSelect = document.getElementById('sheetSelect');
      sheetSelect.innerHTML = '';
      workbookCache.SheetNames.forEach(sheetName => {
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        sheetSelect.appendChild(option);
      });
      document.getElementById('sheetSelectionArea').classList.remove('hidden');
    } catch (error) {
      alert("Error reading file.");
    } finally {
      document.getElementById('loadingOverlay').classList.add('hidden');
      document.getElementById('loadingOverlay').classList.remove('flex');
      document.getElementById('excelInput').value = '';
    }
  };
  reader.readAsArrayBuffer(file);
});

document.getElementById('confirmLoadSheetBtn')?.addEventListener('click', async () => {
  if (!workbookCache) return;
  const selectedSheetName = document.getElementById('sheetSelect').value;
  document.getElementById('loadingOverlay').classList.remove('hidden');
  document.getElementById('loadingOverlay').classList.add('flex');

  try {
    const worksheet = workbookCache.Sheets[selectedSheetName];
    const json = XLSX.utils.sheet_to_json(worksheet, { raw: true });
    if (json.length === 0) return alert("Sheet empty.");
    if (!appData[selectedSheetName]) appData[selectedSheetName] = [];

    for (let i = 0; i < json.length; i++) {
      document.getElementById('loadingText').textContent = `Loading Data... (${i + 1}/${json.length})`;

      const row = json[i];
      const rowKeys = Object.keys(row).reduce((acc, key) => {
        acc[key.trim().toUpperCase()] = row[key];
        return acc;
      }, {});

      const startContract = rowKeys['START OF CONTRACT'] || rowKeys['START DATE'] || '';
      const endContract = rowKeys['END OF CONTRACT'] || rowKeys['END CONTRACT'] || '';
      const rawExcelTerm = String(rowKeys['LEASE TERM'] || rowKeys['TERM'] || '').trim();
      const term = rawExcelTerm || '12';
      const dates = endContract && !startContract
        ? deriveFromContractEndDate(endContract, term)
        : parseDate(startContract, term);

      const newTenant = {
        tenantName: rowKeys['TENANT NAME'] || rowKeys['NAME'] || 'Unknown',
        building: rowKeys['BUILDING'] || "Eve's Residences",
        unit: rowKeys['UNIT'] || '',
        rent: rowKeys['MONTHLY RENT'] || 0,
        computed20th: dates.computed20th,
        sixthDay: dates.sixthDay,
        folderMonth: dates.folderMonth,
        termLabel: formatTermLabel(term),
        hasExplicitTerm: rawExcelTerm !== '',
        status: 'Pending',
        pdfUrl: '',
        reason: rowKeys['REASON'] || '—',
        dateSent: '',
        startDate: startContract,
        endContract: endContract
      };

      appData[selectedSheetName].push(newTenant);
      await syncToDatabase({ action: 'save', tabName: selectedSheetName, tenant: newTenant });
    }

    selectedTenants.clear();
    activeTab = selectedSheetName;
    document.getElementById('searchInput').value = '';
    document.getElementById('statusFilter').value = 'ALL';
    renderTabs();
    renderSpreadsheet();
    document.getElementById('sheetSelectionArea').classList.add('hidden');
    workbookCache = null;
  } catch (error) {
    alert("Error loading sheet.");
  } finally {
    document.getElementById('loadingOverlay').classList.add('hidden');
    document.getElementById('loadingOverlay').classList.remove('flex');
  }
});

window.onload = () => initDatabase(false, true);