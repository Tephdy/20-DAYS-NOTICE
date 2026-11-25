const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwvTv-hCnZoKT87dtu7nNq5E65qWhLCOp_NHrP_-uRNavNU9quBJyq7cRPUPacwTX6n/exec';
    const ROOT_FOLDER_ID = '16qLfbrPeU-H_UH0QUl2lIqUoqf5_nVLQ'; 
    const MAIN_SPREADSHEET_ID = '1eQUsnIRkXBqohI_6F-oV95bKZAn75lilbkx4KEqd4dw';
    const MAIN_SPREADSHEET_ADDITIONAL_TABS = ['PENTHAUZ'];

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

    async function initDatabase() {
      document.getElementById('loadingOverlay').classList.remove('hidden');
      document.getElementById('loadingOverlay').classList.add('flex');
      document.getElementById('loadingText').textContent = "Loading Data from Google Sheets...";

      try {
        const response = await fetch(APPS_SCRIPT_URL, {
          method: 'POST',
          body: JSON.stringify({ action: 'load' })
        });
        const res = await response.json();

        if (res.success) {
          appData = res.data;
          MAIN_SPREADSHEET_ADDITIONAL_TABS.forEach(tabName => {
            if (!Array.isArray(appData[tabName])) appData[tabName] = [];
          });
          for (const tabName in appData) {
            if (!Array.isArray(appData[tabName])) continue;
            appData[tabName].forEach(tenant => {
              tenant.pdfUrl = tenant.pdfUrl || '';
              if (!tenant.status) tenant.status = 'Pending';
              if (!tenant.dateSent) tenant.dateSent = '';
              if (!tenant.reason) tenant.reason = '—';
              if (tenant.endContract) tenant.endContract = normalizeMainSheetDate(tenant.endContract) || tenant.endContract;
              
              if (tenant.startDate && tenant.termLabel) {
                const dates = parseDate(tenant.startDate, tenant.termLabel);
                tenant.computed20th = dates.computed20th;
                tenant.sixthDay = dates.sixthDay;
                tenant.folderMonth = dates.folderMonth;
              } else if (tenant.computed20th) {
                const d = new Date(tenant.computed20th);
                if (!isNaN(d.getTime())) {
                  tenant.computed20th = d.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
                }
              }
            });
          }

          const tabNames = Object.keys(appData);
          if (tabNames.length > 0) {
            activeTab = tabNames[0];
            renderTabs();
            renderSpreadsheet();
          } else {
            document.getElementById('spreadsheetBody').innerHTML = `<tr><td colspan="9" class="px-4 py-12 text-center text-slate-400">No tabs found. Import a sheet to begin.</td></tr>`;
          }
        } else {
          alert("Failed to load database: " + res.error);
        }
      } catch (error) {
        console.error(error);
        alert("Could not connect to the Google Apps Script.");
      } finally {
        document.getElementById('loadingOverlay').classList.add('hidden');
        document.getElementById('loadingOverlay').classList.remove('flex');
      }
    }

    async function syncToDatabase(tenant, tabName, actionType) {
      try {
        const res = await fetch(APPS_SCRIPT_URL, {
          method: 'POST',
          body: JSON.stringify({
            action: actionType, 
            tabName: tabName,
            tenantName: tenant.tenantName,
            unit: tenant.unit,
            building: tenant.building,
            computed20th: tenant.computed20th,
            rent: tenant.rent,
            status: tenant.status,
            pdfUrl: tenant.pdfUrl || '',
            termLabel: tenant.termLabel,
            reason: tenant.reason,
            dateSent: tenant.dateSent || '',
            startDate: tenant.startDate || '',
            endContract: tenant.endContract || ''
          })
        });
        return await res.json();
      } catch (e) {
        console.error(e);
        return { success: false };
      }
    }

    function normalizeSyncValue(value) {
      return String(value == null ? '' : value).trim().toLowerCase().replace(/\s+/g, ' ');
    }

    function normalizeSyncHeader(value) {
      return String(value == null ? '' : value).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
    }

    function normalizeSyncTerm(value) {
      const term = normalizeSyncValue(value);
      if (term.includes('ext') && term.includes('trial')) return 'extended trial';
      if (term.includes('trial')) return 'trial';
      if (term.includes('monthly') || term === 'm-m' || term === 'mm' || term === 'month-month') return 'monthly';
      if (/\b24\b|2 years?|two years?/.test(term)) return '24';
      if (/\b6\b|6 months?/.test(term)) return '6';
      if (/\b12\b|1 year|one year/.test(term)) return '12';
      return term;
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
      return { computed20th, ...deriveFromComputed20th(computed20th) };
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
          finish(null, response.table.rows.map(row =>
            (row.c || []).map(cell => parseMainSheetValue(cell && cell.v, cell && cell.f))
          ));
        };
        script.onerror = () => finish(new Error(`Could not access the "${tabName}" tab. Check that the source spreadsheet is shared for viewing.`));

        const endpoint = `https://docs.google.com/spreadsheets/d/${MAIN_SPREADSHEET_ID}/gviz/tq`;
        const params = new URLSearchParams({
          tqx: `out:json;responseHandler:${callbackName}`,
          sheet: tabName,
          headers: '0'
        });
        script.src = `${endpoint}?${params.toString()}`;
        document.head.appendChild(script);
      });
    }

    function extractMainTenantRows(rows) {
      const headerAliases = {
        tenantName: ['name', 'tenant name', 'tenant'],
        unit: ['unit', 'unit no', 'unit number'],
        building: ['building', 'property', 'property name'],
        startDate: ['start of contract', 'contract start', 'lease start', 'start date', 'renewal/movein', 'renewal', 'movein', 'move in date'],
        endDate: ['end of contract', 'end contract', 'end contract date', 'contract end', 'contract end date', 'lease end', 'lease end date', 'end date', 'expiration date', 'contract expiration', 'lease expiration', 'expiry date', 'contract expiry', 'contract expiration date'],
        term: ['lease term', 'term', 'contract term', 'contract'],
        rent: ['monthly rent', 'rent', 'rate', 'monthly rate'],
        reason: ['reason']
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
        throw new Error('Could not find a header row containing both Name and Unit.');
      }

      const headers = rows[headerRowIndex].map(normalizeSyncHeader);
      const columnIndexes = {};
      Object.entries(aliasSets).forEach(([field, aliases]) => {
        columnIndexes[field] = headers.findIndex(header => aliases.has(header));
      });
      if (columnIndexes.startDate === -1 && headers.length > 8) columnIndexes.startDate = 8;
      if (columnIndexes.rent === -1 && headers.length > 5) columnIndexes.rent = 5;
      if (columnIndexes.endDate === -1) {
        columnIndexes.endDate = headers.findIndex(header =>
          header.includes('contract') && (header.includes('end') || header.includes('expir') || header.includes('expiry'))
        );
      }
      if (columnIndexes.endDate === -1 && headers.length > 9) {
        columnIndexes.endDate = 9;
      }
      return rows.slice(headerRowIndex + 1).map(row => {
        const record = {};
        Object.entries(columnIndexes).forEach(([field, index]) => {
          if (index >= 0) record[field] = row[index] == null ? '' : row[index];
        });
        return record;
      }).filter(record => normalizeSyncValue(record.tenantName) && normalizeSyncValue(record.unit));
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

    async function syncMainSpreadsheet() {
      const button = document.getElementById('syncMainSpreadsheetBtn');
      const loadingOverlay = document.getElementById('loadingOverlay');
      const summary = { imported: 0, updated: 0, unchanged: 0, unmatched: 0, ambiguous: 0, endContractFetched: 0, endContractUpdated: 0, failedTabs: [], renameWarnings: [] };
      button.disabled = true;
      loadingOverlay.classList.remove('hidden');
      loadingOverlay.classList.add('flex');

      try {
        const tabNames = Array.from(new Set([
          ...Object.keys(appData).filter(tabName => Array.isArray(appData[tabName])),
          ...MAIN_SPREADSHEET_ADDITIONAL_TABS
        ]));
        if (tabNames.length === 0) throw new Error('There are no property tabs to synchronize.');

        for (let tabIndex = 0; tabIndex < tabNames.length; tabIndex++) {
          const tabName = tabNames[tabIndex];
          document.getElementById('loadingText').textContent = `Checking ${tabName} (${tabIndex + 1}/${tabNames.length})...`;
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
              const nameKey = normalizeSyncValue(source.tenantName);
              const nameMatches = localRecords.filter(tenant => normalizeSyncValue(tenant.tenantName) === nameKey);
              if (sourceNameCounts.get(nameKey) === 1 && localNameCounts.get(nameKey) === 1 && nameMatches.length === 1) {
                target = nameMatches[0];
              } else {
                const unitKey = normalizeSyncValue(source.unit);
                const unitMatches = localRecords.filter(tenant => normalizeSyncValue(tenant.unit) === unitKey);
                if (sourceUnitCounts.get(unitKey) === 1 && localUnitCounts.get(unitKey) === 1 && unitMatches.length === 1) {
                  target = unitMatches[0];
                }
              }
            }

            if (!target || processedTenants.has(target)) {
              const canImportMissingTenant = !target &&
                MAIN_SPREADSHEET_ADDITIONAL_TABS.includes(tabName) &&
                sourcePairCounts.get(pairKey) === 1 &&
                exactMatches.length === 0;
              if (canImportMissingTenant) {
                const sourceEndDate = normalizeMainSheetDate(source.endDate);
                const sourceStartDate = normalizeMainSheetDate(source.startDate);
                const sourceTerm = source.term || '12';
                const parsedTerm = parseDate('2000-01-01', sourceTerm);
                const dates = sourceEndDate
                  ? deriveFromContractEndDate(sourceEndDate, sourceTerm)
                  : sourceStartDate
                    ? parseDate(sourceStartDate, sourceTerm)
                    : { computed20th: '', sixthDay: '_______________', folderMonth: 'Unknown Month' };
                const importedTenant = {
                  tenantName: String(source.tenantName).trim(),
                  unit: String(source.unit).trim(),
                  building: source.building || tabName,
                  rent: normalizeMainSheetRent(source.rent) || 0,
                  computed20th: dates.computed20th,
                  sixthDay: dates.sixthDay,
                  folderMonth: dates.folderMonth,
                  termLabel: parsedTerm.termLabel || String(sourceTerm),
                  status: 'Pending',
                  pdfUrl: '',
                  reason: source.reason || '—',
                  dateSent: '',
                  startDate: sourceStartDate,
                  endContract: sourceEndDate
                };
                const saveResult = await syncToDatabase(importedTenant, tabName, 'save');
                if (!saveResult || !saveResult.success) {
                  throw new Error(`Could not import "${importedTenant.tenantName}" into "${tabName}".`);
                }
                localRecords.push(importedTenant);
                processedTenants.add(importedTenant);
                summary.imported++;
                if (sourceEndDate) {
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

            const updated = { ...target };
            const setIfPresent = (field, value) => {
              if (value !== '' && value != null) updated[field] = value;
            };
            setIfPresent('tenantName', source.tenantName);
            setIfPresent('unit', source.unit);
            setIfPresent('building', source.building);
            setIfPresent('rent', normalizeMainSheetRent(source.rent));
            setIfPresent('reason', source.reason);

            const sourceStartDate = normalizeMainSheetDate(source.startDate);
            const sourceEndDate = normalizeMainSheetDate(source.endDate);
            if (sourceEndDate) {
              summary.endContractFetched++;
              if (sourceEndDate !== normalizeMainSheetDate(target.endContract)) {
                summary.endContractUpdated++;
              }
            }
            const startDate = sourceStartDate || target.startDate;
            const term = source.term || target.termLabel;
            const startDateChanged = sourceStartDate && sourceStartDate !== normalizeMainSheetDate(target.startDate);
            const termChanged = source.term && normalizeSyncTerm(source.term) !== normalizeSyncTerm(target.termLabel);
            if (sourceStartDate) updated.startDate = sourceStartDate;
            if (sourceEndDate) updated.endContract = sourceEndDate;
            if (startDate && term && (startDateChanged || termChanged)) {
              const dates = parseDate(startDate, term);
              if (dates.computed20th && dates.computed20th !== 'No Date') {
                updated.computed20th = dates.computed20th;
                updated.sixthDay = dates.sixthDay;
                updated.folderMonth = dates.folderMonth;
                updated.termLabel = dates.termLabel;
              }
            }
            if (sourceEndDate) {
              const endDateDerived = deriveFromContractEndDate(sourceEndDate, term);
              updated.computed20th = endDateDerived.computed20th;
              updated.sixthDay = endDateDerived.sixthDay;
              updated.folderMonth = endDateDerived.folderMonth;
            }

            const changed = ['tenantName', 'unit', 'building', 'rent', 'reason', 'startDate', 'endContract', 'termLabel', 'computed20th']
              .some(field => normalizeSyncValue(updated[field]) !== normalizeSyncValue(target[field]));
            if (!changed) {
              summary.unchanged++;
              continue;
            }

            document.getElementById('loadingText').textContent = `Updating ${target.tenantName} (${tabName})...`;
            if (updated.tenantName !== target.tenantName) {
              const saveResult = await syncToDatabase(updated, tabName, 'save');
              if (!saveResult || !saveResult.success) {
                throw new Error(`The renamed tenant "${updated.tenantName}" could not be saved in "${tabName}".`);
              }
              try {
                const deleteResponse = await fetch(APPS_SCRIPT_URL, {
                  method: 'POST',
                  body: JSON.stringify({ action: 'delete', tabName, tenantName: target.tenantName })
                });
                const deleteResult = await deleteResponse.json();
                if (!deleteResponse.ok || !deleteResult.success) {
                  const reason = deleteResult.error || deleteResult.message || `HTTP ${deleteResponse.status}`;
                  summary.renameWarnings.push(`${tabName}: saved "${updated.tenantName}", but could not remove old entry "${target.tenantName}" (${reason}).`);
                }
              } catch (error) {
                console.error(error);
                summary.renameWarnings.push(`${tabName}: saved "${updated.tenantName}", but could not confirm removal of old entry "${target.tenantName}" (${error.message}).`);
              }
            } else {
              const updateResult = await syncToDatabase(updated, tabName, 'update');
              if (!updateResult || !updateResult.success) {
                throw new Error(`Could not update "${target.tenantName}" in "${tabName}".`);
              }
            }
            Object.assign(target, updated);
            summary.updated++;
          }
        }

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
        if (summary.renameWarnings.length) details.push(`Rename cleanup warnings:\n${summary.renameWarnings.join('\n')}`);
        alert(`Main spreadsheet sync complete.\n\n${details.join('\n')}`);
      } catch (error) {
        console.error(error);
        renderTabs();
        renderSpreadsheet();
        alert(`Main spreadsheet sync stopped: ${error.message}`);
      } finally {
        button.disabled = false;
        loadingOverlay.classList.add('hidden');
        loadingOverlay.classList.remove('flex');
      }
    }

    document.getElementById('syncMainSpreadsheetBtn').addEventListener('click', syncMainSpreadsheet);

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
      const isoDateMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
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
        if (date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day) {
          return formatDate(date);
        }
        return '';
      }
      const d = new Date(s);
      return isNaN(d) ? '' : formatDate(d);
    }

    function deriveFromComputed20th(dateStr) {
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

    function isShortTerm(term) {
      const termString = String(term).toUpperCase().trim();
      return (
        termString === 'M-M' || 
        termString === 'MM' || 
        termString.includes('MONTHLY') || 
        termString.includes('TRIAL') || 
        termString === 'MONTH-MONTH'
      );
    }

    function parseDate(startContract, term) {
      let computed20th = '';
      let sixthDay = '';
      let termLabel = '';
      let termMonths = 12;
      let folderMonth = 'Unknown Month'; 

      let startD;
      if (typeof startContract === 'number') {
        startD = new Date(Math.round((startContract - 25569) * 86400 * 1000));
      } else if (startContract instanceof Date) {
        startD = startContract;
      } else if (typeof startContract === 'string') {
        startD = new Date(startContract.trim()); 
      } else {
        startD = new Date(startContract);
      }

      if (!isNaN(startD.getTime())) {
        const termString = String(term).toUpperCase().trim();
        const shortTerm = isShortTerm(term);

        if (shortTerm) {
          termLabel = termString.includes('TRIAL') ? term : 'Monthly';
          termMonths = 1;
        } else {
          termMonths = Number(term) || 12;
          termLabel = termMonths === 6 ? '6 months' : termMonths === 12 ? '1 year (12 months)' : termMonths === 24 ? '2 years (24 months)' : termMonths + ' months';
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
          } catch(e) {}
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
          } catch(e) { matchesMonth = false; }
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
          
          let dateSentHtml = '';
          dateSentHtml = `<input type="text" inputmode="numeric" placeholder="M/D/YYYY" class="table-date-input" value="${formatDisplayDate(tenant.dateSent)}" ${disabledAttr} onchange="updateDateSent('${tenant.tenantName.replace(/'/g, "\\'")}', this.value)" aria-label="Date sent (M/D/YYYY)">`;

          const safeName = tenant.tenantName.replace(/'/g, "\\'");

          const computed20thValue = formatDisplayDate(tenant.computed20th);
          const computed20thHtml = `<input type="text" inputmode="numeric" placeholder="M/D/YYYY" class="table-date-input" value="${computed20thValue}" onchange="updateComputed20th('${safeName}', this.value)" title="Edit Notice Deadline (M/D/YYYY)" aria-label="Notice deadline (M/D/YYYY)">`;
          const endContractDisplay = formatDisplayDate(tenant.endContract) || '—';

          let actionBtn = '';
          if (tenant.status === 'Pending') {
            actionBtn = `
              <div class="flex justify-end items-center gap-1">
                <button onclick="generateSingle('${safeName}')" class="bg-slate-800 text-white px-2 py-1 rounded text-xs font-bold hover:bg-slate-900 transition-colors" title="Generate PDF">Generate</button>
                <button onclick="openEditModal('${safeName}')" class="bg-blue-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-blue-700 transition-colors" title="Edit tenant">Edit</button>
                <button onclick="deleteTenant('${safeName}')" class="bg-red-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-red-700 transition-colors" title="Delete tenant">Del</button>
              </div>`;
          } else if (tenant.status === 'Generated') {
            actionBtn = `
              <div class="flex justify-end items-center gap-1 flex-wrap">
                <button onclick="regenerateSingle('${safeName}')" class="bg-slate-500 text-white px-2 py-1 rounded text-xs font-bold hover:bg-slate-600 transition-colors" title="Re-generate">Regenerate</button>
                <button onclick="markAsSent('${safeName}')" class="bg-green-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-green-700 transition-colors">Sent</button>
                <button onclick="openEditModal('${safeName}')" class="bg-blue-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-blue-700 transition-colors" title="Edit tenant">Edit</button>
                <button onclick="deleteTenant('${safeName}')" class="bg-red-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-red-700 transition-colors" title="Delete tenant">Del</button>
              </div>`;
          } else {
            actionBtn = `
              <div class="flex justify-end items-center gap-1">
                <button onclick="openEditModal('${safeName}')" class="bg-blue-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-blue-700 transition-colors" title="Edit tenant">Edit</button>
                <button onclick="deleteTenant('${safeName}')" class="bg-red-600 text-white px-2 py-1 rounded text-xs font-bold hover:bg-red-700 transition-colors" title="Delete tenant">Del</button>
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
      const selectedCount = activeTab && appData[activeTab]
        ? appData[activeTab].filter(tenant => selectedTenants.has(tenant) && tenant.pdfUrl && String(tenant.pdfUrl).trim() !== '').length
        : 0;
      const selectedPendingCount = activeTab && appData[activeTab]
        ? appData[activeTab].filter(tenant => selectedTenants.has(tenant) && tenant.status === 'Pending').length
        : 0;
      const generateSelectedButton = document.getElementById('batchGenerateSelectedBtn');
      document.getElementById('batchGenerateSelectedLabel').textContent = `Generate Selected (${selectedPendingCount})`;
      generateSelectedButton.disabled = selectedPendingCount === 0;
      generateSelectedButton.classList.toggle('opacity-50', selectedPendingCount === 0);
      generateSelectedButton.classList.toggle('cursor-not-allowed', selectedPendingCount === 0);

      const downloadButton = document.getElementById('batchDownloadZipBtn');
      document.getElementById('batchDownloadZipLabel').textContent = `Download Selected ZIP (${selectedCount})`;
      downloadButton.disabled = selectedCount === 0;
      downloadButton.classList.toggle('opacity-50', selectedCount === 0);
      downloadButton.classList.toggle('cursor-not-allowed', selectedCount === 0);

      const selectAllCheckbox = document.getElementById('selectAllTenantsCheckbox');
      const visibleCheckboxes = Array.from(document.querySelectorAll('#spreadsheetBody .tenant-select-checkbox'));
      const visibleSelectedCount = visibleCheckboxes.filter(checkbox => checkbox.checked).length;
      selectAllCheckbox.disabled = visibleCheckboxes.length === 0;
      selectAllCheckbox.checked = visibleCheckboxes.length > 0 && visibleSelectedCount === visibleCheckboxes.length;
      selectAllCheckbox.indeterminate = visibleSelectedCount > 0 && visibleSelectedCount < visibleCheckboxes.length;
    }

    document.getElementById('selectAllTenantsCheckbox').addEventListener('change', event => {
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
      document.getElementById('editBuilding').value = "Eve's Residences";
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

    window.openEditModal = function(tenantName) {
      if (!activeTab) return;
      const tenant = appData[activeTab].find(t => t.tenantName === tenantName);
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
      if (tl.includes('monthly') || tl.includes('m-m')) termVal = 'M-M';
      else if (tl.includes('ext. trial')) termVal = 'Ext. Trial';
      else if (tl.includes('trial')) termVal = 'Trial';
      else if (tl.includes('6')) termVal = '6';
      else if (tl.includes('2 years') || tl.includes('24')) termVal = '24';
      else if (tl.includes('1 year') || tl.includes('12')) termVal = '12';
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
        derived = deriveFromComputed20th(manualComputed20th);
      } else {
        finalComputed20th = autoDates.computed20th;
        derived = { sixthDay: autoDates.sixthDay, folderMonth: autoDates.folderMonth };
      }

      document.getElementById('loadingOverlay').classList.remove('hidden');
      document.getElementById('loadingOverlay').classList.add('flex');
      document.getElementById('loadingText').textContent = isAddMode ? 'Creating tenant...' : 'Saving changes...';

      try {
        if (isAddMode) {
          if (appData[activeTab].some(t => t.tenantName === name)) {
            throw new Error("A tenant with this name already exists in this tab.");
          }
          const newTenant = {
            tenantName: name,
            unit, building, rent,
            computed20th: finalComputed20th,
            sixthDay: derived.sixthDay,
            folderMonth: derived.folderMonth,
            termLabel: autoDates.termLabel,
            status,
            pdfUrl: '',
            reason,
            dateSent,
            startDate,
            endContract
          };
          appData[activeTab].push(newTenant);
          await syncToDatabase(newTenant, activeTab, 'save');
        } else {
          if (!currentEditingTenant) throw new Error("No tenant selected.");
          const originalName = currentEditingTenant.tenantName;
          
          if (name !== originalName && appData[activeTab].some(t => t.tenantName === name)) {
            throw new Error("Another tenant with this name already exists.");
          }

          const updated = {
            ...currentEditingTenant,
            tenantName: name,
            unit, building, rent,
            computed20th: finalComputed20th,
            sixthDay: derived.sixthDay,
            folderMonth: derived.folderMonth,
            termLabel: autoDates.termLabel,
            status,
            reason,
            dateSent,
            startDate,
            endContract
          };

          const idx = appData[activeTab].findIndex(t => t.tenantName === originalName);
          if (idx !== -1) appData[activeTab][idx] = updated;

          if (name !== originalName) {
            await fetch(APPS_SCRIPT_URL, {
              method: 'POST',
              body: JSON.stringify({ action: 'delete', tabName: activeTab, tenantName: originalName })
            });
            await syncToDatabase(updated, activeTab, 'save');
          } else {
            await syncToDatabase(updated, activeTab, 'update');
          }
        }

        closeEditModal();
        renderSpreadsheet();
      } catch (err) {
        console.error(err);
        alert("Error: " + err.message);
      } finally {
        document.getElementById('loadingOverlay').classList.add('hidden');
        document.getElementById('loadingOverlay').classList.remove('flex');
      }
    };

    window.deleteTenant = async function(tenantName) {
      if (!activeTab) return;
      const idx = appData[activeTab].findIndex(t => t.tenantName === tenantName);
      if (idx === -1) return alert("Tenant not found.");
      const tenant = appData[activeTab][idx];

      const confirmDelete = confirm(`Delete tenant "${tenantName}" (Unit ${tenant.unit})? This will remove the row from Google Sheets and cannot be undone.`);
      if (!confirmDelete) return;

      document.getElementById('loadingOverlay').classList.remove('hidden');
      document.getElementById('loadingOverlay').classList.add('flex');
      document.getElementById('loadingText').textContent = `Deleting ${tenantName}...`;

      try {
        const res = await fetch(APPS_SCRIPT_URL, {
          method: 'POST',
          body: JSON.stringify({ action: 'delete', tabName: activeTab, tenantName })
        });
        const resData = await res.json();
        if (resData.success) {
          appData[activeTab].splice(idx, 1);
          renderSpreadsheet();
        } else {
          alert("Failed to delete: " + (resData.error || "Unknown error"));
        }
      } catch (e) {
        console.error(e);
        alert("Error deleting tenant: " + e.message);
      } finally {
        document.getElementById('loadingOverlay').classList.add('hidden');
        document.getElementById('loadingOverlay').classList.remove('flex');
      }
    };

    document.getElementById('searchInput').addEventListener('input', renderSpreadsheet);
    document.getElementById('statusFilter').addEventListener('change', renderSpreadsheet);
    document.getElementById('monthFilter').addEventListener('change', renderSpreadsheet);

    window.updateDateSent = async function(tenantName, newDateStr) {
      const tenantIndex = appData[activeTab].findIndex(t => t.tenantName === tenantName);
      if (tenantIndex === -1) return;
      const tenant = appData[activeTab][tenantIndex];

      const normalizedDate = newDateStr.trim() ? toDateInputValue(newDateStr) : '';
      if (newDateStr.trim() && !normalizedDate) {
        alert("Enter the date as M/D/YYYY, for example 7/22/2026.");
        renderSpreadsheet();
        return;
      }
      tenant.dateSent = normalizedDate;
      renderSpreadsheet();
      await syncToDatabase(tenant, activeTab, 'update');
    };

    window.updateComputed20th = async function(tenantName, newDateStr) {
      const tenantIndex = appData[activeTab].findIndex(t => t.tenantName === tenantName);
      if (tenantIndex === -1) return;
      const tenant = appData[activeTab][tenantIndex];

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

      const derived = deriveFromComputed20th(normalizedDate);
      tenant.computed20th = normalizedDate;
      tenant.sixthDay = derived.sixthDay;
      tenant.folderMonth = derived.folderMonth;

      renderSpreadsheet();
      await syncToDatabase(tenant, activeTab, 'update');
    };

    async function generatePDFDoc(tenant) {
      const safeSetText = (id, val) => {
        const el = document.getElementById(id);
        if (el) el.textContent = val;
      };

      const shortTerm = isShortTerm(tenant.termLabel);
      const dayCountStr = shortTerm ? '10' : '20';
      const daySuffixStr = shortTerm ? '10th' : '20th';
      const titleText = shortTerm ? '10‑DAY NOTICE & ADVANCE AGREEMENT FOR RENEWAL OF CONTRACT' : '20‑DAY NOTICE & ADVANCE AGREEMENT FOR RENEWAL OF CONTRACT';

      try {
        const titleEl = document.getElementById('pdfTitleHeader');
        if (titleEl) titleEl.textContent = titleText;

        safeSetText('outNoticeDaysLabel', dayCountStr);
        safeSetText('outTermDays1', daySuffixStr);
        safeSetText('outTermDays2', daySuffixStr);

        safeSetText('outDate', formatDisplayDate(tenant.date) || tenant.date || '');
        safeSetText('outTenant', tenant.tenantName || '');
        safeSetText('outBuilding', tenant.building || '');
        safeSetText('outUnit', tenant.unit || '');
        
        const pdfDateStr = formatDisplayDate(tenant.computed20th) || tenant.computed20th;
        const contractEndDateStr = formatDisplayDate(tenant.endContract) || tenant.endContract || '';
        const propertyName = String(tenant.building || activeTab || '').toLowerCase().trim().replace(/\s+/g, ' ');
        const isBedAndBath = propertyName === 'bed and bath';
        safeSetText('outOccupancySupportFee', isBedAndBath ? '' : ' The tenant shall also pay an Occupancy Support Fee of ₱298.00 per month.');

        safeSetText('outNoticeDeadline', pdfDateStr || '');
        safeSetText('outEndContract', contractEndDateStr);
        safeSetText('outDay20b', pdfDateStr || '');
        safeSetText('outEndContract2', contractEndDateStr);
        safeSetText('outSixthDay', formatDisplayDate(tenant.sixthDay) || tenant.sixthDay || '');
        safeSetText('outCheckYes', (tenant.choice && tenant.choice.toUpperCase() === 'YES') ? '[ X ]' : '[   ]');
        safeSetText('outCheckNo', (tenant.choice && tenant.choice.toUpperCase() === 'NO') ? '[ X ]' : '[   ]');
        safeSetText('outReason', tenant.reason || '—');
        safeSetText('outSignature1', tenant.tenantName || '');
        
        safeSetText('outTenant2', tenant.tenantName || '');
        safeSetText('outTermLabel', tenant.termLabel || '');
        safeSetText('outRent', formatPeso(tenant.rent));
        safeSetText('outRent2', formatPeso(tenant.rent));
        safeSetText('outSignature2', tenant.tenantName || '');
      } catch (err) {
        throw new Error("Failed to populate template: " + err.message);
      }

      try {
        if (typeof html2canvas === 'undefined' || typeof window.jspdf === 'undefined') {
          throw new Error("Required libraries not loaded.");
        }

        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('p', 'mm', 'a4');
        const page1 = document.getElementById('page-1');
        const page2 = document.getElementById('page-2');

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

    window.generateSingle = async function(tenantName) {
      const tenantIndex = appData[activeTab].findIndex(t => t.tenantName === tenantName);
      if (tenantIndex === -1) return;
      const tenant = appData[activeTab][tenantIndex];
      
      document.getElementById('loadingOverlay').classList.remove('hidden');
      document.getElementById('loadingOverlay').classList.add('flex');
      document.getElementById('loadingText').textContent = `Generating PDF for ${tenant.tenantName}...`;

      try {
        const doc = await generatePDFDoc(tenant);
        const fileName = `Lease Renewal Notice - ${tenant.tenantName}.pdf`;

        if (APPS_SCRIPT_URL !== 'YOUR_APPS_SCRIPT_WEB_APP_URL_HERE') {
          document.getElementById('loadingText').textContent = `Uploading to Drive...`;
          
          let monthName = tenant.folderMonth;
          if (!monthName || monthName === 'Unknown Month' || monthName.trim() === '') {
            if (tenant.computed20th && tenant.computed20th.includes('-')) {
              monthName = new Date(tenant.computed20th).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
            } else {
              monthName = new Date().toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
            }
          }

          const propertyName = tenant.building || "Unknown Property";
          const base64Data = doc.output('datauristring').split(',')[1];
          
          const uploadRes = await fetch(APPS_SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({
              fileName: fileName, base64: base64Data, rootFolderId: ROOT_FOLDER_ID,
              propertyName: propertyName, monthName: monthName, tenantName: tenant.tenantName
            })
          });

          const uploadData = await uploadRes.json();
          if (uploadData.success && uploadData.fileUrl) {
            tenant.pdfUrl = uploadData.fileUrl;
          } else {
            throw new Error("Upload failed or no URL returned.");
          }
        } else {
          doc.save(fileName);
        }

        tenant.status = 'Generated';
        renderSpreadsheet(); 
        await syncToDatabase(tenant, activeTab, 'update');
      } catch (error) {
        console.error(error);
        alert("Error generating PDF: " + error.message);
      } finally {
        document.getElementById('loadingOverlay').classList.add('hidden');
        document.getElementById('loadingOverlay').classList.remove('flex');
      }
    };

    window.regenerateSingle = async function(tenantName) {
      const tenantIndex = appData[activeTab].findIndex(t => t.tenantName === tenantName);
      if (tenantIndex === -1) return;
      const tenant = appData[activeTab][tenantIndex];
      
      document.getElementById('loadingOverlay').classList.remove('hidden');
      document.getElementById('loadingOverlay').classList.add('flex');
      document.getElementById('loadingText').textContent = `Regenerating PDF for ${tenant.tenantName}...`;

      try {
        const doc = await generatePDFDoc(tenant);
        const fileName = `Lease Renewal Notice - ${tenant.tenantName}.pdf`;

        if (APPS_SCRIPT_URL !== 'YOUR_APPS_SCRIPT_WEB_APP_URL_HERE') {
          document.getElementById('loadingText').textContent = `Re-uploading to Drive...`;
          
          let monthName = tenant.folderMonth;
          if (!monthName || monthName === 'Unknown Month' || monthName.trim() === '') {
            if (tenant.computed20th && tenant.computed20th.includes('-')) {
              monthName = new Date(tenant.computed20th).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
            } else {
              monthName = new Date().toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
            }
          }

          const propertyName = tenant.building || "Unknown Property";
          const base64Data = doc.output('datauristring').split(',')[1];
          
          const uploadRes = await fetch(APPS_SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({
              fileName: fileName, base64: base64Data, rootFolderId: ROOT_FOLDER_ID,
              propertyName: propertyName, monthName: monthName, tenantName: tenant.tenantName
            })
          });

          const uploadData = await uploadRes.json();
          if (uploadData.success && uploadData.fileUrl) {
            tenant.pdfUrl = uploadData.fileUrl;
          } else {
            throw new Error("Upload failed or no URL returned.");
          }
        } else {
          doc.save(fileName);
        }
        
        renderSpreadsheet();
        await syncToDatabase(tenant, activeTab, 'update');
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
      const confirmGenerate = confirm(`Are you sure you want to generate PDFs for ${selectionLabel === 'selected' ? '' : 'all '}${pendingTenants.length} ${scopeText} tenants in "${activeTab}"?`);
      if (!confirmGenerate) return;
      document.getElementById('loadingOverlay').classList.remove('hidden');
      document.getElementById('loadingOverlay').classList.add('flex');

      try {
        for (let i = 0; i < pendingTenants.length; i++) {
          const tenant = pendingTenants[i];
          document.getElementById('loadingText').textContent = `Generating (${i + 1}/${pendingTenants.length}): ${tenant.tenantName}...`;
          
          const doc = await generatePDFDoc(tenant);
          const fileName = `Lease Renewal - ${tenant.tenantName}.pdf`;

          if (APPS_SCRIPT_URL !== 'YOUR_APPS_SCRIPT_WEB_APP_URL_HERE') {
            document.getElementById('loadingText').textContent = `Uploading (${i + 1}/${pendingTenants.length}): ${tenant.tenantName}...`;
            
            let monthName = tenant.folderMonth;
            if (!monthName || monthName === 'Unknown Month' || monthName.trim() === '') {
              if (tenant.computed20th && tenant.computed20th.includes('-')) {
                monthName = new Date(tenant.computed20th).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
              } else {
                monthName = new Date().toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
              }
            }
            
            const propertyName = tenant.building || "Unknown Property";
            const base64Data = doc.output('datauristring').split(',')[1];
            
            const uploadRes = await fetch(APPS_SCRIPT_URL, {
              method: 'POST',
              body: JSON.stringify({
                fileName: fileName, base64: base64Data, rootFolderId: ROOT_FOLDER_ID,
                propertyName: propertyName, monthName: monthName, tenantName: tenant.tenantName
              })
            });

            const uploadData = await uploadRes.json();
            if (uploadData.success && uploadData.fileUrl) {
              tenant.pdfUrl = uploadData.fileUrl;
            }
          } else {
            doc.save(fileName);
          }

          tenant.status = 'Generated';
          renderSpreadsheet(); 
          await syncToDatabase(tenant, activeTab, 'update');
        }
        
        alert(`Successfully generated and uploaded ${pendingTenants.length} ${selectionLabel === 'selected' ? 'selected ' : ''}notices.`);
      } catch (error) {
        console.error(error);
        alert("An error occurred during batch generation: " + error.message);
      } finally {
        document.getElementById('loadingOverlay').classList.add('hidden');
        document.getElementById('loadingOverlay').classList.remove('flex');
      }
    }

    document.getElementById('batchGenerateBtn').addEventListener('click', () => {
      if (!activeTab) return alert("No active tab selected.");
      const pendingTenants = appData[activeTab].filter(tenant => tenant.status === 'Pending');
      batchGenerateTenants(pendingTenants, 'all');
    });

    document.getElementById('batchGenerateSelectedBtn').addEventListener('click', () => {
      if (!activeTab) return alert("No active tab selected.");
      const pendingTenants = appData[activeTab].filter(tenant => selectedTenants.has(tenant) && tenant.status === 'Pending');
      batchGenerateTenants(pendingTenants, 'selected');
    });

    document.getElementById('batchDownloadZipBtn').addEventListener('click', async () => {
      if (!activeTab) return alert("No active tab selected.");

      const generatedTenants = appData[activeTab].filter(t => selectedTenants.has(t) && t.pdfUrl && String(t.pdfUrl).trim() !== '');
      
      if (generatedTenants.length === 0) {
        return alert("Select one or more tenants with generated PDFs to download.");
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
            const fileName = `Lease Renewal Notice - ${tenant.tenantName}.pdf`;
            
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

    window.markAsSent = async function(tenantName) {
      const tenantIndex = appData[activeTab].findIndex(t => t.tenantName === tenantName);
      if (tenantIndex === -1) return;
      const tenant = appData[activeTab][tenantIndex];
      tenant.status = 'Sent';
      renderSpreadsheet();
      await syncToDatabase(tenant, activeTab, 'update');
    };

    document.getElementById('clearCurrentTabBtn').addEventListener('click', async () => {
      if (!activeTab) return alert("No active tab selected.");
      
      const confirmClear = confirm(`Are you sure you want to clear ALL data in the "${activeTab}" tab? This will wipe Google Sheets data.`);
      
      if (confirmClear) {
        document.getElementById('loadingOverlay').classList.remove('hidden');
        document.getElementById('loadingOverlay').classList.add('flex');
        document.getElementById('loadingText').textContent = `Clearing tab ${activeTab}...`;

        try {
          const res = await fetch(APPS_SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'clearTab', tabName: activeTab })
          });
          
          const resData = await res.json();
          
          if (resData.success) {
            appData[activeTab] = [];
            renderSpreadsheet();
            alert(`Tab "${activeTab}" cleared.`);
          } else {
            alert("Failed to clear: " + (resData.error || "Unknown"));
          }
        } catch (e) {
          alert("Error clearing tab.");
        } finally {
          document.getElementById('loadingOverlay').classList.add('hidden');
          document.getElementById('loadingOverlay').classList.remove('flex');
        }
      }
    });

    document.getElementById('addTenantFormBtn').addEventListener('click', async () => {
      const tenantName = document.getElementById('inputTenant').value;
      if (!tenantName) return alert("Enter Tenant Name");
      if (!activeTab) return alert("Select a tab first.");

      const startContract = document.getElementById('inputStartDate').value;
      const endContract = document.getElementById('inputEndContract').value;
      const term = document.getElementById('inputTerm').value;
      const dates = parseDate(startContract, term);

      const newTenant = {
        tenantName: tenantName,
        building: document.getElementById('inputBuilding').value,
        unit: document.getElementById('inputUnit').value,
        rent: document.getElementById('inputRent').value,
        computed20th: dates.computed20th,
        termLabel: dates.termLabel,
        status: 'Pending',
        pdfUrl: '',
        reason: '—',
        dateSent: '',
        startDate: startContract,
        endContract
      };

      appData[activeTab].push(newTenant);
      renderSpreadsheet();
      document.getElementById('inputTenant').value = '';
      document.getElementById('inputEndContract').value = '';

      await syncToDatabase(newTenant, activeTab, 'save');
    });

    document.getElementById('recalculateDatesBtn').addEventListener('click', async () => {
      if (!activeTab) return alert("No active tab selected.");
      
      const confirmRecalc = confirm(`Are you sure you want to recalculate ALL dates in the "${activeTab}" tab based on the new formula? This will update Google Sheets.`);
      if (!confirmRecalc) return;

      document.getElementById('loadingOverlay').classList.remove('hidden');
      document.getElementById('loadingOverlay').classList.add('flex');
      document.getElementById('loadingText').textContent = `Recalculating dates...`;

      try {
        const tenants = appData[activeTab];
        for (let i = 0; i < tenants.length; i++) {
          const tenant = tenants[i];
          if (tenant.startDate && tenant.termLabel) {
            const dates = parseDate(tenant.startDate, tenant.termLabel);
            tenant.computed20th = dates.computed20th;
            tenant.sixthDay = dates.sixthDay;
            tenant.folderMonth = dates.folderMonth;
            
            await syncToDatabase(tenant, activeTab, 'update');
          }
        }
        
        renderSpreadsheet();
        alert("Dates recalculated successfully!");
      } catch (error) {
        console.error(error);
        alert("Error recalculating dates.");
      } finally {
        document.getElementById('loadingOverlay').classList.add('hidden');
        document.getElementById('loadingOverlay').classList.remove('flex');
      }
    });

    document.getElementById('excelInput').addEventListener('change', (e) => {
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

    document.getElementById('confirmLoadSheetBtn').addEventListener('click', async () => {
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
          document.getElementById('loadingText').textContent = `Uploading Data... (${i + 1}/${json.length})`;
          
          const row = json[i];
          const rowKeys = Object.keys(row).reduce((acc, key) => {
            acc[key.trim().toUpperCase()] = row[key];
            return acc;
          }, {});

          const startContract = rowKeys['START OF CONTRACT'] || '';
          const term = rowKeys['LEASE TERM'] || '12';
          const dates = parseDate(startContract, term);

          const newTenant = {
            tenantName: rowKeys['NAME'] || 'Unknown',
            building: rowKeys['BUILDING'] || "Eve's Residences",
            unit: rowKeys['UNIT'] || '',
            rent: rowKeys['MONTHLY RENT'] || 0,
            computed20th: dates.computed20th,
            termLabel: dates.termLabel,
            status: 'Pending',
            pdfUrl: '',
            reason: rowKeys['REASON'] || '—',
            dateSent: '',
            startDate: startContract
          };

          appData[selectedSheetName].push(newTenant);
          await syncToDatabase(newTenant, selectedSheetName, 'save');
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

    window.onload = initDatabase;