//
//    Timestamp: 2025-05-09T07:48:22EDT
//    Summary: Removed '% Quarterly Revenue Target' from the Attach Rates table display logic (columns array and averageMetrics array).
//
document.addEventListener('DOMContentLoaded', () => {
    // --- Theme Constants and Elements ---
    const LIGHT_THEME_CLASS = 'light-theme';
    const THEME_STORAGE_KEY = 'themePreference';
    const DARK_THEME_ICON = '🌙'; 
    const LIGHT_THEME_ICON = '☀️'; 
    const DARK_THEME_META_COLOR = '#2c2c2c';
    const LIGHT_THEME_META_COLOR = '#f4f4f8'; 

    const themeToggleBtn = document.getElementById('themeToggleBtn');
    const metaThemeColorTag = document.querySelector('meta[name="theme-color"]');

    // --- Configuration ---
    const REQUIRED_HEADERS = [ 
        'Store', 'REGION', 'DISTRICT', 'Q2 Territory', 'FSM NAME', 'CHANNEL',
        'SUB_CHANNEL', 'DEALER_NAME', 'Revenue w/DF', 'QTD Revenue Target',
        'Quarterly Revenue Target', 'QTD Gap', '% Quarterly Revenue Target', 'Rev AR%', 
        'Unit w/ DF', 'Unit Target', 'Unit Achievement', 'Visit count', 'Trainings',
        'Retail Mode Connectivity', 'Rep Skill Ach', '(V)PMR Ach', 'Elite', 'Post Training Score',
        'Tablet Attach Rate', 'PC Attach Rate', 'NC Attach Rate', 'TWS Attach Rate',
        'WW Attach Rate', 'ME Attach Rate', 'NCME Attach Rate', 'SUPER STORE', 'GOLDEN RHINO',
        'GCE', 'AI_Zone', 'Hispanic_Market', 'EV ROUTE',
        'STORE ID', 'ADDRESS1', 'CITY', 'STATE', 'ZIPCODE',
        'LATITUDE_ORG', 'LONGITUDE_ORG', 
        'ORG_STORE_ID', 'CV_STORE_ID', 'CINGLEPOINT_ID', 
        'STORE_TYPE_NAME', 'National_Tier', 'Merchandising_Level', 'Combined_Tier', 
        '%Quarterly Territory Rev Target', 'Region Rev%', 'District Rev%', 'Territory Rev%'
    ]; // '% Quarterly Revenue Target' is still required for other parts of the dashboard
    const FLAG_HEADERS = ['SUPER STORE', 'GOLDEN RHINO', 'GCE', 'AI_Zone', 'Hispanic_Market', 'EV ROUTE'];
    const CURRENCY_FORMAT = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' });
    const PERCENT_FORMAT = new Intl.NumberFormat('en-US', { style: 'percent', minimumFractionDigits: 1, maximumFractionDigits: 1 });
    const NUMBER_FORMAT = new Intl.NumberFormat('en-US');
    const CHART_COLORS = ['#58a6ff', '#ffb758', '#86dc86', '#ff7f7f', '#b796e6', '#ffda8a', '#8ad7ff', '#ff9ba6'];
    const TOP_N_CHART = 15; 
    const TOP_N_TABLES = 5;

    // --- DOM Elements ---
    const excelFileInput = document.getElementById('excelFile');
    const statusDiv = document.getElementById('status');
    const loadingIndicator = document.getElementById('loadingIndicator');
    const filterLoadingIndicator = document.getElementById('filterLoadingIndicator');
    const filterArea = document.getElementById('filterArea');
    const resultsArea = document.getElementById('resultsArea');
    const applyFiltersButton = document.getElementById('applyFiltersButton');
    const resetFiltersButton = document.getElementById('resetFiltersButton'); 
    const regionFilter = document.getElementById('regionFilter');
    const districtFilter = document.getElementById('districtFilter');
    const territoryFilter = document.getElementById('territoryFilter');
    const fsmFilter = document.getElementById('fsmFilter');
    const channelFilter = document.getElementById('channelFilter');
    const subchannelFilter = document.getElementById('subchannelFilter');
    const dealerFilter = document.getElementById('dealerFilter');
    const storeFilter = document.getElementById('storeFilter');
    const storeSearch = document.getElementById('storeSearch');
    const flagFiltersCheckboxes = FLAG_HEADERS.reduce((acc, header) => {
        let expectedId = '';
        switch (header) {
            case 'SUPER STORE':       expectedId = 'superStoreFilter'; break;
            case 'GOLDEN RHINO':      expectedId = 'goldenRhinoFilter'; break;
            case 'GCE':               expectedId = 'gceFilter'; break;
            case 'AI_Zone':           expectedId = 'aiZoneFilter'; break;
            case 'Hispanic_Market':   expectedId = 'hispanicMarketFilter'; break;
            case 'EV ROUTE':          expectedId = 'evRouteFilter'; break;
            default: console.warn(`Unknown flag header encountered during mapping: ${header}`); return acc;
        }
        const element = document.getElementById(expectedId);
        if (element) { acc[header] = element; } 
        else { console.warn(`Flag filter checkbox not found for ID: ${expectedId} (Header: ${header}) upon initial mapping. Check HTML.`);}
        return acc;
    }, {});
    const territorySelectAll = document.getElementById('territorySelectAll');
    const territoryDeselectAll = document.getElementById('territoryDeselectAll');
    const storeSelectAll = document.getElementById('storeSelectAll');
    const storeDeselectAll = document.getElementById('storeDeselectAll');
    const revenueWithDFValue = document.getElementById('revenueWithDFValue');
    const qtdRevenueTargetValue = document.getElementById('qtdRevenueTargetValue');
    const qtdGapValue = document.getElementById('qtdGapValue');
    const quarterlyRevenueTargetValue = document.getElementById('quarterlyRevenueTargetValue');
    const percentQuarterlyStoreTargetValue = document.getElementById('percentQuarterlyStoreTargetValue');
    const revARValue = document.getElementById('revARValue');
    const unitsWithDFValue = document.getElementById('unitsWithDFValue');
    const unitTargetValue = document.getElementById('unitTargetValue');
    const unitAchievementValue = document.getElementById('unitAchievementValue');
    const visitCountValue = document.getElementById('visitCountValue');
    const trainingCountValue = document.getElementById('trainingCountValue');
    const retailModeConnectivityValue = document.getElementById('retailModeConnectivityValue');
    const repSkillAchValue = document.getElementById('repSkillAchValue');
    const vPmrAchValue = document.getElementById('vPmrAchValue');
    const postTrainingScoreValue = document.getElementById('postTrainingScoreValue');
    const eliteValue = document.getElementById('eliteValue');
    const percentQuarterlyTerritoryTargetP = document.getElementById('percentQuarterlyTerritoryTargetP');
    const territoryRevPercentP = document.getElementById('territoryRevPercentP');
    const districtRevPercentP = document.getElementById('districtRevPercentP');
    const regionRevPercentP = document.getElementById('regionRevPercentP');
    const percentQuarterlyTerritoryTargetValue = document.getElementById('percentQuarterlyTerritoryTargetValue');
    const territoryRevPercentValue = document.getElementById('territoryRevPercentValue');
    const districtRevPercentValue = document.getElementById('districtRevPercentValue');
    const regionRevPercentValue = document.getElementById('regionRevPercentValue');
    const attachRateTableBody = document.getElementById('attachRateTableBody');
    const attachRateTableFooter = document.getElementById('attachRateTableFooter');
    const attachTableStatus = document.getElementById('attachTableStatus');
    const attachRateTable = document.getElementById('attachRateTable');
    const exportCsvButton = document.getElementById('exportCsvButton');
    const topBottomSection = document.getElementById('topBottomSection');
    const top5TableBody = document.getElementById('top5TableBody');
    const bottom5TableBody = document.getElementById('bottom5TableBody');
    const mainChartCanvas = document.getElementById('mainChartCanvas')?.getContext('2d');
    const storeDetailsSection = document.getElementById('storeDetailsSection');
    const storeDetailsContent = document.getElementById('storeDetailsContent');
    const closeStoreDetailsButton = document.getElementById('closeStoreDetailsButton');
    const emailRecipientInput = document.getElementById('emailRecipient');
    const shareEmailButton = document.getElementById('shareEmailButton');
    const shareStatus = document.getElementById('shareStatus');

    // --- Global State ---
    let rawData = [];
    let filteredData = [];
    let mainChartInstance = null;
    let storeOptions = [];
    let allPossibleStores = [];
    let currentSort = { column: 'Store', ascending: true };
    let selectedStoreRow = null;

    // --- Theme Management ---
    const applyTheme = (theme) => {
        if (theme === 'light') {
            document.body.classList.add(LIGHT_THEME_CLASS);
            if (themeToggleBtn) themeToggleBtn.textContent = DARK_THEME_ICON;
            if (metaThemeColorTag) metaThemeColorTag.setAttribute('content', LIGHT_THEME_META_COLOR);
        } else { 
            document.body.classList.remove(LIGHT_THEME_CLASS);
            if (themeToggleBtn) themeToggleBtn.textContent = LIGHT_THEME_ICON;
            if (metaThemeColorTag) metaThemeColorTag.setAttribute('content', DARK_THEME_META_COLOR);
        }
        if (mainChartInstance && (filteredData.length > 0 || (rawData.length > 0 && filteredData.length === 0) )) { 
             updateCharts(filteredData); 
        }
    };

    const toggleTheme = () => {
        const isLight = document.body.classList.contains(LIGHT_THEME_CLASS);
        const newTheme = isLight ? 'dark' : 'light';
        applyTheme(newTheme);
        localStorage.setItem(THEME_STORAGE_KEY, newTheme);
    };
    
    const getChartThemeColors = () => {
        const isLight = document.body.classList.contains(LIGHT_THEME_CLASS);
        return {
            tickColor: isLight ? '#495057' : '#e0e0e0',
            gridColor: isLight ? 'rgba(0, 0, 0, 0.1)' : 'rgba(224, 224, 224, 0.2)',
            legendColor: isLight ? '#333333' : '#e0e0e0'
        };
    };

    // --- Helper Functions ---
    const formatCurrency = (value) => isNaN(value) ? 'N/A' : CURRENCY_FORMAT.format(value);
    const formatPercent = (value) => isNaN(value) ? 'N/A' : PERCENT_FORMAT.format(value);
    const formatNumber = (value) => isNaN(value) ? 'N/A' : NUMBER_FORMAT.format(value);
    const parseNumber = (value) => {
        if (value === null || value === undefined || String(value).trim() === '') return NaN;
        if (typeof value === 'number') return value;
        if (typeof value === 'string') { const numStr = value.replace(/[$,%]/g, ''); const num = parseFloat(numStr); return isNaN(num) ? NaN : num; }
        return NaN;
    };
    const parsePercent = (value) => {
         if (value === null || value === undefined || String(value).trim() === '') return NaN;
         if (typeof value === 'number') return value;
         if (typeof value === 'string') { const numStr = value.replace('%', ''); const num = parseFloat(numStr); return isNaN(num) ? NaN : num / 100; }
         return NaN;
    };
    const safeGet = (obj, path, defaultValue = 'N/A') => {
        const value = obj ? obj[path] : undefined;
        return (value !== undefined && value !== null && String(value).trim() !== '') ? value : defaultValue;
    };
    const isValidForAverage = (value) => {
         if (value === null || value === undefined || String(value).trim() === '') return false;
         return !isNaN(parseNumber(String(value).replace('%','')));
    };
    const calculateQtdGap = (row) => {
        const revenue = parseNumber(safeGet(row, 'Revenue w/DF', 0)); 
        const target = parseNumber(safeGet(row, 'QTD Revenue Target', 0));
        if (isNaN(revenue) || isNaN(target)) { return Infinity; } 
        return revenue - target;
    };
    const calculateRevARPercentForRow = (row) => { 
        const revenue = parseNumber(safeGet(row, 'Revenue w/DF', 0));
        const target = parseNumber(safeGet(row, 'QTD Revenue Target', 0));
        if (target === 0 || isNaN(revenue) || isNaN(target)) { return NaN; } 
        return revenue / target;
    };
    const calculateUnitAchievementPercentForRow = (row) => { 
        const units = parseNumber(safeGet(row, 'Unit w/ DF', 0));
        const target = parseNumber(safeGet(row, 'Unit Target', 0));
        if (target === 0 || isNaN(units) || isNaN(target)) { return NaN; } 
        return units / target;
    };
    const getUniqueValues = (data, column) => {
        const values = new Set(data.map(item => safeGet(item, column, '')).filter(val => String(val).trim() !== ''));
        return ['ALL', ...Array.from(values).sort()];
    };
    const setOptions = (selectElement, options, disable = false) => {
        if (!selectElement) return;
        selectElement.innerHTML = '';
        options.forEach(optionValue => { const option = document.createElement('option'); option.value = optionValue; option.textContent = optionValue === 'ALL' ? `-- ${optionValue} --` : optionValue; option.title = optionValue; selectElement.appendChild(option); });
        selectElement.disabled = disable;
    };
    const setMultiSelectOptions = (selectElement, options, disable = false) => {
        if (!selectElement) return;
        selectElement.innerHTML = '';
        options.forEach(optionValue => { if (optionValue === 'ALL') return; const option = document.createElement('option'); option.value = optionValue; option.textContent = optionValue; option.title = optionValue; selectElement.appendChild(option); });
        selectElement.disabled = disable;
    };
    const showLoading = (isLoading, isFiltering = false) => {
        const displayStyle = isLoading ? 'flex' : 'none';
        if (isFiltering) { if (filterLoadingIndicator) filterLoadingIndicator.style.display = displayStyle; if (applyFiltersButton) applyFiltersButton.disabled = isLoading; } 
        else { if (loadingIndicator) loadingIndicator.style.display = displayStyle; if (excelFileInput) excelFileInput.disabled = isLoading; }
    };    

    // --- Core Functions ---
    const handleFile = async (event) => {
        const file = event.target.files[0];
        if (!file) { if (statusDiv) statusDiv.textContent = 'No file selected.'; return; }
        if (statusDiv) statusDiv.textContent = 'Reading file...';
        showLoading(true); if (filterArea) filterArea.style.display = 'none'; if (resultsArea) resultsArea.style.display = 'none'; resetUI(); 
        try {
            const data = await file.arrayBuffer(); const workbook = XLSX.read(data); const firstSheetName = workbook.SheetNames[0]; const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: null });
            if (jsonData.length > 0) { const headers = Object.keys(jsonData[0]); const missingHeaders = REQUIRED_HEADERS.filter(h => !headers.includes(h)); if (missingHeaders.length > 0) { console.warn(`Warning: Missing expected columns: ${missingHeaders.join(', ')}.`); }
            } else { throw new Error("Excel sheet appears to be empty."); }
            rawData = jsonData; allPossibleStores = [...new Set(rawData.map(r => safeGet(r, 'Store', null)).filter(s => s && String(s).trim() !== ''))].sort().map(s => ({ value: s, text: s }));
            if (statusDiv) statusDiv.textContent = `Loaded ${rawData.length} rows. Adjust filters and click 'Apply Filters'.`;
            populateFilters(rawData); if (filterArea) filterArea.style.display = 'block';
        } catch (error) {
            console.error('Error processing file:', error); if (statusDiv) statusDiv.textContent = `Error: ${error.message}`;
            rawData = []; allPossibleStores = []; filteredData = []; resetUI();
        } finally { showLoading(false); if (excelFileInput) excelFileInput.value = ''; }
    };
    const populateFilters = (data) => {
        setOptions(regionFilter, getUniqueValues(data, 'REGION')); setOptions(districtFilter, getUniqueValues(data, 'DISTRICT')); setMultiSelectOptions(territoryFilter, getUniqueValues(data, 'Q2 Territory').slice(1));
        setOptions(fsmFilter, getUniqueValues(data, 'FSM NAME')); setOptions(channelFilter, getUniqueValues(data, 'CHANNEL')); setOptions(subchannelFilter, getUniqueValues(data, 'SUB_CHANNEL')); setOptions(dealerFilter, getUniqueValues(data, 'DEALER_NAME'));
        Object.values(flagFiltersCheckboxes).forEach(input => { if(input) input.disabled = false; });
        storeOptions = [...allPossibleStores]; setStoreFilterOptions(storeOptions, false);
        if (territorySelectAll) territorySelectAll.disabled = false; if (territoryDeselectAll) territoryDeselectAll.disabled = false;
        if (storeSelectAll) storeSelectAll.disabled = false; if (storeDeselectAll) storeDeselectAll.disabled = false;
        if (storeSearch) storeSearch.disabled = false; 
        if (applyFiltersButton) applyFiltersButton.disabled = false;
        if (resetFiltersButton) resetFiltersButton.disabled = false; 
        addDependencyFilterListeners();
    };
    const addDependencyFilterListeners = () => {
        const handler = updateStoreFilterOptionsBasedOnHierarchy; const filtersToListen = [regionFilter, districtFilter, territoryFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter];
        filtersToListen.forEach(filter => { if (filter) { filter.removeEventListener('change', handler); filter.addEventListener('change', handler); } });
        Object.values(flagFiltersCheckboxes).forEach(input => { if (input) { input.removeEventListener('change', handler); input.addEventListener('change', handler); } });
    };
    const updateStoreFilterOptionsBasedOnHierarchy = () => {
        if (rawData.length === 0) return;
        const selectedRegion = regionFilter?.value; const selectedDistrict = districtFilter?.value; const selectedTerritories = territoryFilter ? Array.from(territoryFilter.selectedOptions).map(opt => opt.value) : [];
        const selectedFsm = fsmFilter?.value; const selectedChannel = channelFilter?.value; const selectedSubchannel = subchannelFilter?.value; const selectedDealer = dealerFilter?.value;
        const selectedFlags = {}; Object.entries(flagFiltersCheckboxes).forEach(([key, input]) => { if (input?.checked) { selectedFlags[key] = true; } });
        const potentiallyValidStoresData = rawData.filter(row => {
            if (selectedRegion !== 'ALL' && safeGet(row, 'REGION', null) !== selectedRegion) return false; if (selectedDistrict !== 'ALL' && safeGet(row, 'DISTRICT', null) !== selectedDistrict) return false;
            if (selectedTerritories.length > 0 && !selectedTerritories.includes(safeGet(row, 'Q2 Territory', null))) return false; if (selectedFsm !== 'ALL' && safeGet(row, 'FSM NAME', null) !== selectedFsm) return false;
            if (selectedChannel !== 'ALL' && safeGet(row, 'CHANNEL', null) !== selectedChannel) return false; if (selectedSubchannel !== 'ALL' && safeGet(row, 'SUB_CHANNEL', null) !== selectedSubchannel) return false;
            if (selectedDealer !== 'ALL' && safeGet(row, 'DEALER_NAME', null) !== selectedDealer) return false;
            for (const flag in selectedFlags) { const flagValue = safeGet(row, flag, 'NO'); if (!(flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1')) { return false; } }
            return true;
        });
        const validStoreNames = new Set(potentiallyValidStoresData.map(row => safeGet(row, 'Store', null)).filter(s => s && String(s).trim() !== ''));
        storeOptions = Array.from(validStoreNames).sort().map(s => ({ value: s, text: s }));
        const previouslySelectedStores = storeFilter ? new Set(Array.from(storeFilter.selectedOptions).map(opt => opt.value)) : new Set();
        setStoreFilterOptions(storeOptions, false); filterStoreOptions(); 
        if (storeFilter) { Array.from(storeFilter.options).forEach(option => { if (previouslySelectedStores.has(option.value)) { option.selected = true; } }); if (storeFilter.selectedOptions.length === 0) { storeFilter.selectedIndex = -1; } }
    };
    const setStoreFilterOptions = (optionsToShow, disable = true) => {
        if (!storeFilter) return; const currentSearchTerm = storeSearch?.value || ''; storeFilter.innerHTML = '';
        optionsToShow.forEach(opt => { const option = document.createElement('option'); option.value = opt.value; option.textContent = opt.text; option.title = opt.text; storeFilter.appendChild(option); });
        storeFilter.disabled = disable; if (storeSearch) storeSearch.disabled = disable;
        if (storeSelectAll) storeSelectAll.disabled = disable || optionsToShow.length === 0; if (storeDeselectAll) storeDeselectAll.disabled = disable || optionsToShow.length === 0;
        if (storeSearch) storeSearch.value = currentSearchTerm;
    };
    const filterStoreOptions = () => {
        if (!storeFilter || !storeSearch) return; const searchTerm = storeSearch.value.toLowerCase();
        const filteredOptions = storeOptions.filter(opt => opt.text.toLowerCase().includes(searchTerm)); const selectedValues = new Set(Array.from(storeFilter.selectedOptions).map(opt => opt.value));
        storeFilter.innerHTML = ''; filteredOptions.forEach(opt => { const option = document.createElement('option'); option.value = opt.value; option.textContent = opt.text; option.title = opt.text; if (selectedValues.has(opt.value)) { option.selected = true; } storeFilter.appendChild(option); });
        if (storeSelectAll) storeSelectAll.disabled = storeFilter.disabled || filteredOptions.length === 0; if (storeDeselectAll) storeDeselectAll.disabled = storeFilter.disabled || filteredOptions.length === 0;
    };

    const updateTopBottomTables = (data) => {
        if (!topBottomSection || !top5TableBody || !bottom5TableBody) { console.warn("Top/Bottom table elements not found."); return; }
        top5TableBody.innerHTML = ''; bottom5TableBody.innerHTML = '';
        const territoriesInData = new Set(data.map(row => safeGet(row, 'Q2 Territory', null)).filter(Boolean));
        const isSingleTerritorySelected = territoriesInData.size === 1;
        if (!isSingleTerritorySelected || data.length === 0) { topBottomSection.style.display = 'none'; return; }
        topBottomSection.style.display = 'flex'; 
        const top5Data = [...data].sort((a, b) => parseNumber(safeGet(b, 'Revenue w/DF', -Infinity)) - parseNumber(safeGet(a, 'Revenue w/DF', -Infinity))).slice(0, TOP_N_TABLES);
        top5Data.forEach(row => {
            const tr = top5TableBody.insertRow(); const storeName = safeGet(row, 'Store', 'N/A'); tr.dataset.storeName = storeName;
            tr.onclick = () => { showStoreDetails(row); highlightTableRow(storeName); };
            const revenue = parseNumber(safeGet(row, 'Revenue w/DF', NaN)); const revAR = calculateRevARPercentForRow(row); const unitAch = calculateUnitAchievementPercentForRow(row); const visits = parseNumber(safeGet(row, 'Visit count', NaN));
            tr.insertCell().textContent = storeName; tr.cells[0].title = storeName; tr.insertCell().textContent = formatCurrency(revenue); tr.cells[1].title = formatCurrency(revenue);
            tr.insertCell().textContent = formatPercent(revAR); tr.cells[2].title = formatPercent(revAR); tr.insertCell().textContent = formatPercent(unitAch); tr.cells[3].title = formatPercent(unitAch);
            tr.insertCell().textContent = formatNumber(visits); tr.cells[4].title = formatNumber(visits);
        });
        const bottom5Data = [...data].sort((a, b) => calculateQtdGap(a) - calculateQtdGap(b)).slice(0, TOP_N_TABLES);
        bottom5Data.forEach(row => {
            const tr = bottom5TableBody.insertRow(); const storeName = safeGet(row, 'Store', 'N/A'); tr.dataset.storeName = storeName;
            tr.onclick = () => { showStoreDetails(row); highlightTableRow(storeName); };
            const qtdGap = calculateQtdGap(row); const revAR = calculateRevARPercentForRow(row); const unitAch = calculateUnitAchievementPercentForRow(row); const visits = parseNumber(safeGet(row, 'Visit count', NaN));
            tr.insertCell().textContent = storeName; tr.cells[0].title = storeName; tr.insertCell().textContent = formatCurrency(qtdGap === Infinity ? NaN : qtdGap); tr.cells[1].title = formatCurrency(qtdGap === Infinity ? NaN : qtdGap);
            tr.insertCell().textContent = formatPercent(revAR); tr.cells[2].title = formatPercent(revAR); tr.insertCell().textContent = formatPercent(unitAch); tr.cells[3].title = formatPercent(unitAch);
            tr.insertCell().textContent = formatNumber(visits); tr.cells[4].title = formatNumber(visits);
        });
    };

    const applyFilters = () => {
        showLoading(true, true); if (resultsArea) resultsArea.style.display = 'none';
        setTimeout(() => {
            try {
                const selectedRegion = regionFilter?.value; const selectedDistrict = districtFilter?.value; const selectedTerritories = territoryFilter ? Array.from(territoryFilter.selectedOptions).map(opt => opt.value) : [];
                const selectedFsm = fsmFilter?.value; const selectedChannel = channelFilter?.value; const selectedSubchannel = subchannelFilter?.value; const selectedDealer = dealerFilter?.value;
                const selectedStores = storeFilter ? Array.from(storeFilter.selectedOptions).map(opt => opt.value) : [];
                const selectedFlags = {}; Object.entries(flagFiltersCheckboxes).forEach(([key, input]) => { if (input?.checked) { selectedFlags[key] = true; } });
                filteredData = rawData.filter(row => {
                    if (selectedRegion !== 'ALL' && safeGet(row, 'REGION', null) !== selectedRegion) return false; if (selectedDistrict !== 'ALL' && safeGet(row, 'DISTRICT', null) !== selectedDistrict) return false;
                    if (selectedTerritories.length > 0 && !selectedTerritories.includes(safeGet(row, 'Q2 Territory', null))) return false; if (selectedFsm !== 'ALL' && safeGet(row, 'FSM NAME', null) !== selectedFsm) return false;
                    if (selectedChannel !== 'ALL' && safeGet(row, 'CHANNEL', null) !== selectedChannel) return false; if (selectedSubchannel !== 'ALL' && safeGet(row, 'SUB_CHANNEL', null) !== selectedSubchannel) return false;
                    if (selectedDealer !== 'ALL' && safeGet(row, 'DEALER_NAME', null) !== selectedDealer) return false; if (selectedStores.length > 0 && !selectedStores.includes(safeGet(row, 'Store', null))) return false;
                    for (const flag in selectedFlags) { const flagValue = safeGet(row, flag, 'NO'); if (!(flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1')) { return false; } }
                    return true;
                });
                updateSummary(filteredData); updateTopBottomTables(filteredData); updateCharts(filteredData); updateAttachRateTable(filteredData); 
                if (filteredData.length === 1) { showStoreDetails(filteredData[0]); highlightTableRow(safeGet(filteredData[0], 'Store', null)); } else { hideStoreDetails(); }
                if (statusDiv) statusDiv.textContent = `Displaying ${filteredData.length} of ${rawData.length} rows based on filters.`;
                if (resultsArea) resultsArea.style.display = 'block'; if (exportCsvButton) exportCsvButton.disabled = filteredData.length === 0;
            } catch (error) {
                console.error("Error applying filters:", error); if (statusDiv) statusDiv.textContent = "Error applying filters. Check console for details.";
                filteredData = []; if (resultsArea) resultsArea.style.display = 'none'; if (exportCsvButton) exportCsvButton.disabled = true;
                updateSummary([]); updateTopBottomTables([]); updateCharts([]); updateAttachRateTable([]); hideStoreDetails();
            } finally { showLoading(false, true); }
        }, 10);
    };
    
    const resetFiltersForFullUIReset = () => {
         const allOptionHTML = '<option value="ALL">-- Load File First --</option>';
         [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(sel => { 
             if (sel) { sel.innerHTML = allOptionHTML; sel.value = 'ALL'; sel.disabled = true;}
         });
         if (territoryFilter) { territoryFilter.innerHTML = '<option value="ALL">-- Load File First --</option>'; territoryFilter.selectedIndex = -1; territoryFilter.disabled = true; }
         if (storeFilter) { storeFilter.innerHTML = '<option value="ALL">-- Load File First --</option>'; storeFilter.selectedIndex = -1; storeFilter.disabled = true; }
         if (storeSearch) { storeSearch.value = ''; storeSearch.disabled = true; }
         storeOptions = []; 
         Object.values(flagFiltersCheckboxes).forEach(input => { if(input) {input.checked = false; input.disabled = true;} });
         if (applyFiltersButton) applyFiltersButton.disabled = true;
         if (resetFiltersButton) resetFiltersButton.disabled = true; 
         if (territorySelectAll) territorySelectAll.disabled = true; if (territoryDeselectAll) territoryDeselectAll.disabled = true;
         if (storeSelectAll) storeSelectAll.disabled = true; if (storeDeselectAll) storeDeselectAll.disabled = true;
         if (exportCsvButton) exportCsvButton.disabled = true;
         const handler = updateStoreFilterOptionsBasedOnHierarchy;
         [regionFilter, districtFilter, territoryFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(filter => { if (filter) filter.removeEventListener('change', handler); });
         Object.values(flagFiltersCheckboxes).forEach(input => { if (input) input.removeEventListener('change', handler); });
    };

     const resetUI = () => {
         resetFiltersForFullUIReset(); 
         if (filterArea) filterArea.style.display = 'none'; 
         if (resultsArea) resultsArea.style.display = 'none';
         if (mainChartInstance) { mainChartInstance.destroy(); mainChartInstance = null; }
         if (attachRateTableBody) attachRateTableBody.innerHTML = ''; 
         if (attachRateTableFooter) attachRateTableFooter.innerHTML = ''; 
         if (attachTableStatus) attachTableStatus.textContent = '';
         if (topBottomSection) topBottomSection.style.display = 'none'; 
         if (top5TableBody) top5TableBody.innerHTML = ''; 
         if (bottom5TableBody) bottom5TableBody.innerHTML = '';
         hideStoreDetails(); 
         updateSummary([]); 
         if(statusDiv) statusDiv.textContent = 'No file selected.';
         allPossibleStores = []; 
         rawData = []; 
         filteredData = [];
     };

    const handleResetFiltersClick = () => {
        [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(sel => { 
            if (sel) sel.value = 'ALL'; 
        });
        if (territoryFilter) territoryFilter.selectedIndex = -1;
        if (storeFilter) storeFilter.selectedIndex = -1; 
        if (storeSearch) storeSearch.value = ''; 
        Object.values(flagFiltersCheckboxes).forEach(input => { if(input) input.checked = false; });

        if (rawData.length > 0) {
            updateStoreFilterOptionsBasedOnHierarchy(); 
        } else {
            if (storeFilter) { storeFilter.innerHTML = '<option value="ALL">-- Load File First --</option>'; storeFilter.disabled = true; }
            if (storeSearch) storeSearch.disabled = true;
            if (storeSelectAll) storeSelectAll.disabled = true;
            if (storeDeselectAll) storeDeselectAll.disabled = true;
        }

        if (resultsArea) resultsArea.style.display = 'none';
        if (topBottomSection) topBottomSection.style.display = 'none';
        if (mainChartInstance) { mainChartInstance.destroy(); mainChartInstance = null; }
        if (attachRateTableBody) attachRateTableBody.innerHTML = '';
        if (attachRateTableFooter) attachRateTableFooter.innerHTML = '';
        if (attachTableStatus) attachTableStatus.textContent = '';
        hideStoreDetails(); 

        filteredData = [];
        if (exportCsvButton) exportCsvButton.disabled = true;

        if (statusDiv) {
            if (rawData.length > 0) {
                statusDiv.textContent = 'Filters reset. Click "Apply Filters" to see results.';
            } else {
                statusDiv.textContent = 'No file selected. Load a file to use filters.';
            }
        }
    };

    const updateSummary = (data) => {
        const totalCount = data.length;
        const fieldsToClearText = [revenueWithDFValue, qtdRevenueTargetValue, qtdGapValue, quarterlyRevenueTargetValue, percentQuarterlyStoreTargetValue, revARValue, unitsWithDFValue, unitTargetValue, unitAchievementValue, visitCountValue, trainingCountValue, retailModeConnectivityValue, repSkillAchValue, vPmrAchValue, postTrainingScoreValue, eliteValue, percentQuarterlyTerritoryTargetValue, territoryRevPercentValue, districtRevPercentValue, regionRevPercentValue];
        fieldsToClearText.forEach(el => { if (el) el.textContent = 'N/A'; });
        [percentQuarterlyTerritoryTargetP, territoryRevPercentP, districtRevPercentP, regionRevPercentP].forEach(p => { if(p) p.style.display = 'none';});
        if (totalCount === 0) return;
        const sumRevenue = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Revenue w/DF', 0)), 0);
        const sumQtdTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'QTD Revenue Target', 0)), 0);
        const sumQuarterlyTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Quarterly Revenue Target', 0)), 0);
        const sumUnits = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Unit w/ DF', 0)), 0);
        const sumUnitTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Unit Target', 0)), 0);
        const sumVisits = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Visit count', 0)), 0);
        const sumTrainings = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Trainings', 0)), 0);
        let sumConnectivity = 0, countConnectivity = 0; let sumRepSkill = 0, countRepSkill = 0; let sumPmr = 0, countPmr = 0;
        let sumPostTraining = 0, countPostTraining = 0; let sumElite = 0, countElite = 0; 
        data.forEach(row => {
            let valStr; const subChannel = safeGet(row, 'SUB_CHANNEL', null); 
            valStr = safeGet(row, 'Retail Mode Connectivity', null); if (isValidForAverage(valStr)) { sumConnectivity += parsePercent(valStr); countConnectivity++; }
            valStr = safeGet(row, 'Rep Skill Ach', null); if (isValidForAverage(valStr)) { sumRepSkill += parsePercent(valStr); countRepSkill++; }
            valStr = safeGet(row, '(V)PMR Ach', null); if (isValidForAverage(valStr)) { sumPmr += parsePercent(valStr); countPmr++; }
            valStr = safeGet(row, 'Post Training Score', null); if (isValidForAverage(valStr)) { sumPostTraining += parseNumber(valStr); countPostTraining++; }
            if (subChannel !== "Verizon COR") { valStr = safeGet(row, 'Elite', null); if (isValidForAverage(valStr)) { sumElite += parsePercent(valStr); countElite++; } }
        });
        const calculatedRevAR = sumQtdTarget === 0 ? NaN : sumRevenue / sumQtdTarget;
        const avgConnectivity = countConnectivity > 0 ? sumConnectivity / countConnectivity : NaN; const avgRepSkill = countRepSkill > 0 ? sumRepSkill / countRepSkill : NaN;
        const avgPmr = countPmr > 0 ? sumPmr / countPmr : NaN; const avgPostTraining = countPostTraining > 0 ? sumPostTraining / countPostTraining : NaN;
        const avgElite = countElite > 0 ? sumElite / countElite : NaN; 
        const overallPercentStoreTarget = sumQuarterlyTarget !== 0 ? sumRevenue / sumQuarterlyTarget : NaN; const overallUnitAchievement = sumUnitTarget !== 0 ? sumUnits / sumUnitTarget : NaN;
        if (revenueWithDFValue) { revenueWithDFValue.textContent = formatCurrency(sumRevenue); revenueWithDFValue.title = `Sum of 'Revenue w/DF' for ${totalCount} filtered stores`; }
        if (qtdRevenueTargetValue) { qtdRevenueTargetValue.textContent = formatCurrency(sumQtdTarget); qtdRevenueTargetValue.title = `Sum of 'QTD Revenue Target' for ${totalCount} filtered stores`; }
        if (qtdGapValue) { qtdGapValue.textContent = formatCurrency(sumRevenue - sumQtdTarget); qtdGapValue.title = `Calculated Gap (Total Revenue - QTD Target) for ${totalCount} filtered stores`; }
        if (quarterlyRevenueTargetValue) { quarterlyRevenueTargetValue.textContent = formatCurrency(sumQuarterlyTarget); quarterlyRevenueTargetValue.title = `Sum of 'Quarterly Revenue Target' for ${totalCount} filtered stores`; }
        if (unitsWithDFValue) { unitsWithDFValue.textContent = formatNumber(sumUnits); unitsWithDFValue.title = `Sum of 'Unit w/ DF' for ${totalCount} filtered stores`; }
        if (unitTargetValue) { unitTargetValue.textContent = formatNumber(sumUnitTarget); unitTargetValue.title = `Sum of 'Unit Target' for ${totalCount} filtered stores`; }
        if (visitCountValue) { visitCountValue.textContent = formatNumber(sumVisits); visitCountValue.title = `Sum of 'Visit count' for ${totalCount} filtered stores`; }
        if (trainingCountValue) { trainingCountValue.textContent = formatNumber(sumTrainings); trainingCountValue.title = `Sum of 'Trainings' for ${totalCount} filtered stores`; }
        if (revARValue) { revARValue.textContent = formatPercent(calculatedRevAR); revARValue.title = "Rev AR% for selected stores with data"; }
        if (percentQuarterlyStoreTargetValue) { percentQuarterlyStoreTargetValue.textContent = formatPercent(overallPercentStoreTarget); percentQuarterlyStoreTargetValue.title = `Overall % Quarterly Target (Total Revenue / Total Quarterly Target)`; }
        if (unitAchievementValue) { unitAchievementValue.textContent = formatPercent(overallUnitAchievement); unitAchievementValue.title = `Overall Unit Achievement % (Total Units / Total Unit Target)`; }
        if (retailModeConnectivityValue) { retailModeConnectivityValue.textContent = formatPercent(avgConnectivity); retailModeConnectivityValue.title = `Average 'Retail Mode Connectivity' across ${countConnectivity} stores with data`; }
        if (repSkillAchValue) { repSkillAchValue.textContent = formatPercent(avgRepSkill); repSkillAchValue.title = `Average 'Rep Skill Ach' across ${countRepSkill} stores with data`; }
        if (vPmrAchValue) { vPmrAchValue.textContent = formatPercent(avgPmr); vPmrAchValue.title = `Average '(V)PMR Ach' across ${countPmr} stores with data`; }
        if (postTrainingScoreValue) { postTrainingScoreValue.textContent = isNaN(avgPostTraining) ? 'N/A' : avgPostTraining.toFixed(1); postTrainingScoreValue.title = `Average 'Post Training Score' across ${countPostTraining} stores with data`; }
        if (eliteValue) { eliteValue.textContent = formatPercent(avgElite); eliteValue.title = `Average 'Elite' score % across ${countElite} stores with data (excluding Verizon COR sub-channel)`;}
        updateContextualSummary(data);
    };

    const updateContextualSummary = (data) => {
        [percentQuarterlyTerritoryTargetP, territoryRevPercentP, districtRevPercentP, regionRevPercentP].forEach(p => {if (p) p.style.display = 'none'});
        if (data.length === 0) return;
        const singleRegion = regionFilter?.value !== 'ALL'; const singleDistrict = districtFilter?.value !== 'ALL'; const singleTerritory = territoryFilter && territoryFilter.selectedOptions.length === 1;
        const calculateAverageExcludeBlanks = (column) => { let sum = 0, count = 0; data.forEach(row => { const valStr = safeGet(row, column, null); if (isValidForAverage(valStr)) { sum += parsePercent(valStr); count++; } }); return count > 0 ? sum / count : NaN; };
        const avgPercentTerritoryTarget = calculateAverageExcludeBlanks('%Quarterly Territory Rev Target'); const avgTerritoryRevPercent = calculateAverageExcludeBlanks('Territory Rev%');
        const avgDistrictRevPercent = calculateAverageExcludeBlanks('District Rev%'); const avgRegionRevPercent = calculateAverageExcludeBlanks('Region Rev%');
        if (percentQuarterlyTerritoryTargetValue) percentQuarterlyTerritoryTargetValue.textContent = formatPercent(avgPercentTerritoryTarget); if (percentQuarterlyTerritoryTargetP && !isNaN(avgPercentTerritoryTarget)) percentQuarterlyTerritoryTargetP.style.display = 'block';
        if (singleTerritory || singleDistrict || singleRegion) { if (territoryRevPercentValue) territoryRevPercentValue.textContent = formatPercent(avgTerritoryRevPercent); if (territoryRevPercentP && !isNaN(avgTerritoryRevPercent)) territoryRevPercentP.style.display = 'block'; }
        if (singleDistrict || singleRegion) { if (districtRevPercentValue) districtRevPercentValue.textContent = formatPercent(avgDistrictRevPercent); if (districtRevPercentP && !isNaN(avgDistrictRevPercent)) districtRevPercentP.style.display = 'block'; }
        if (singleRegion) { if (regionRevPercentValue) regionRevPercentValue.textContent = formatPercent(avgRegionRevPercent); if (regionRevPercentP && !isNaN(avgRegionRevPercent)) regionRevPercentP.style.display = 'block'; }
    };

    const updateCharts = (data) => {
        if (mainChartInstance) { mainChartInstance.destroy(); mainChartInstance = null; }
        if (!mainChartCanvas || (data.length === 0 && rawData.length === 0) ) { if (mainChartCanvas && mainChartInstance) { mainChartInstance = new Chart(mainChartCanvas, {type: 'bar', data: {labels:[], datasets:[]}}); mainChartInstance.destroy(); mainChartInstance = null;} return; }
        const chartThemeColors = getChartThemeColors(); 
        const sortedData = [...data].sort((a, b) => parseNumber(safeGet(b, 'Revenue w/DF', 0)) - parseNumber(safeGet(a, 'Revenue w/DF', 0)));
        const chartData = sortedData.slice(0, TOP_N_CHART);
        const labels = chartData.map(row => safeGet(row, 'Store', 'Unknown Store'));
        const revenueDataSet = chartData.map(row => parseNumber(safeGet(row, 'Revenue w/DF', 0)));
        const targetDataSet = chartData.map(row => parseNumber(safeGet(row, 'QTD Revenue Target', 0)));
        const backgroundColors = chartData.map((_, index) => revenueDataSet[index] >= targetDataSet[index] ? 'rgba(75, 192, 192, 0.6)' : 'rgba(255, 99, 132, 0.6)'); 
        const borderColors = chartData.map((_, index) => revenueDataSet[index] >= targetDataSet[index] ? 'rgba(75, 192, 192, 1)' : 'rgba(255, 99, 132, 1)');
        mainChartInstance = new Chart(mainChartCanvas, {
            type: 'bar', data: { labels: labels, datasets: [ { label: 'Total Revenue (incl. DF)', data: revenueDataSet, backgroundColor: backgroundColors, borderColor: borderColors, borderWidth: 1 }, { label: 'QTD Revenue Target', data: targetDataSet, type: 'line', borderColor: 'rgba(255, 206, 86, 1)', backgroundColor: 'rgba(255, 206, 86, 0.2)', fill: false, tension: 0.1, borderWidth: 2, pointRadius: 3, pointHoverRadius: 5 } ] },
            options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, ticks: { color: chartThemeColors.tickColor, callback: value => formatCurrency(value) }, grid: { color: chartThemeColors.gridColor } }, x: { ticks: { color: chartThemeColors.tickColor }, grid: { display: false } } }, plugins: { legend: { labels: { color: chartThemeColors.legendColor } }, tooltip: { callbacks: { label: function(context) { let label = context.dataset.label || ''; if (label) label += ': '; if (context.parsed.y !== null) { label += formatCurrency(context.parsed.y); if (context.dataset.type !== 'line' && chartData[context.dataIndex]) { const storeData = chartData[context.dataIndex]; const percentQtrTarget = parsePercent(safeGet(storeData, '% Quarterly Revenue Target', 0)); if (!isNaN(percentQtrTarget)) { label += ` (${formatPercent(percentQtrTarget)} of Qtr Target)`; } } } return label; } } } }, onClick: (_, elements) => { if (elements.length > 0) { const index = elements[0].index; const storeName = labels[index]; const storeData = filteredData.find(row => safeGet(row, 'Store', null) === storeName); if (storeData) { showStoreDetails(storeData); highlightTableRow(storeName); } } } }
        });
    };

    const updateAttachRateTable = (data) => {
        if (!attachRateTableBody || !attachRateTableFooter) return;
        attachRateTableBody.innerHTML = ''; attachRateTableFooter.innerHTML = '';
        if (data.length === 0) { if(attachTableStatus) attachTableStatus.textContent = 'No data to display based on filters.'; return; }
        
        const sortedData = [...data].sort((a, b) => {
             let valA = safeGet(a, currentSort.column, null); let valB = safeGet(b, currentSort.column, null);
             if (valA === null && valB === null) return 0; if (valA === null) return currentSort.ascending ? -1 : 1; if (valB === null) return currentSort.ascending ? 1 : -1;
             const isPercentCol = currentSort.column.includes('Attach Rate') || currentSort.column.includes('% Target'); // Keep this for sorting main table
             const numA = isPercentCol ? parsePercent(valA) : parseNumber(valA); const numB = isPercentCol ? parsePercent(valB) : parseNumber(valB);
             if (!isNaN(numA) && !isNaN(numB)) { return currentSort.ascending ? numA - numB : numB - numA; }
             else { valA = String(valA).toLowerCase(); valB = String(valB).toLowerCase(); return currentSort.ascending ? valA.localeCompare(valB) : valB.localeCompare(valA); }
        });

        // ** MODIFIED: Define columns for Attach Rate Table (excluding '% Quarterly Revenue Target') **
        const columns = [
            { key: 'Store', format: (val) => val },
            // '% Quarterly Revenue Target' removed from here
            { key: 'Tablet Attach Rate', format: formatPercent, highlight: true },
            { key: 'PC Attach Rate', format: formatPercent, highlight: true },
            { key: 'NC Attach Rate', format: formatPercent, highlight: true },
            { key: 'TWS Attach Rate', format: formatPercent, highlight: true },
            { key: 'WW Attach Rate', format: formatPercent, highlight: true },
            { key: 'ME Attach Rate', format: formatPercent, highlight: true },
            { key: 'NCME Attach Rate', format: formatPercent, highlight: true },
        ];
        
        // ** MODIFIED: averageMetrics for footer (excluding '% Quarterly Revenue Target') **
        const averageMetrics = [
             // '% Quarterly Revenue Target' removed
             'Tablet Attach Rate', 'PC Attach Rate', 'NC Attach Rate',
             'TWS Attach Rate', 'WW Attach Rate', 'ME Attach Rate', 'NCME Attach Rate'
        ];
        const averages = {};
        averageMetrics.forEach(key => { 
            let sum = 0, count = 0; 
            data.forEach(row => { 
                const valStr = safeGet(row, key, null); 
                if (isValidForAverage(valStr)) { sum += parsePercent(valStr); count++; } 
            }); 
            averages[key] = count > 0 ? sum / count : NaN; 
        });

        sortedData.forEach(row => {
            const tr = document.createElement('tr'); const storeName = safeGet(row, 'Store', null);
            if (storeName && String(storeName).trim() !== '') {
                 tr.dataset.storeName = storeName; tr.onclick = () => { showStoreDetails(row); highlightTableRow(storeName); };
                 // Use the MODIFIED 'columns' array here
                 columns.forEach(col => {
                     const td = document.createElement('td'); const rawValue = safeGet(row, col.key, null); 
                     const isPercentCol = col.key.includes('Attach Rate'); // Simplified as '% Target' is removed
                     const numericValue = (col.key === 'Store') ? rawValue : (isPercentCol ? parsePercent(rawValue) : parseNumber(rawValue));
                     let formattedValue; 
                     if (rawValue === null || (col.key !== 'Store' && isNaN(numericValue)) || String(rawValue).trim() === '') { formattedValue = 'N/A'; } 
                     else { formattedValue = typeof col.format === 'function' ? col.format(numericValue) : numericValue; }
                     td.textContent = formattedValue; td.title = `${col.key}: ${formattedValue}`;
                     if (col.highlight && !isNaN(averages[col.key]) && typeof numericValue === 'number' && !isNaN(numericValue)) { 
                         td.classList.toggle('highlight-green', numericValue >= averages[col.key]); 
                         td.classList.toggle('highlight-red', numericValue < averages[col.key]); 
                     }
                     tr.appendChild(td);
                 }); 
                 attachRateTableBody.appendChild(tr);
            }
        });

        if (data.length > 0) {
            const footerRow = attachRateTableFooter.insertRow(); 
            const avgLabelCell = footerRow.insertCell(); 
            avgLabelCell.textContent = 'Filtered Avg*';
            avgLabelCell.title = 'Average calculated only using stores with valid data for each column'; 
            avgLabelCell.style.textAlign = "right"; 
            avgLabelCell.style.fontWeight = "bold";
            // Use the MODIFIED 'averageMetrics' array here
            averageMetrics.forEach(key => { 
                const td = footerRow.insertCell(); 
                const avgValue = averages[key]; 
                td.textContent = formatPercent(avgValue); 
                let validCount = data.filter(r => isValidForAverage(safeGet(r, key, null))).length; 
                td.title = `Average ${key}: ${formatPercent(avgValue)} (from ${validCount} stores)`; 
                td.style.textAlign = "right"; 
            });
        }
        if(attachTableStatus) attachTableStatus.textContent = `Showing ${attachRateTableBody.rows.length} stores. Click row for details. Click headers to sort.`;
        updateSortArrows();
    };

    const handleSort = (event) => {
         const headerCell = event.target.closest('th'); if (!headerCell?.classList.contains('sortable')) return;
         const sortKey = headerCell.dataset.sort; if (!sortKey) return;
         if (currentSort.column === sortKey) { currentSort.ascending = !currentSort.ascending; } else { currentSort.column = sortKey; currentSort.ascending = true; }
         updateAttachRateTable(filteredData);
    };

    const updateSortArrows = () => {
        if (!attachRateTable) return;
        attachRateTable.querySelectorAll('th.sortable .sort-arrow').forEach(arrow => { arrow.className = 'sort-arrow'; arrow.textContent = ''; });
        const currentHeaderArrow = attachRateTable.querySelector(`th[data-sort="${CSS.escape(currentSort.column)}"] .sort-arrow`);
        if (currentHeaderArrow) { currentHeaderArrow.classList.add(currentSort.ascending ? 'asc' : 'desc'); }
    };

    const showStoreDetails = (storeData) => {
        if (!storeDetailsContent || !storeDetailsSection || !closeStoreDetailsButton) return;
        const addressParts = [ safeGet(storeData, 'ADDRESS1', null), safeGet(storeData, 'CITY', null), safeGet(storeData, 'STATE', null), safeGet(storeData, 'ZIPCODE', null) ].filter(part => part && String(part).trim() !== '');
        const addressString = addressParts.length > 0 ? addressParts.join(', ') : 'N/A';
        const latitude = parseNumber(safeGet(storeData, 'LATITUDE_ORG', NaN)); const longitude = parseNumber(safeGet(storeData, 'LONGITUDE_ORG', NaN));
        let mapsLinkHtml = `<p style="color: #aaa; font-style: italic;">(Map coordinates not available)</p>`;
        if (!isNaN(latitude) && !isNaN(longitude)) { const mapsUrl = `https://www.google.com/maps?q=${latitude},${longitude}`; mapsLinkHtml = `<p><a href="${mapsUrl}" target="_blank" title="Open in Google Maps">View on Google Maps</a></p>`; }
        let flagSummaryHtml = FLAG_HEADERS.map(flag => { const flagValue = safeGet(storeData, flag, 'NO'); const isTrue = (flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1'); return `<span title="${flag.replace(/_/g, ' ')}" data-flag="${isTrue}">${flag.replace(/_/g, ' ')} ${isTrue ? '✔' : '✘'}</span>`; }).join(' | ');
        storeDetailsContent.innerHTML = ` <p><strong>Store:</strong> ${safeGet(storeData, 'Store')}</p> <p><strong>Address:</strong> ${addressString}</p> ${mapsLinkHtml} <hr> <p><strong>IDs:</strong> Store: ${safeGet(storeData, 'STORE ID')} | Org: ${safeGet(storeData, 'ORG_STORE_ID')} | CV: ${safeGet(storeData, 'CV_STORE_ID')} | CinglePoint: ${safeGet(storeData, 'CINGLEPOINT_ID')}</p> <p><strong>Type:</strong> ${safeGet(storeData, 'STORE_TYPE_NAME')} | Nat Tier: ${safeGet(storeData, 'National_Tier')} | Merch Lvl: ${safeGet(storeData, 'Merchandising_Level')} | Comb Tier: ${safeGet(storeData, 'Combined_Tier')}</p> <hr> <p><strong>Hierarchy:</strong> ${safeGet(storeData, 'REGION')} > ${safeGet(storeData, 'DISTRICT')} > ${safeGet(storeData, 'Q2 Territory')}</p> <p><strong>FSM:</strong> ${safeGet(storeData, 'FSM NAME')}</p> <p><strong>Channel:</strong> ${safeGet(storeData, 'CHANNEL')} / ${safeGet(storeData, 'SUB_CHANNEL')}</p> <p><strong>Dealer:</strong> ${safeGet(storeData, 'DEALER_NAME')}</p> <hr> <p><strong>Visits:</strong> ${formatNumber(parseNumber(safeGet(storeData, 'Visit count', 0)))} | <strong>Trainings:</strong> ${formatNumber(parseNumber(safeGet(storeData, 'Trainings', 0)))}</p> <p><strong>Connectivity:</strong> ${formatPercent(parsePercent(safeGet(storeData, 'Retail Mode Connectivity', 0)))}</p> <hr> <p><strong>Flags:</strong> ${flagSummaryHtml}</p> `;
        storeDetailsSection.style.display = 'block'; closeStoreDetailsButton.style.display = 'inline-block';
    };

    const hideStoreDetails = () => {
        if (!storeDetailsContent || !storeDetailsSection || !closeStoreDetailsButton) return;
        storeDetailsContent.innerHTML = 'Select a store from the table or chart for details, or apply filters resulting in a single store.';
        storeDetailsSection.style.display = 'none'; closeStoreDetailsButton.style.display = 'none'; highlightTableRow(null);
    };

    const highlightTableRow = (storeName) => {
        if (selectedStoreRow) { selectedStoreRow.classList.remove('selected-row'); selectedStoreRow = null; }
        if (storeName) {
            const tablesToSearch = [attachRateTableBody, top5TableBody, bottom5TableBody];
            for (const tableBody of tablesToSearch) {
                if (tableBody) { 
                    try { const rowToHighlight = tableBody.querySelector(`tr[data-store-name="${CSS.escape(storeName)}"]`); if (rowToHighlight) { rowToHighlight.classList.add('selected-row'); selectedStoreRow = rowToHighlight; break; }
                    } catch (e) { console.error("Error selecting table row in highlightTableRow:", e, "StoreName:", storeName, "Table:", tableBody); }
                }
            }
        }
    };

    const exportData = () => {
        if (filteredData.length === 0) { alert("No filtered data to export."); return; }
        try {
            if (!attachRateTable) throw new Error("Attach rate table not found for headers.");
            // Get headers from the modified attachRateTable (which no longer has % Qtr Rev Target)
            const headers = Array.from(attachRateTable.querySelectorAll('thead th'))
                                 .map(th => th.dataset.sort || th.textContent.replace(/ [▲▼]$/, '').trim());
            const dataToExport = filteredData.map(row => {
                return headers.map(headerKey => { // This headerKey comes from the current table headers
                    let value = safeGet(row, headerKey, ''); 
                    // Note: If headerKey is for a column that was removed (like '% Qtr Rev Target'), 
                    // it won't be in `headers` anymore, so safeGet will use its default for that key if accessed directly.
                    // But since we iterate `headers` from the table, this is fine.
                    const isPercentLike = headerKey.includes('%') || headerKey.includes('Rate') || headerKey.includes('Ach') || headerKey.includes('Connectivity') || headerKey.includes('Elite');
                    if (isPercentLike) { const numVal = parsePercent(value); return isNaN(numVal) ? '' : numVal; } 
                    else { const numVal = parseNumber(value); if (!isNaN(numVal) && typeof value !== 'boolean' && String(value).trim() !== '') { return numVal; } if (typeof value === 'string' && (value.includes(',') || value.includes('"') || value.includes('\n'))) { return `"${value.replace(/"/g, '""')}"`; } return value; }
                });
            });
            let csvContent = "data:text/csv;charset=utf-8," + headers.join(",") + "\n" + dataToExport.map(e => e.join(",")).join("\n");
            const encodedUri = encodeURI(csvContent); const link = document.createElement("a"); link.setAttribute("href", encodedUri); link.setAttribute("download", "fsm_dashboard_export.csv"); document.body.appendChild(link); link.click(); document.body.removeChild(link);
        } catch (error) { console.error("Error exporting CSV:", error); alert("Error generating CSV export. See console for details."); }
    };

    const generateEmailBody = () => {
        if (filteredData.length === 0) return "No data available based on current filters.";
        let body = "FSM Dashboard Summary:\n---------------------------------\n"; body += `Filters Applied: ${getFilterSummary()}\nStores Found: ${filteredData.length}\n---------------------------------\n\n`;
        body += "Performance Summary:\n"; body += `- Total Revenue (incl. DF): ${revenueWithDFValue?.textContent || 'N/A'}\n`; body += `- QTD Revenue Target: ${qtdRevenueTargetValue?.textContent || 'N/A'}\n`;
        body += `- QTD Gap: ${qtdGapValue?.textContent || 'N/A'}\n`; body += `- Rev AR%: ${revARValue?.textContent || 'N/A'} (Calculated as: Total Revenue w/DF / Total QTD Revenue Target)\n`;
        body += `- % Store Quarterly Target: ${percentQuarterlyStoreTargetValue?.textContent || 'N/A'}\n`; body += `- Total Units (incl. DF): ${unitsWithDFValue?.textContent || 'N/A'}\n`;
        body += `- Unit Achievement %: ${unitAchievementValue?.textContent || 'N/A'}\n`; body += `- Total Visits: ${visitCountValue?.textContent || 'N/A'}\n`; body += `- Avg. Connectivity: ${retailModeConnectivityValue?.textContent || 'N/A'}\n\n`;
        body += "Mysteryshop & Training (Avg*):\n"; body += `- Rep Skill Ach: ${repSkillAchValue?.textContent || 'N/A'}\n`; body += `- (V)PMR Ach: ${vPmrAchValue?.textContent || 'N/A'}\n`;
        body += `- Post Training Score: ${postTrainingScoreValue?.textContent || 'N/A'}\n`; body += `- Elite Score %: ${eliteValue?.textContent || 'N/A'} (Excludes Verizon COR sub-channel)\n\n`;
        body += "*Averages calculated only using stores with valid data for each metric.\n\n";
        const territoriesInData = new Set(filteredData.map(row => safeGet(row, 'Q2 Territory', null)).filter(Boolean));
        if (territoriesInData.size === 1) {
            const territoryName = territoriesInData.values().next().value; body += `--- Key Performers for Territory: ${territoryName} ---\n`;
            const top5ForEmail = [...filteredData].sort((a, b) => parseNumber(safeGet(b, 'Revenue w/DF', -Infinity)) - parseNumber(safeGet(a, 'Revenue w/DF', -Infinity))).slice(0, TOP_N_TABLES);
            if (top5ForEmail.length > 0) { body += `Top ${top5ForEmail.length} (Revenue):\n`; top5ForEmail.forEach((store, i) => { body += `${i+1}. ${safeGet(store, 'Store')} - Rev: ${formatCurrency(parseNumber(safeGet(store, 'Revenue w/DF')))}\n`; }); body += "\n"; }
            const bottom5ForEmail = [...filteredData].sort((a, b) => calculateQtdGap(a) - calculateQtdGap(b)).slice(0, TOP_N_TABLES);
            if (bottom5ForEmail.length > 0) { body += `Bottom ${bottom5ForEmail.length} (Opportunities by QTD Gap):\n`; bottom5ForEmail.forEach((store, i) => { body += `${i+1}. ${safeGet(store, 'Store')} - Gap: ${formatCurrency(calculateQtdGap(store) === Infinity ? NaN : calculateQtdGap(store))}\n`; }); body += "\n"; }
        }
        body += "---------------------------------\nGenerated by FSM Dashboard\n"; return body;
    };

    const getFilterSummary = () => {
        let summary = [];
        if (regionFilter?.value !== 'ALL') summary.push(`Region: ${regionFilter.value}`); if (districtFilter?.value !== 'ALL') summary.push(`District: ${districtFilter.value}`);
        const territories = territoryFilter ? Array.from(territoryFilter.selectedOptions).map(o => o.value) : []; if (territories.length > 0) summary.push(`Territories: ${territories.length === 1 ? territories[0] : `${territories.length} selected`}`);
        if (fsmFilter?.value !== 'ALL') summary.push(`FSM: ${fsmFilter.value}`); if (channelFilter?.value !== 'ALL') summary.push(`Channel: ${channelFilter.value}`);
        if (subchannelFilter?.value !== 'ALL') summary.push(`Subchannel: ${subchannelFilter.value}`); if (dealerFilter?.value !== 'ALL') summary.push(`Dealer: ${dealerFilter.value}`);
        const stores = storeFilter ? Array.from(storeFilter.selectedOptions).map(o => o.value) : []; if (stores.length > 0) summary.push(`Stores: ${stores.length === 1 ? stores[0] : `${stores.length} selected`}`);
        const flags = Object.entries(flagFiltersCheckboxes).filter(([, input]) => input?.checked).map(([key])=> key.replace(/_/g, ' ')); if (flags.length > 0) summary.push(`Attributes: ${flags.join(', ')}`);
        return summary.length > 0 ? summary.join('; ') : 'None';
    };

    const handleShareEmail = () => {
        if (!emailRecipientInput || !shareStatus) return; const recipient = emailRecipientInput.value;
        if (!recipient || !/\S+@\S+\.\S+/.test(recipient)) { shareStatus.textContent = "Please enter a valid recipient email address."; return; }
        try {
            const subject = `FSM Dashboard Summary - ${new Date().toLocaleDateString()}`; const bodyContent = generateEmailBody();
            const mailtoLink = `mailto:${recipient}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(bodyContent)}`;
            if (mailtoLink.length > 2000) { shareStatus.textContent = "Generated email body is too long for a mailto link. Try applying more filters or copy the content manually."; console.warn("Mailto link length exceeds 2000 characters:", mailtoLink.length); return; }
            window.location.href = mailtoLink; shareStatus.textContent = "Your email client should open. Please review and send the email.";
        } catch (error) { console.error("Error generating mailto link:", error); shareStatus.textContent = "Error generating email content. Check console for details."; }
    };

     const selectAllOptions = (selectElement) => {
         if (!selectElement) return; Array.from(selectElement.options).forEach(option => option.selected = true);
         if (selectElement === territoryFilter) updateStoreFilterOptionsBasedOnHierarchy();
    };

     const deselectAllOptions = (selectElement) => {
         if (!selectElement) return; selectElement.selectedIndex = -1;
         if (selectElement === territoryFilter) updateStoreFilterOptionsBasedOnHierarchy();
    };

    // --- Event Listeners ---
    if (themeToggleBtn) themeToggleBtn.addEventListener('click', toggleTheme);
    if (resetFiltersButton) resetFiltersButton.addEventListener('click', handleResetFiltersClick); 
    
    excelFileInput?.addEventListener('change', handleFile);
    applyFiltersButton?.addEventListener('click', applyFilters);
    storeSearch?.addEventListener('input', filterStoreOptions);
    exportCsvButton?.addEventListener('click', exportData);
    shareEmailButton?.addEventListener('click', handleShareEmail);
    closeStoreDetailsButton?.addEventListener('click', hideStoreDetails);
    territorySelectAll?.addEventListener('click', () => selectAllOptions(territoryFilter));
    territoryDeselectAll?.addEventListener('click', () => deselectAllOptions(territoryFilter));
    storeSelectAll?.addEventListener('click', () => selectAllOptions(storeFilter));
    storeDeselectAll?.addEventListener('click', () => deselectAllOptions(storeFilter));
    attachRateTable?.querySelector('thead')?.addEventListener('click', handleSort);

    // --- Initial Setup ---
    const savedTheme = localStorage.getItem(THEME_STORAGE_KEY); 
    applyTheme(savedTheme || 'dark'); 

    resetUI(); 
    if (!mainChartCanvas) console.warn("Main chart canvas context not found on load. Chart will not render.");

}); // End DOMContentLoaded
