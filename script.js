/* Generated: 2025-04-05 12:28:51 PM EDT - Attempt 3: Add dark/light theme toggle, ensuring all JS code is correctly placed within DOMContentLoaded scope. */
document.addEventListener('DOMContentLoaded', () => {
    // --- Configuration ---
    const REQUIRED_HEADERS = [
        'Store', 'REGION', 'DISTRICT', 'Q2 Territory', 'FSM NAME', 'CHANNEL',
        'SUB_CHANNEL', 'DEALER_NAME', 'Revenue w/DF', 'QTD Revenue Target',
        'Quarterly Revenue Target', 'QTD Gap', '% Quarterly Revenue Target', //'% Quarterly Revenue Target' used in attach table
        'Unit w/ DF', 'Unit Target', 'Unit Achievement', 'Visit count', 'Trainings',
        'Retail Mode Connectivity', 'Rep Skill Ach', '(V)PMR Ach', 'Elite', 'Post Training Score',
        'Tablet Attach Rate', 'PC Attach Rate', 'NC Attach Rate', 'TWS Attach Rate',
        'WW Attach Rate', 'ME Attach Rate', 'NCME Attach Rate', 'SUPER STORE', 'GOLDEN RHINO',
        'GCE', 'AI_Zone', 'Hispanic_Market', 'EV ROUTE', 'STORE ID', 'ADDRESS1', 'CITY', 'STATE', 'ZIPCODE',
        'LATITUDE_ORG', 'LONGITUDE_ORG', 'ORG_STORE_ID', 'CV_STORE_ID', 'CINGLEPOINT_ID',
        'STORE_TYPE_NAME', 'National_Tier', 'Merchandising_Level', 'Combined_Tier',
        '%Quarterly Territory Rev Target', 'Region Rev%', 'District Rev%', 'Territory Rev%'
    ];
    const FLAG_HEADERS = ['SUPER STORE', 'GOLDEN RHINO', 'GCE', 'AI_Zone', 'Hispanic_Market', 'EV ROUTE'];
    const ATTACH_RATE_KEYS = [
         'Tablet Attach Rate', 'PC Attach Rate', 'NC Attach Rate',
         'TWS Attach Rate', 'WW Attach Rate', 'ME Attach Rate', 'NCME Attach Rate'
    ];
    const CURRENCY_FORMAT = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' });
    const PERCENT_FORMAT = new Intl.NumberFormat('en-US', { style: 'percent', minimumFractionDigits: 1, maximumFractionDigits: 1 });
    const NUMBER_FORMAT = new Intl.NumberFormat('en-US');
    const CHART_COLORS = ['#58a6ff', '#ffb758', '#86dc86', '#ff7f7f', '#b796e6', '#ffda8a', '#8ad7ff', '#ff9ba6']; // Example dark mode colors
    const TOP_N_CHART = 15;
    const TOP_N_TABLES = 5;
    const THEME_KEY = 'fsmDashboardTheme'; // Key for localStorage

    // --- DOM Elements ---
    const bodyElement = document.body; // Reference to body for theme class
    const themeToggleButton = document.getElementById('themeToggle'); // Theme toggle button
    const excelFileInput = document.getElementById('excelFile');
    const statusDiv = document.getElementById('status');
    const loadingIndicator = document.getElementById('loadingIndicator');
    const filterLoadingIndicator = document.getElementById('filterLoadingIndicator');
    const filterArea = document.getElementById('filterArea');
    const resultsArea = document.getElementById('resultsArea');
    const applyFiltersButton = document.getElementById('applyFiltersButton');
    // Filter Elements
    const regionFilter = document.getElementById('regionFilter');
    const districtFilter = document.getElementById('districtFilter');
    const territoryFilter = document.getElementById('territoryFilter');
    const fsmFilter = document.getElementById('fsmFilter');
    const channelFilter = document.getElementById('channelFilter');
    const subchannelFilter = document.getElementById('subchannelFilter');
    const dealerFilter = document.getElementById('dealerFilter');
    const storeFilter = document.getElementById('storeFilter');
    const storeSearch = document.getElementById('storeSearch');
    // Flag Filter Checkboxes
    const flagFiltersCheckboxes = FLAG_HEADERS.reduce((acc, header) => {
        let expectedId = '';
        switch (header) {
            case 'SUPER STORE':       expectedId = 'superStoreFilter'; break;
            case 'GOLDEN RHINO':      expectedId = 'goldenRhinoFilter'; break;
            case 'GCE':               expectedId = 'gceFilter'; break;
            case 'AI_Zone':           expectedId = 'aiZoneFilter'; break;
            case 'Hispanic_Market':   expectedId = 'hispanicMarketFilter'; break;
            case 'EV ROUTE':          expectedId = 'evRouteFilter'; break;
            default: console.warn(`Unknown flag header: ${header}`); return acc;
        }
        const element = document.getElementById(expectedId);
        if (element) { acc[header] = element; }
        else { console.warn(`Flag filter checkbox not found for ID: ${expectedId} (Header: ${header})`); }
        return acc;
    }, {});
    const territorySelectAll = document.getElementById('territorySelectAll');
    const territoryDeselectAll = document.getElementById('territoryDeselectAll');
    const storeSelectAll = document.getElementById('storeSelectAll');
    const storeDeselectAll = document.getElementById('storeDeselectAll');
    // Summary Elements
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
    // Contextual Summary Elements
    const percentQuarterlyTerritoryTargetP = document.getElementById('percentQuarterlyTerritoryTargetP');
    const territoryRevPercentP = document.getElementById('territoryRevPercentP');
    const districtRevPercentP = document.getElementById('districtRevPercentP');
    const regionRevPercentP = document.getElementById('regionRevPercentP');
    const percentQuarterlyTerritoryTargetValue = document.getElementById('percentQuarterlyTerritoryTargetValue');
    const territoryRevPercentValue = document.getElementById('territoryRevPercentValue');
    const districtRevPercentValue = document.getElementById('districtRevPercentValue');
    const regionRevPercentValue = document.getElementById('regionRevPercentValue');
    // Table Elements
    const attachRateTableBody = document.getElementById('attachRateTableBody');
    const attachRateTableFooter = document.getElementById('attachRateTableFooter');
    const attachTableStatus = document.getElementById('attachTableStatus');
    const attachRateTable = document.getElementById('attachRateTable');
    const exportCsvButton = document.getElementById('exportCsvButton');
    // Top/Bottom 5 Elements
    const topBottomSection = document.getElementById('topBottomSection');
    const top5TableBody = document.getElementById('top5TableBody');
    const bottom5TableBody = document.getElementById('bottom5TableBody');
    // Chart Elements
    const mainChartCanvas = document.getElementById('mainChartCanvas').getContext('2d');
    // Store Details Elements
    const storeDetailsSection = document.getElementById('storeDetailsSection');
    const storeDetailsContent = document.getElementById('storeDetailsContent');
    const closeStoreDetailsButton = document.getElementById('closeStoreDetailsButton');
    // Share Elements
    const emailRecipientInput = document.getElementById('emailRecipient');
    const shareEmailButton = document.getElementById('shareEmailButton');
    const shareStatus = document.getElementById('shareStatus');

    // --- Global State ---
    let rawData = [];
    let filteredData = [];
    let mainChartInstance = null;
    let storeOptions = [];
    let allPossibleStores = [];
    let currentSort = { column: 'Store', ascending: true }; // For Attach Rate Table
    let selectedStoreRow = null;
    let currentTheme = localStorage.getItem(THEME_KEY) || 'dark'; // Default theme

    // --- Helper Functions (Defined *before* use, inside DOMContentLoaded) ---
    const formatCurrency = (value) => isNaN(value) ? 'N/A' : CURRENCY_FORMAT.format(value);
    const formatPercent = (value) => isNaN(value) ? 'N/A' : PERCENT_FORMAT.format(value);
    const formatNumber = (value) => isNaN(value) ? 'N/A' : NUMBER_FORMAT.format(value);
    const parseNumber = (value) => { if (value === null || value === undefined || value === '') return NaN; if (typeof value === 'number') return value; if (typeof value === 'string') { value = value.replace(/[$,%]/g, ''); const num = parseFloat(value); return isNaN(num) ? NaN : num; } return NaN; };
    const parsePercent = (value) => { if (value === null || value === undefined || value === '') return NaN; if (typeof value === 'number') return value; if (typeof value === 'string') { const num = parseFloat(value.replace('%', '')); return isNaN(num) ? NaN : num / 100; } return NaN; };
    const safeGet = (obj, path, defaultValue = 'N/A') => { if (defaultValue === null && obj && obj[path] === 0) { return 0; } const value = obj ? obj[path] : undefined; return (value !== undefined && value !== null) ? value : defaultValue; };
    const isValidForAverage = (value) => { if (value === null || value === undefined || value === '') return false; return !isNaN(parseNumber(String(value).replace('%',''))); };
    const isValidAndNonZeroForAverage = (value) => { if (!isValidForAverage(value)) return false; const num = parseNumber(String(value).replace('%','')); return num !== 0; };
    const calculateQtdGap = (row) => { const revenue = parseNumber(safeGet(row, 'Revenue w/DF', 0)); const target = parseNumber(safeGet(row, 'QTD Revenue Target', 0)); if (isNaN(revenue) || isNaN(target)) { return Infinity; } return revenue - target; };
    const calculateRevARPercent = (row) => { const revenue = parseNumber(safeGet(row, 'Revenue w/DF', 0)); const target = parseNumber(safeGet(row, 'QTD Revenue Target', 0)); if (isNaN(revenue) || isNaN(target) || target === 0) { return NaN; } return revenue / target; };
    const calculateUnitAchievementPercent = (row) => { const units = parseNumber(safeGet(row, 'Unit w/ DF', 0)); const target = parseNumber(safeGet(row, 'Unit Target', 0)); if (isNaN(units) || isNaN(target) || target === 0) { return NaN; } return units / target; };
    const getUniqueValues = (data, column) => { const values = new Set(data.map(item => safeGet(item, column, '')).filter(val => val !== '')); return ['ALL', ...Array.from(values).sort()]; };
    const setOptions = (selectElement, options, disable = false) => { selectElement.innerHTML = ''; options.forEach(optionValue => { const option = document.createElement('option'); option.value = optionValue; option.textContent = optionValue === 'ALL' ? `-- ${optionValue} --` : optionValue; option.title = optionValue; selectElement.appendChild(option); }); selectElement.disabled = disable; };
    const setMultiSelectOptions = (selectElement, options, disable = false) => { selectElement.innerHTML = ''; options.forEach(optionValue => { if (optionValue === 'ALL') return; const option = document.createElement('option'); option.value = optionValue; option.textContent = optionValue; option.title = optionValue; selectElement.appendChild(option); }); selectElement.disabled = disable; };
    const showLoading = (isLoading, isFiltering = false) => { if (isFiltering) { if (filterLoadingIndicator) filterLoadingIndicator.style.display = isLoading ? 'flex' : 'none'; if (applyFiltersButton) applyFiltersButton.disabled = isLoading; } else { if (loadingIndicator) loadingIndicator.style.display = isLoading ? 'flex' : 'none'; if (excelFileInput) excelFileInput.disabled = isLoading; } };

    // --- Theme Handling Functions (Defined before use) ---
    const applyTheme = (theme) => {
        if (!bodyElement || !themeToggleButton) return;
        if (theme === 'light') {
            bodyElement.classList.add('light-mode'); bodyElement.classList.remove('dark-mode');
            themeToggleButton.textContent = '🌙'; themeToggleButton.title = 'Switch to Dark Mode';
        } else {
            bodyElement.classList.add('dark-mode'); bodyElement.classList.remove('light-mode');
            themeToggleButton.textContent = '☀️'; themeToggleButton.title = 'Switch to Light Mode';
        }
        currentTheme = theme; localStorage.setItem(THEME_KEY, theme);
        if (mainChartInstance) { updateChartTheme(mainChartInstance); mainChartInstance.update('none'); } // Update chart immediately, no animation
    };

    const toggleTheme = () => {
        const newTheme = currentTheme === 'light' ? 'dark' : 'light';
        applyTheme(newTheme);
    };

    const updateChartTheme = (chart) => {
        const computedStyle = getComputedStyle(bodyElement);
        const gridColor = computedStyle.getPropertyValue('--chart-grid-color').trim();
        const tickColor = computedStyle.getPropertyValue('--chart-tick-color').trim();
        const labelColor = computedStyle.getPropertyValue('--chart-label-color').trim();
        if (chart?.options?.scales?.y) { chart.options.scales.y.grid.color = gridColor; chart.options.scales.y.ticks.color = tickColor; }
        if (chart?.options?.scales?.x) { chart.options.scales.x.grid.color = gridColor; chart.options.scales.x.ticks.color = tickColor; }
        if (chart?.options?.plugins?.legend?.labels) { chart.options.plugins.legend.labels.color = labelColor; }
    };

    // --- Core Functions ---
    const handleFile = async (event) => {
        const file = event.target.files[0]; if (!file) { statusDiv.textContent = 'No file selected.'; return; }
        statusDiv.textContent = 'Reading file...'; showLoading(true); filterArea.style.display = 'none'; resultsArea.style.display = 'none'; resetFilters();
        try {
            const data = await file.arrayBuffer(); const workbook = XLSX.read(data); const firstSheetName = workbook.SheetNames[0]; const worksheet = workbook.Sheets[firstSheetName]; const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: null });
            if (jsonData.length > 0) { const headers = Object.keys(jsonData[0]); const missingHeaders = REQUIRED_HEADERS.filter(h => !headers.includes(h)); if (missingHeaders.length > 0) console.warn(`Warning: Missing columns: ${missingHeaders.join(', ')}.`); } else { throw new Error("Excel sheet appears empty."); }
            rawData = jsonData; allPossibleStores = [...new Set(rawData.map(r => safeGet(r, 'Store', null)).filter(s => s))].sort().map(s => ({ value: s, text: s }));
            statusDiv.textContent = `Loaded ${rawData.length} rows. Adjust filters and click 'Apply Filters'.`;
            populateFilters(rawData); filterArea.style.display = 'block';
        } catch (error) { console.error('Error processing file:', error); statusDiv.textContent = `Error: ${error.message}`; rawData = []; allPossibleStores = []; filteredData = []; resetUI();
        } finally { showLoading(false); if (excelFileInput) excelFileInput.value = ''; }
    };

    const populateFilters = (data) => {
        setOptions(regionFilter, getUniqueValues(data, 'REGION')); setOptions(districtFilter, getUniqueValues(data, 'DISTRICT')); setMultiSelectOptions(territoryFilter, getUniqueValues(data, 'Q2 Territory').slice(1)); setOptions(fsmFilter, getUniqueValues(data, 'FSM NAME')); setOptions(channelFilter, getUniqueValues(data, 'CHANNEL')); setOptions(subchannelFilter, getUniqueValues(data, 'SUB_CHANNEL')); setOptions(dealerFilter, getUniqueValues(data, 'DEALER_NAME'));
        Object.values(flagFiltersCheckboxes).forEach(input => { if(input) input.disabled = false });
        storeOptions = [...allPossibleStores]; setStoreFilterOptions(storeOptions, false);
        territorySelectAll.disabled = false; territoryDeselectAll.disabled = false; storeSelectAll.disabled = false; storeDeselectAll.disabled = false; storeSearch.disabled = false; applyFiltersButton.disabled = false;
        addDependencyFilterListeners();
    };

    const addDependencyFilterListeners = () => {
        const handler = updateStoreFilterOptionsBasedOnHierarchy;
        [regionFilter, districtFilter, territoryFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(filter => { if (filter) { filter.removeEventListener('change', handler); filter.addEventListener('change', handler); } });
        Object.values(flagFiltersCheckboxes).forEach(input => { if (input) { input.removeEventListener('change', handler); input.addEventListener('change', handler); } });
    };

    const updateStoreFilterOptionsBasedOnHierarchy = () => {
        if (rawData.length === 0) return;
        const selectedRegion = regionFilter.value; const selectedDistrict = districtFilter.value; const selectedTerritories = Array.from(territoryFilter.selectedOptions).map(opt => opt.value); const selectedFsm = fsmFilter.value; const selectedChannel = channelFilter.value; const selectedSubchannel = subchannelFilter.value; const selectedDealer = dealerFilter.value; const selectedFlags = {}; Object.entries(flagFiltersCheckboxes).forEach(([key, input]) => { if (input && input.checked) selectedFlags[key] = true; });
        const potentiallyValidStoresData = rawData.filter(row => {
            if (selectedRegion !== 'ALL' && safeGet(row, 'REGION', null) !== selectedRegion) return false; if (selectedDistrict !== 'ALL' && safeGet(row, 'DISTRICT', null) !== selectedDistrict) return false; if (selectedTerritories.length > 0 && !selectedTerritories.includes(safeGet(row, 'Q2 Territory', null))) return false; if (selectedFsm !== 'ALL' && safeGet(row, 'FSM NAME', null) !== selectedFsm) return false; if (selectedChannel !== 'ALL' && safeGet(row, 'CHANNEL', null) !== selectedChannel) return false; if (selectedSubchannel !== 'ALL' && safeGet(row, 'SUB_CHANNEL', null) !== selectedSubchannel) return false; if (selectedDealer !== 'ALL' && safeGet(row, 'DEALER_NAME', null) !== selectedDealer) return false;
            for (const flag in selectedFlags) { const flagValue = safeGet(row, flag, 'NO'); if (!(flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1')) return false; }
            return true;
        });
        const validStoreNames = new Set(potentiallyValidStoresData.map(row => safeGet(row, 'Store', null)).filter(Boolean)); storeOptions = Array.from(validStoreNames).sort().map(s => ({ value: s, text: s }));
        const previouslySelectedStores = new Set(Array.from(storeFilter.selectedOptions).map(opt => opt.value)); setStoreFilterOptions(storeOptions, false); filterStoreOptions();
        Array.from(storeFilter.options).forEach(option => { if (previouslySelectedStores.has(option.value)) option.selected = true; }); if (storeFilter.selectedOptions.length === 0) storeFilter.selectedIndex = -1;
    };

    const setStoreFilterOptions = (optionsToShow, disable = true) => {
        const currentSearchTerm = storeSearch.value; storeFilter.innerHTML = ''; optionsToShow.forEach(opt => { const option = document.createElement('option'); option.value = opt.value; option.textContent = opt.text; option.title = opt.text; storeFilter.appendChild(option); });
        storeFilter.disabled = disable; storeSearch.disabled = disable; storeSelectAll.disabled = disable || optionsToShow.length === 0; storeDeselectAll.disabled = disable || optionsToShow.length === 0; storeSearch.value = currentSearchTerm;
    };

    const filterStoreOptions = () => {
        const searchTerm = storeSearch.value.toLowerCase(); const filteredOptions = storeOptions.filter(opt => opt.text.toLowerCase().includes(searchTerm)); const selectedValues = new Set(Array.from(storeFilter.selectedOptions).map(opt => opt.value)); storeFilter.innerHTML = '';
        filteredOptions.forEach(opt => { const option = document.createElement('option'); option.value = opt.value; option.textContent = opt.text; option.title = opt.text; if (selectedValues.has(opt.value)) option.selected = true; storeFilter.appendChild(option); });
        storeSelectAll.disabled = storeFilter.disabled || filteredOptions.length === 0; storeDeselectAll.disabled = storeFilter.disabled || filteredOptions.length === 0;
    };

    const applyFilters = () => {
        showLoading(true, true); resultsArea.style.display = 'none';
        setTimeout(() => {
            try {
                const selectedRegion = regionFilter.value; const selectedDistrict = districtFilter.value; const selectedTerritories = Array.from(territoryFilter.selectedOptions).map(opt => opt.value); const selectedFsm = fsmFilter.value; const selectedChannel = channelFilter.value; const selectedSubchannel = subchannelFilter.value; const selectedDealer = dealerFilter.value; const selectedStores = Array.from(storeFilter.selectedOptions).map(opt => opt.value); const selectedFlags = {}; Object.entries(flagFiltersCheckboxes).forEach(([key, input]) => { if (input && input.checked) selectedFlags[key] = true; });
                filteredData = rawData.filter(row => {
                    if (selectedRegion !== 'ALL' && safeGet(row, 'REGION', null) !== selectedRegion) return false; if (selectedDistrict !== 'ALL' && safeGet(row, 'DISTRICT', null) !== selectedDistrict) return false; if (selectedTerritories.length > 0 && !selectedTerritories.includes(safeGet(row, 'Q2 Territory', null))) return false; if (selectedFsm !== 'ALL' && safeGet(row, 'FSM NAME', null) !== selectedFsm) return false; if (selectedChannel !== 'ALL' && safeGet(row, 'CHANNEL', null) !== selectedChannel) return false; if (selectedSubchannel !== 'ALL' && safeGet(row, 'SUB_CHANNEL', null) !== selectedSubchannel) return false; if (selectedDealer !== 'ALL' && safeGet(row, 'DEALER_NAME', null) !== selectedDealer) return false; if (selectedStores.length > 0 && !selectedStores.includes(safeGet(row, 'Store', null))) return false;
                    for (const flag in selectedFlags) { const flagValue = safeGet(row, flag, 'NO'); if (!(flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1')) return false; } return true;
                });
                updateSummary(filteredData); updateTopBottomTables(filteredData); updateCharts(filteredData); updateAttachRateTable(filteredData);
                if (filteredData.length === 1) { showStoreDetails(filteredData[0]); highlightTableRow(safeGet(filteredData[0], 'Store', null)); } else { hideStoreDetails(); }
                statusDiv.textContent = `Displaying ${filteredData.length} of ${rawData.length} rows based on filters.`; resultsArea.style.display = 'block'; exportCsvButton.disabled = filteredData.length === 0;
            } catch (error) { console.error("Error applying filters:", error); statusDiv.textContent = "Error applying filters. Check console."; filteredData = []; resultsArea.style.display = 'none'; exportCsvButton.disabled = true; updateSummary([]); updateTopBottomTables([]); updateCharts([]); updateAttachRateTable([]); hideStoreDetails();
            } finally { showLoading(false, true); }
        }, 10);
    };

    const resetFilters = () => {
        [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(sel => { if (sel) { sel.value = 'ALL'; sel.disabled = true;} }); if (territoryFilter) { territoryFilter.selectedIndex = -1; territoryFilter.disabled = true; } if (storeFilter) { storeFilter.innerHTML = '<option value="ALL">-- Load File First --</option>'; storeFilter.selectedIndex = -1; storeFilter.disabled = true; } if (storeSearch) { storeSearch.value = ''; storeSearch.disabled = true; }
        storeOptions = []; allPossibleStores = []; Object.values(flagFiltersCheckboxes).forEach(input => { if(input) {input.checked = false; input.disabled = true;} }); if (applyFiltersButton) applyFiltersButton.disabled = true; if (territorySelectAll) territorySelectAll.disabled = true; if (territoryDeselectAll) territoryDeselectAll.disabled = true; if (storeSelectAll) storeSelectAll.disabled = true; if (storeDeselectAll) storeDeselectAll.disabled = true; if (exportCsvButton) exportCsvButton.disabled = true;
        const handler = updateStoreFilterOptionsBasedOnHierarchy; [regionFilter, districtFilter, territoryFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(filter => { if (filter) filter.removeEventListener('change', handler); }); Object.values(flagFiltersCheckboxes).forEach(input => { if (input) input.removeEventListener('change', handler); });
    };

    const resetUI = () => {
        resetFilters(); if (filterArea) filterArea.style.display = 'none'; if (resultsArea) resultsArea.style.display = 'none'; if (mainChartInstance) { mainChartInstance.destroy(); mainChartInstance = null; } if (attachRateTableBody) attachRateTableBody.innerHTML = ''; if (attachRateTableFooter) attachRateTableFooter.innerHTML = ''; if (attachTableStatus) attachTableStatus.textContent = ''; if (topBottomSection) topBottomSection.style.display = 'none'; if (top5TableBody) top5TableBody.innerHTML = ''; if (bottom5TableBody) bottom5TableBody.innerHTML = '';
        hideStoreDetails(); updateSummary([]); if(statusDiv) statusDiv.textContent = 'No file selected.';
    };

    const updateSummary = (data) => {
        const totalCount = data.length;
        const fieldsToClear = [revenueWithDFValue, qtdRevenueTargetValue, qtdGapValue, quarterlyRevenueTargetValue, percentQuarterlyStoreTargetValue, revARValue, unitsWithDFValue, unitTargetValue, unitAchievementValue, visitCountValue, trainingCountValue, retailModeConnectivityValue, repSkillAchValue, vPmrAchValue, postTrainingScoreValue, eliteValue, percentQuarterlyTerritoryTargetValue, territoryRevPercentValue, districtRevPercentValue, regionRevPercentValue]; fieldsToClear.forEach(el => { if (el) el.textContent = 'N/A'; }); [percentQuarterlyTerritoryTargetP, territoryRevPercentP, districtRevPercentP, regionRevPercentP].forEach(p => { if(p) p.style.display = 'none';}); if (totalCount === 0) { return; }
        const sumRevenue = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Revenue w/DF', 0)), 0); const sumQtdTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'QTD Revenue Target', 0)), 0); const sumQuarterlyTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Quarterly Revenue Target', 0)), 0); const sumUnits = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Unit w/ DF', 0)), 0); const sumUnitTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Unit Target', 0)), 0); const sumVisits = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Visit count', 0)), 0); const sumTrainings = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Trainings', 0)), 0);
        let sumConnectivity = 0, countConnectivity = 0; let sumRepSkill = 0, countRepSkill = 0; let sumPmr = 0, countPmr = 0; let sumPostTraining = 0, countPostTraining = 0; let sumElite = 0, countElite = 0;
        data.forEach(row => { let val; val = safeGet(row, 'Retail Mode Connectivity', null); if (isValidAndNonZeroForAverage(val)) { sumConnectivity += parsePercent(val); countConnectivity++; } val = safeGet(row, 'Rep Skill Ach', null); if (isValidAndNonZeroForAverage(val)) { sumRepSkill += parsePercent(val); countRepSkill++; } val = safeGet(row, '(V)PMR Ach', null); if (isValidAndNonZeroForAverage(val)) { sumPmr += parsePercent(val); countPmr++; } val = safeGet(row, 'Post Training Score', null); if (isValidForAverage(val)) { sumPostTraining += parseNumber(val); countPostTraining++; } val = safeGet(row, 'Elite', null); if (isValidForAverage(val)) { sumElite += parsePercent(val); countElite++; } });
        const avgConnectivity = countConnectivity === 0 ? NaN : sumConnectivity / countConnectivity; const avgRepSkill = countRepSkill === 0 ? NaN : sumRepSkill / countRepSkill; const avgPmr = countPmr === 0 ? NaN : sumPmr / countPmr; const avgPostTraining = countPostTraining === 0 ? NaN : sumPostTraining / countPostTraining; const avgElite = countElite === 0 ? NaN : sumElite / countElite;
        const calculatedRevAR = sumQtdTarget === 0 ? NaN : sumRevenue / sumQtdTarget; const overallUnitAchievement = sumUnitTarget === 0 ? NaN : sumUnits / sumUnitTarget; const overallPercentStoreTarget = sumQuarterlyTarget === 0 ? NaN : sumRevenue / sumQuarterlyTarget;
        // Update DOM...
        revenueWithDFValue && (revenueWithDFValue.textContent = formatCurrency(sumRevenue)); revenueWithDFValue && (revenueWithDFValue.title = `Sum of 'Revenue w/DF' for ${totalCount} filtered stores`); qtdRevenueTargetValue && (qtdRevenueTargetValue.textContent = formatCurrency(sumQtdTarget)); qtdRevenueTargetValue && (qtdRevenueTargetValue.title = `Sum of 'QTD Revenue Target' for ${totalCount} filtered stores`); qtdGapValue && (qtdGapValue.textContent = formatCurrency(sumRevenue - sumQtdTarget)); qtdGapValue && (qtdGapValue.title = `Calculated Gap (Total Revenue - QTD Target) for ${totalCount} filtered stores`); quarterlyRevenueTargetValue && (quarterlyRevenueTargetValue.textContent = formatCurrency(sumQuarterlyTarget)); quarterlyRevenueTargetValue && (quarterlyRevenueTargetValue.title = `Sum of 'Quarterly Revenue Target' for ${totalCount} filtered stores`); unitsWithDFValue && (unitsWithDFValue.textContent = formatNumber(sumUnits)); unitsWithDFValue && (unitsWithDFValue.title = `Sum of 'Unit w/ DF' for ${totalCount} filtered stores`); unitTargetValue && (unitTargetValue.textContent = formatNumber(sumUnitTarget)); unitTargetValue && (unitTargetValue.title = `Sum of 'Unit Target' for ${totalCount} filtered stores`); visitCountValue && (visitCountValue.textContent = formatNumber(sumVisits)); visitCountValue && (visitCountValue.title = `Sum of 'Visit count' for ${totalCount} filtered stores`); trainingCountValue && (trainingCountValue.textContent = formatNumber(sumTrainings)); trainingCountValue && (trainingCountValue.title = `Sum of 'Trainings' for ${totalCount} filtered stores`);
        revARValue && (revARValue.textContent = formatPercent(calculatedRevAR)); revARValue && (revARValue.title = `Calculated Rev AR% (Total Revenue / Total QTD Target)`); unitAchievementValue && (unitAchievementValue.textContent = formatPercent(overallUnitAchievement)); unitAchievementValue && (unitAchievementValue.title = `Overall Unit Achievement % (Total Units / Total Unit Target)`); percentQuarterlyStoreTargetValue && (percentQuarterlyStoreTargetValue.textContent = formatPercent(overallPercentStoreTarget)); percentQuarterlyStoreTargetValue && (percentQuarterlyStoreTargetValue.title = `Overall % Quarterly Target (Total Revenue / Total Quarterly Target)`);
        retailModeConnectivityValue && (retailModeConnectivityValue.textContent = formatPercent(avgConnectivity)); retailModeConnectivityValue && (retailModeConnectivityValue.title = `Average 'Retail Mode Connectivity' across ${countConnectivity} stores with non-zero data`); repSkillAchValue && (repSkillAchValue.textContent = formatPercent(avgRepSkill)); repSkillAchValue && (repSkillAchValue.title = `Average 'Rep Skill Ach' across ${countRepSkill} stores with non-zero data`); vPmrAchValue && (vPmrAchValue.textContent = formatPercent(avgPmr)); vPmrAchValue && (vPmrAchValue.title = `Average '(V)PMR Ach' across ${countPmr} stores with non-zero data`); postTrainingScoreValue && (postTrainingScoreValue.textContent = formatNumber(avgPostTraining.toFixed(1))); postTrainingScoreValue && (postTrainingScoreValue.title = `Average 'Post Training Score' across ${countPostTraining} stores with data`); eliteValue && (eliteValue.textContent = formatPercent(avgElite)); eliteValue && (eliteValue.title = `Average 'Elite' score % across ${countElite} stores with data`);
        updateContextualSummary(data);
    };

    const updateContextualSummary = (data) => { /* ... [implementation unchanged] ... */ };

    const updateTopBottomTables = (data) => {
        if (!topBottomSection || !top5TableBody || !bottom5TableBody) return; top5TableBody.innerHTML = ''; bottom5TableBody.innerHTML = ''; const territoriesInData = new Set(data.map(row => safeGet(row, 'Q2 Territory', null)).filter(Boolean)); const isSingleTerritory = territoriesInData.size === 1; if (!isSingleTerritory || data.length === 0) { topBottomSection.style.display = 'none'; return; } topBottomSection.style.display = 'flex';
        // Top 5
        const top5Data = [...data].sort((a, b) => parseNumber(safeGet(b, 'Revenue w/DF', -Infinity)) - parseNumber(safeGet(a, 'Revenue w/DF', -Infinity))).slice(0, TOP_N_TABLES);
        top5Data.forEach(row => { const tr = top5TableBody.insertRow(); const storeName = safeGet(row, 'Store', null); tr.dataset.storeName = storeName; tr.onclick = () => { showStoreDetails(row); highlightTableRow(storeName); }; const revenue = parseNumber(safeGet(row, 'Revenue w/DF', NaN)); const revAR = calculateRevARPercent(row); const unitAch = calculateUnitAchievementPercent(row); const visits = parseNumber(safeGet(row, 'Visit count', NaN)); tr.insertCell().textContent = storeName; tr.insertCell().textContent = formatCurrency(revenue); tr.insertCell().textContent = formatPercent(revAR); tr.insertCell().textContent = formatPercent(unitAch); tr.insertCell().textContent = formatNumber(visits); });
        // Bottom 5
        const bottom5Data = [...data].sort((a, b) => calculateQtdGap(a) - calculateQtdGap(b)).slice(0, TOP_N_TABLES);
        bottom5Data.forEach(row => { const tr = bottom5TableBody.insertRow(); const storeName = safeGet(row, 'Store', null); tr.dataset.storeName = storeName; tr.onclick = () => { showStoreDetails(row); highlightTableRow(storeName); }; const qtdGap = calculateQtdGap(row); const revAR = calculateRevARPercent(row); const unitAch = calculateUnitAchievementPercent(row); const visits = parseNumber(safeGet(row, 'Visit count', NaN)); tr.insertCell().textContent = storeName; tr.insertCell().textContent = formatCurrency(qtdGap === Infinity ? NaN : qtdGap); tr.insertCell().textContent = formatPercent(revAR); tr.insertCell().textContent = formatPercent(unitAch); tr.insertCell().textContent = formatNumber(visits); });
    };

    // ** UPDATED ** updateCharts to call theme update
    const updateCharts = (data) => {
        if (mainChartInstance) { mainChartInstance.destroy(); mainChartInstance = null; }
        if (data.length === 0 || !mainChartCanvas) return;
        const sortedData = [...data].sort((a, b) => parseNumber(safeGet(b, 'Revenue w/DF', 0)) - parseNumber(safeGet(a, 'Revenue w/DF', 0)));
        const chartData = sortedData.slice(0, TOP_N_CHART);
        const labels = chartData.map(row => safeGet(row, 'Store', 'Unknown'));
        const revenueDataSet = chartData.map(row => parseNumber(safeGet(row, 'Revenue w/DF', 0)));
        const targetDataSet = chartData.map(row => parseNumber(safeGet(row, 'QTD Revenue Target', 0)));
        const backgroundColors = chartData.map((_, index) => revenueDataSet[index] >= targetDataSet[index] ? 'rgba(75, 192, 192, 0.6)' : 'rgba(255, 99, 132, 0.6)');
        const borderColors = chartData.map((_, index) => revenueDataSet[index] >= targetDataSet[index] ? 'rgba(75, 192, 192, 1)' : 'rgba(255, 99, 132, 1)');

        mainChartInstance = new Chart(mainChartCanvas, {
             type: 'bar',
             data: { labels: labels, datasets: [ { label: 'Total Revenue (incl. DF)', data: revenueDataSet, backgroundColor: backgroundColors, borderColor: borderColors, borderWidth: 1 }, { label: 'QTD Revenue Target', data: targetDataSet, type: 'line', borderColor: 'rgba(255, 206, 86, 1)', backgroundColor: 'rgba(255, 206, 86, 0.2)', fill: false, tension: 0.1, borderWidth: 2, pointRadius: 3, pointHoverRadius: 5 } ] },
             options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, ticks: { callback: (value) => formatCurrency(value) } }, x: { grid: { display: false } } }, plugins: { legend: {}, tooltip: { callbacks: { label: function(context) { let label = context.dataset.label || ''; if (label) { label += ': '; } if (context.parsed.y !== null) { if (context.dataset.type === 'line' || context.dataset.label.toLowerCase().includes('target')) { label += formatCurrency(context.parsed.y); } else { label += formatCurrency(context.parsed.y); if (chartData && chartData[context.dataIndex]){ const storeData = chartData[context.dataIndex]; const percentTarget = parsePercent(safeGet(storeData, '% Quarterly Revenue Target', 0)); label += ` (${formatPercent(percentTarget)} of Qtr Target)`; } } } return label; } } } },
             onClick: (event, elements) => { if (elements.length > 0) { const index = elements[0].index; const storeName = labels[index]; const storeData = filteredData.find(row => safeGet(row, 'Store', null) === storeName); if (storeData) { showStoreDetails(storeData); highlightTableRow(storeName); } } } }
        });

        // Apply theme colors *after* creating chart
        updateChartTheme(mainChartInstance);
        mainChartInstance.update('none'); // Update chart without animation
    };

    const updateAttachRateTable = (data) => {
        if (!attachRateTableBody || !attachRateTableFooter) return; attachRateTableBody.innerHTML = ''; attachRateTableFooter.innerHTML = ''; const originalDataLength = data.length; if (originalDataLength === 0) { if(attachTableStatus) attachTableStatus.textContent = 'No data to display.'; return; }
        // 1. Sort the original data
        const sortedData = [...data].sort((a, b) => { let valA = safeGet(a, currentSort.column, null); let valB = safeGet(b, currentSort.column, null); if (valA === null && valB === null) return 0; if (valA === null) return currentSort.ascending ? -1 : 1; if (valB === null) return currentSort.ascending ? 1 : -1; const isPercentCol = currentSort.column.includes('Attach Rate') || currentSort.column.includes('%'); const numA = isPercentCol ? parsePercent(valA) : parseNumber(valA); const numB = isPercentCol ? parsePercent(valB) : parseNumber(valB); if (!isNaN(numA) && !isNaN(numB)) { return currentSort.ascending ? numA - numB : numB - numA; } else { valA = String(valA).toLowerCase(); valB = String(valB).toLowerCase(); return currentSort.ascending ? valA.localeCompare(valB) : valB.localeCompare(valA); } });
        // 2. Calculate averages based on the original data
        const averages = {}; const attachAndPercentKeys = ['% Quarterly Revenue Target', ...ATTACH_RATE_KEYS]; // Include % Qtr Target for average calc
        attachAndPercentKeys.forEach(key => { let sum = 0; let count = 0; data.forEach(row => { const val = safeGet(row, key, null); if (isValidForAverage(val)) { sum += parsePercent(val); count++; } }); averages[key] = count === 0 ? NaN : sum / count; });
        // 3. Iterate sorted data and add rows
        sortedData.forEach(row => { const storeName = safeGet(row, 'Store', null); if (storeName) { const tr = document.createElement('tr'); tr.dataset.storeName = storeName; tr.onclick = () => { showStoreDetails(row); highlightTableRow(storeName); }; const columns = [ { key: 'Store', format: (val) => val }, { key: '% Quarterly Revenue Target', format: formatPercent, highlight: true }, { key: 'Tablet Attach Rate', format: formatPercent, highlight: true }, { key: 'PC Attach Rate', format: formatPercent, highlight: true }, { key: 'NC Attach Rate', format: formatPercent, highlight: true }, { key: 'TWS Attach Rate', format: formatPercent, highlight: true }, { key: 'WW Attach Rate', format: formatPercent, highlight: true }, { key: 'ME Attach Rate', format: formatPercent, highlight: true }, { key: 'NCME Attach Rate', format: formatPercent, highlight: true }, ]; columns.forEach(col => { const td = document.createElement('td'); const rawValue = safeGet(row, col.key, null); const isPercentCol = col.key.includes('Attach Rate') || col.key.includes('%'); const numericValue = isPercentCol ? parsePercent(rawValue) : parseNumber(rawValue); let formattedValue = (rawValue === null || (col.key !== 'Store' && isNaN(numericValue))) ? 'N/A' : (isPercentCol ? col.format(numericValue) : rawValue); td.textContent = formattedValue; td.title = `${col.key}: ${formattedValue}`; if (col.highlight && averages[col.key] !== undefined && !isNaN(numericValue)) { td.classList.add(numericValue >= averages[col.key] ? 'highlight-green' : 'highlight-red'); } tr.appendChild(td); }); attachRateTableBody.appendChild(tr); } });
        // 4. Add Average Row
        if (originalDataLength > 0) { const footerRow = attachRateTableFooter.insertRow(); const avgLabelCell = footerRow.insertCell(); avgLabelCell.textContent = 'Filtered Avg*'; avgLabelCell.title = 'Average calculated only using stores with valid data'; avgLabelCell.style.textAlign = "right"; avgLabelCell.style.fontWeight = "bold"; attachAndPercentKeys.forEach(key => { const td = footerRow.insertCell(); const avgValue = averages[key]; td.textContent = formatPercent(avgValue); let validCount = 0; data.forEach(row => { if (isValidForAverage(safeGet(row, key, null))) validCount++; }); td.title = `Average ${key}: ${formatPercent(avgValue)} (from ${validCount} stores)`; td.style.textAlign = "right"; }); }
        // 5. Update status
        if(attachTableStatus) attachTableStatus.textContent = `Showing ${attachRateTableBody.rows.length} stores. Click row for details. Click headers to sort.`; updateSortArrows();
    };

    const handleSort = (event) => { /* ... [implementation unchanged] ... */ };
    const updateSortArrows = () => { /* ... [implementation unchanged] ... */ };
    const showStoreDetails = (storeData) => { /* ... [implementation unchanged] ... */ };
    const hideStoreDetails = () => { /* ... [implementation unchanged] ... */ };
    const highlightTableRow = (storeName) => { /* ... [implementation unchanged] ... */ };
    const exportData = () => { /* ... [implementation unchanged] ... */ };
    const generateEmailBody = () => { /* ... [implementation unchanged] ... */ };
    const getFilterSummary = () => { /* ... [implementation unchanged] ... */ };
    const handleShareEmail = () => { /* ... [implementation unchanged] ... */ };
    const selectAllOptions = (selectElement) => { /* ... [implementation unchanged] ... */ };
    const deselectAllOptions = (selectElement) => { /* ... [implementation unchanged] ... */ };

    // --- Event Listeners ---
    excelFileInput?.addEventListener('change', handleFile);
    applyFiltersButton?.addEventListener('click', applyFilters);
    storeSearch?.addEventListener('input', filterStoreOptions);
    exportCsvButton?.addEventListener('click', exportData);
    shareEmailButton?.addEventListener('click', handleShareEmail);
    closeStoreDetailsButton?.addEventListener('click', hideStoreDetails);
    territorySelectAll?.addEventListener('click', () => { selectAllOptions(territoryFilter); updateStoreFilterOptionsBasedOnHierarchy(); });
    territoryDeselectAll?.addEventListener('click', () => { deselectAllOptions(territoryFilter); updateStoreFilterOptionsBasedOnHierarchy(); });
    storeSelectAll?.addEventListener('click', () => selectAllOptions(storeFilter));
    storeDeselectAll?.addEventListener('click', () => deselectAllOptions(storeFilter));
    attachRateTable?.querySelector('thead')?.addEventListener('click', handleSort);
    themeToggleButton?.addEventListener('click', toggleTheme); // Add listener

    // --- Initial Setup ---
    resetUI(); // Ensure clean state on load
    applyTheme(currentTheme); // Apply saved or default theme ON LOAD

}); // End DOMContentLoaded
