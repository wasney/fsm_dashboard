/* Generated: 2025-04-07 10:23:55 PM EDT - Comment out theme toggle code and add logging to debug broken file upload functionality. */
document.addEventListener('DOMContentLoaded', () => {
    // --- Configuration ---
    const REQUIRED_HEADERS = [
        'Store', 'REGION', 'DISTRICT', 'Q2 Territory', 'FSM NAME', 'CHANNEL',
        'SUB_CHANNEL', 'DEALER_NAME', 'Revenue w/DF', 'QTD Revenue Target',
        'Quarterly Revenue Target', 'QTD Gap', '% Quarterly Revenue Target',
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
    const CHART_COLORS = ['#58a6ff', '#ffb758', '#86dc86', '#ff7f7f', '#b796e6', '#ffda8a', '#8ad7ff', '#ff9ba6'];
    const TOP_N_CHART = 15;
    const TOP_N_TABLES = 5;
    // const THEME_KEY = 'fsmDashboardTheme'; // Theme code commented out

    // --- DOM Elements ---
    // const bodyElement = document.body; // Theme code commented out
    // const themeToggleButton = document.getElementById('themeToggle'); // Theme code commented out
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
    // let currentTheme = localStorage.getItem(THEME_KEY) || 'dark'; // Theme code commented out

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

    // --- Theme Handling Functions (Commented Out) ---
    /*
    const applyTheme = (theme) => {
        if (!bodyElement || !themeToggleButton) return;
        if (theme === 'light') { bodyElement.classList.add('light-mode'); bodyElement.classList.remove('dark-mode'); themeToggleButton.textContent = '🌙'; themeToggleButton.title = 'Switch to Dark Mode'; }
        else { bodyElement.classList.add('dark-mode'); bodyElement.classList.remove('light-mode'); themeToggleButton.textContent = '☀️'; themeToggleButton.title = 'Switch to Light Mode'; }
        currentTheme = theme; localStorage.setItem(THEME_KEY, theme);
        if (mainChartInstance) { updateChartTheme(mainChartInstance); mainChartInstance.update('none'); }
    };
    const toggleTheme = () => { const newTheme = currentTheme === 'light' ? 'dark' : 'light'; applyTheme(newTheme); };
    const updateChartTheme = (chart) => {
        const computedStyle = getComputedStyle(bodyElement); const gridColor = computedStyle.getPropertyValue('--chart-grid-color').trim(); const tickColor = computedStyle.getPropertyValue('--chart-tick-color').trim(); const labelColor = computedStyle.getPropertyValue('--chart-label-color').trim();
        if (chart?.options?.scales?.y) { chart.options.scales.y.grid.color = gridColor; chart.options.scales.y.ticks.color = tickColor; }
        if (chart?.options?.scales?.x) { chart.options.scales.x.grid.color = gridColor; chart.options.scales.x.ticks.color = tickColor; }
        if (chart?.options?.plugins?.legend?.labels) { chart.options.plugins.legend.labels.color = labelColor; }
    };
    */

    // --- Core Functions ---
    const handleFile = async (event) => {
        console.log("handleFile started"); // DEBUG LOG ADDED
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

    const populateFilters = (data) => { /* ... [implementation unchanged] ... */ };
    const addDependencyFilterListeners = () => { /* ... [implementation unchanged] ... */ };
    const updateStoreFilterOptionsBasedOnHierarchy = () => { /* ... [implementation unchanged] ... */ };
    const setStoreFilterOptions = (optionsToShow, disable = true) => { /* ... [implementation unchanged] ... */ };
    const filterStoreOptions = () => { /* ... [implementation unchanged] ... */ };
    const applyFilters = () => { /* ... [implementation unchanged - includes debug logging from previous attempt] ... */ };
    const resetFilters = () => { /* ... [implementation unchanged] ... */ };
    const resetUI = () => { /* ... [implementation unchanged] ... */ };
    const updateSummary = (data) => { /* ... [implementation unchanged] ... */ };
    const updateContextualSummary = (data) => { /* ... [implementation unchanged] ... */ };
    const updateTopBottomTables = (data) => { /* ... [implementation unchanged] ... */ };

    // ** UPDATED ** updateCharts to call theme update (but theme logic is commented out now)
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

        // Define default colors (dark theme) as theme logic is commented out
        const gridColor = 'rgba(224, 224, 224, 0.2)';
        const tickColor = '#e0e0e0';
        const labelColor = '#e0e0e0';

        mainChartInstance = new Chart(mainChartCanvas, {
             type: 'bar',
             data: { labels: labels, datasets: [ { label: 'Total Revenue (incl. DF)', data: revenueDataSet, backgroundColor: backgroundColors, borderColor: borderColors, borderWidth: 1 }, { label: 'QTD Revenue Target', data: targetDataSet, type: 'line', borderColor: 'rgba(255, 206, 86, 1)', backgroundColor: 'rgba(255, 206, 86, 0.2)', fill: false, tension: 0.1, borderWidth: 2, pointRadius: 3, pointHoverRadius: 5 } ] },
             options: { responsive: true, maintainAspectRatio: false,
                 scales: { y: { beginAtZero: true, ticks: { color: tickColor, callback: (value) => formatCurrency(value) }, grid: { color: gridColor } }, x: { ticks: { color: tickColor }, grid: { display: false } } }, // Apply default colors
                 plugins: { legend: { labels : { color: labelColor } }, // Apply default color
                 tooltip: { callbacks: { label: function(context) { let label = context.dataset.label || ''; if (label) { label += ': '; } if (context.parsed.y !== null) { if (context.dataset.type === 'line' || context.dataset.label.toLowerCase().includes('target')) { label += formatCurrency(context.parsed.y); } else { label += formatCurrency(context.parsed.y); if (chartData && chartData[context.dataIndex]){ const storeData = chartData[context.dataIndex]; const percentTarget = parsePercent(safeGet(storeData, '% Quarterly Revenue Target', 0)); label += ` (${formatPercent(percentTarget)} of Qtr Target)`; } } } return label; } } } },
             onClick: (event, elements) => { if (elements.length > 0) { const index = elements[0].index; const storeName = labels[index]; const storeData = filteredData.find(row => safeGet(row, 'Store', null) === storeName); if (storeData) { showStoreDetails(storeData); highlightTableRow(storeName); } } } }
        });
        // No theme update needed here as logic is commented out
        // updateChartTheme(mainChartInstance);
        // mainChartInstance.update('none');
    };

    const updateAttachRateTable = (data) => { /* ... [implementation unchanged] ... */ };
    const handleSort = (event) => { /* ... [implementation unchanged] ... */ };
    const updateSortArrows = () => { /* ... [implementation unchanged] ... */ };
    const showStoreDetails = (storeData) => { /* ... [implementation unchanged - includes debug logging] ... */ };
    const hideStoreDetails = () => { /* ... [implementation unchanged - includes debug logging] ... */ };
    const highlightTableRow = (storeName) => { /* ... [implementation unchanged - includes debug logging] ... */ };
    const exportData = () => { /* ... [implementation unchanged] ... */ };
    const generateEmailBody = () => { /* ... [implementation unchanged] ... */ };
    const getFilterSummary = () => { /* ... [implementation unchanged] ... */ };
    const handleShareEmail = () => { /* ... [implementation unchanged] ... */ };
    const selectAllOptions = (selectElement) => { /* ... [implementation unchanged] ... */ };
    const deselectAllOptions = (selectElement) => { /* ... [implementation unchanged] ... */ };

    // --- Event Listeners ---
    excelFileInput?.addEventListener('change', handleFile);
    console.log("File input listener attached"); // DEBUG LOG ADDED
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
    // themeToggleButton?.addEventListener('click', toggleTheme); // Theme listener commented out

    // --- Initial Setup ---
    resetUI(); // Ensure clean state on load
    // applyTheme(currentTheme); // Theme initial application commented out

}); // End DOMContentLoaded
