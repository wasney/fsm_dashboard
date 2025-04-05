/* Generated: 2025-04-05 12:13:21 PM EDT - Add dark/light theme toggle with localStorage persistence and chart color adaptation. */
document.addEventListener('DOMContentLoaded', () => {
    // --- Configuration ---
    const REQUIRED_HEADERS = [ // Add all essential headers needed for calculations/display
        'Store', 'REGION', 'DISTRICT', 'Q2 Territory', 'FSM NAME', 'CHANNEL',
        'SUB_CHANNEL', 'DEALER_NAME', 'Revenue w/DF', 'QTD Revenue Target',
        'Quarterly Revenue Target', 'QTD Gap', '% Quarterly Revenue Target', 'Rev AR%',
        'Unit w/ DF', 'Unit Target', 'Unit Achievement', 'Visit count', 'Trainings',
        'Retail Mode Connectivity', 'Rep Skill Ach', '(V)PMR Ach', 'Elite', 'Post Training Score',
        'Tablet Attach Rate', 'PC Attach Rate', 'NC Attach Rate', 'TWS Attach Rate',
        'WW Attach Rate', 'ME Attach Rate', 'NCME Attach Rate', 'SUPER STORE', 'GOLDEN RHINO',
        'GCE', 'AI_Zone', 'Hispanic_Market', 'EV ROUTE',
        // Store Details Headers:
        'STORE ID', 'ADDRESS1', 'CITY', 'STATE', 'ZIPCODE',
        'LATITUDE_ORG', 'LONGITUDE_ORG', // For Google Maps
        'ORG_STORE_ID', 'CV_STORE_ID', 'CINGLEPOINT_ID', // Additional IDs
        'STORE_TYPE_NAME', 'National_Tier', 'Merchandising_Level', 'Combined_Tier', // Store Type/Tier info
        // Context headers if needed for display/logic
        '%Quarterly Territory Rev Target', 'Region Rev%', 'District Rev%', 'Territory Rev%'
    ];
    const FLAG_HEADERS = ['SUPER STORE', 'GOLDEN RHINO', 'GCE', 'AI_Zone', 'Hispanic_Market', 'EV ROUTE']; // Used for Flag summary in details
    const CURRENCY_FORMAT = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' });
    const PERCENT_FORMAT = new Intl.NumberFormat('en-US', { style: 'percent', minimumFractionDigits: 1, maximumFractionDigits: 1 });
    const NUMBER_FORMAT = new Intl.NumberFormat('en-US');
    const CHART_COLORS = ['#58a6ff', '#ffb758', '#86dc86', '#ff7f7f', '#b796e6', '#ffda8a', '#8ad7ff', '#ff9ba6'];
    const TOP_N_CHART = 15; // Max items to show on the bar chart
    const THEME_KEY = 'fsmDashboardTheme'; // Key for localStorage

    // --- DOM Elements ---
    const bodyElement = document.body; // Reference to body
    const themeToggleButton = document.getElementById('themeToggle');
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
    let currentSort = { column: 'Store', ascending: true };
    let selectedStoreRow = null;
    let currentTheme = localStorage.getItem(THEME_KEY) || 'dark'; // Default to dark

    // --- Helper Functions ---
    const formatCurrency = (value) => isNaN(value) ? 'N/A' : CURRENCY_FORMAT.format(value);
    const formatPercent = (value) => isNaN(value) ? 'N/A' : PERCENT_FORMAT.format(value);
    const formatNumber = (value) => isNaN(value) ? 'N/A' : NUMBER_FORMAT.format(value);

    const parseNumber = (value) => { /* ... [implementation unchanged] ... */ };
    const parsePercent = (value) => { /* ... [implementation unchanged] ... */ };
    const safeGet = (obj, path, defaultValue = 'N/A') => { /* ... [implementation unchanged] ... */ };
    const isValidForAverage = (value) => { /* ... [implementation unchanged] ... */ };
    const getUniqueValues = (data, column) => { /* ... [implementation unchanged] ... */ };
    const setOptions = (selectElement, options, disable = false) => { /* ... [implementation unchanged] ... */ };
    const setMultiSelectOptions = (selectElement, options, disable = false) => { /* ... [implementation unchanged] ... */ };
    const showLoading = (isLoading, isFiltering = false) => { /* ... [implementation unchanged] ... */ };

    // --- Theme Handling ---
    const applyTheme = (theme) => {
        if (theme === 'light') {
            bodyElement.classList.add('light-mode');
            bodyElement.classList.remove('dark-mode');
            themeToggleButton.textContent = '🌙'; // Moon for switching to dark
            themeToggleButton.title = 'Switch to Dark Mode';
        } else {
            bodyElement.classList.add('dark-mode');
            bodyElement.classList.remove('light-mode');
            themeToggleButton.textContent = '☀️'; // Sun for switching to light
            themeToggleButton.title = 'Switch to Light Mode';
        }
        currentTheme = theme; // Update global state
        localStorage.setItem(THEME_KEY, theme); // Save preference

        // Update chart if it exists
        if (mainChartInstance) {
            updateChartTheme(mainChartInstance);
            mainChartInstance.update(); // Redraw chart with new colors
        }
    };

    const toggleTheme = () => {
        const newTheme = bodyElement.classList.contains('light-mode') ? 'dark' : 'light';
        applyTheme(newTheme);
    };

    const updateChartTheme = (chart) => {
        const isDarkMode = currentTheme === 'dark';
        const gridColor = isDarkMode ? 'rgba(224, 224, 224, 0.2)' : 'rgba(0, 0, 0, 0.1)';
        const tickColor = isDarkMode ? '#e0e0e0' : '#666';
        const labelColor = isDarkMode ? '#e0e0e0' : '#333';

        // Update options dynamically
        if (chart && chart.options && chart.options.scales) {
            if (chart.options.scales.y) {
                chart.options.scales.y.grid = chart.options.scales.y.grid || {};
                chart.options.scales.y.grid.color = gridColor;
                chart.options.scales.y.ticks = chart.options.scales.y.ticks || {};
                chart.options.scales.y.ticks.color = tickColor;
            }
            if (chart.options.scales.x) {
                 chart.options.scales.x.grid = chart.options.scales.x.grid || {};
                 chart.options.scales.x.grid.color = gridColor; // Can set display: false if needed
                chart.options.scales.x.ticks = chart.options.scales.x.ticks || {};
                chart.options.scales.x.ticks.color = tickColor;
            }
        }
         if (chart && chart.options && chart.options.plugins && chart.options.plugins.legend) {
             chart.options.plugins.legend.labels = chart.options.plugins.legend.labels || {};
             chart.options.plugins.legend.labels.color = labelColor;
         }
          // Tooltip colors are often handled by Chart.js defaults or might need specific styling if customized heavily
    };


    // --- Core Functions ---
    const handleFile = async (event) => { /* ... [implementation unchanged] ... */ };
    const populateFilters = (data) => { /* ... [implementation unchanged] ... */ };
    const addDependencyFilterListeners = () => { /* ... [implementation unchanged] ... */ };
    const updateStoreFilterOptionsBasedOnHierarchy = () => { /* ... [implementation unchanged] ... */ };
    const setStoreFilterOptions = (optionsToShow, disable = true) => { /* ... [implementation unchanged] ... */ };
    const filterStoreOptions = () => { /* ... [implementation unchanged] ... */ };
    const applyFilters = () => { /* ... [implementation unchanged] ... */ };
    const resetFilters = () => { /* ... [implementation unchanged] ... */ };
    const resetUI = () => { /* ... [implementation unchanged] ... */ };
    const updateSummary = (data) => { /* ... [implementation unchanged] ... */ };
    const updateContextualSummary = (data) => { /* ... [implementation unchanged] ... */ };

    // ** UPDATED ** Chart Update Function
    const updateCharts = (data) => {
        if (mainChartInstance) {
            mainChartInstance.destroy();
            mainChartInstance = null;
        }

        if (data.length === 0 || !mainChartCanvas) return;

        // Sort data and prepare chart datasets (unchanged)
        const sortedData = [...data].sort((a, b) => parseNumber(safeGet(b, 'Revenue w/DF', 0)) - parseNumber(safeGet(a, 'Revenue w/DF', 0)));
        const chartData = sortedData.slice(0, TOP_N_CHART);
        const labels = chartData.map(row => safeGet(row, 'Store', 'Unknown'));
        const revenueDataSet = chartData.map(row => parseNumber(safeGet(row, 'Revenue w/DF', 0)));
        const targetDataSet = chartData.map(row => parseNumber(safeGet(row, 'QTD Revenue Target', 0)));
        const backgroundColors = chartData.map((_, index) => revenueDataSet[index] >= targetDataSet[index] ? 'rgba(75, 192, 192, 0.6)' : 'rgba(255, 99, 132, 0.6)');
        const borderColors = chartData.map((_, index) => revenueDataSet[index] >= targetDataSet[index] ? 'rgba(75, 192, 192, 1)' : 'rgba(255, 99, 132, 1)');

        // Determine theme-specific colors
        const isDarkMode = currentTheme === 'dark';
        const gridColor = isDarkMode ? 'rgba(224, 224, 224, 0.2)' : 'rgba(0, 0, 0, 0.1)';
        const tickColor = isDarkMode ? '#e0e0e0' : '#666';
        const labelColor = isDarkMode ? '#e0e0e0' : '#333';

        mainChartInstance = new Chart(mainChartCanvas, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [
                    { label: 'Total Revenue (incl. DF)', data: revenueDataSet, backgroundColor: backgroundColors, borderColor: borderColors, borderWidth: 1 },
                    { label: 'QTD Revenue Target', data: targetDataSet, type: 'line', borderColor: 'rgba(255, 206, 86, 1)', backgroundColor: 'rgba(255, 206, 86, 0.2)', fill: false, tension: 0.1, borderWidth: 2, pointRadius: 3, pointHoverRadius: 5 }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: { color: tickColor, callback: (value) => formatCurrency(value) }, // Use theme color
                        grid: { color: gridColor } // Use theme color
                    },
                    x: {
                        ticks: { color: tickColor }, // Use theme color
                        grid: { display: false } // Keep grid off for x-axis
                    }
                },
                plugins: {
                    legend: { labels: { color: labelColor } }, // Use theme color
                    tooltip: {
                         // You might need to adjust tooltip colors explicitly if default doesn't adapt well
                         // titleColor: '...', bodyColor: '...', etc.
                         callbacks: {
                             label: function(context) { /* ... [unchanged tooltip label logic] ... */ }
                         }
                    }
                },
                onClick: (event, elements) => { /* ... [unchanged click logic] ... */ }
            }
        });
    };


    const updateAttachRateTable = (data) => { /* ... [implementation unchanged from previous step] ... */ };
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
    themeToggleButton?.addEventListener('click', toggleTheme); // ** NEW ** Theme toggle listener

    // --- Initial Setup ---
    resetUI(); // Ensure clean state on load
    applyTheme(currentTheme); // Apply saved or default theme on load

}); // End DOMContentLoaded

// --- Helper Functions Implementation (Copied from previous version - Assuming no changes needed) ---
const parseNumber = (value) => { if (value === null || value === undefined || value === '') return NaN; if (typeof value === 'number') return value; if (typeof value === 'string') { value = value.replace(/[$,%]/g, ''); const num = parseFloat(value); return isNaN(num) ? NaN : num; } return NaN; };
const parsePercent = (value) => { if (value === null || value === undefined || value === '') return NaN; if (typeof value === 'number') return value; if (typeof value === 'string') { const num = parseFloat(value.replace('%', '')); return isNaN(num) ? NaN : num / 100; } return NaN; };
const safeGet = (obj, path, defaultValue = 'N/A') => { const value = obj ? obj[path] : undefined; return (value !== undefined && value !== null) ? value : defaultValue; };
const isValidForAverage = (value) => { if (value === null || value === undefined || value === '') return false; return !isNaN(parseNumber(String(value).replace('%',''))); };
const getUniqueValues = (data, column) => { const values = new Set(data.map(item => safeGet(item, column, '')).filter(val => val !== '')); return ['ALL', ...Array.from(values).sort()]; };
const setOptions = (selectElement, options, disable = false) => { selectElement.innerHTML = ''; options.forEach(optionValue => { const option = document.createElement('option'); option.value = optionValue; option.textContent = optionValue === 'ALL' ? `-- ${optionValue} --` : optionValue; option.title = optionValue; selectElement.appendChild(option); }); selectElement.disabled = disable; };
const setMultiSelectOptions = (selectElement, options, disable = false) => { selectElement.innerHTML = ''; options.forEach(optionValue => { if (optionValue === 'ALL') return; const option = document.createElement('option'); option.value = optionValue; option.textContent = optionValue; option.title = optionValue; selectElement.appendChild(option); }); selectElement.disabled = disable; };
const showLoading = (isLoading, isFiltering = false) => { const loadingIndicator = document.getElementById('loadingIndicator'); const filterLoadingIndicator = document.getElementById('filterLoadingIndicator'); const applyFiltersButton = document.getElementById('applyFiltersButton'); const excelFileInput = document.getElementById('excelFile'); if (isFiltering) { if (filterLoadingIndicator) filterLoadingIndicator.style.display = isLoading ? 'flex' : 'none'; if (applyFiltersButton) applyFiltersButton.disabled = isLoading; } else { if (loadingIndicator) loadingIndicator.style.display = isLoading ? 'flex' : 'none'; if (excelFileInput) excelFileInput.disabled = isLoading; } };
