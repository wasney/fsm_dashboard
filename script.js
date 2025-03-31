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
    let chartOptionsConfig;

    const TOP_N_CHART = 15; // Max items to show on the bar chart

    // --- DOM Elements ---
    const excelFileInput = document.getElementById('excelFile');
    const statusDiv = document.getElementById('status');
    const loadingIndicator = document.getElementById('loadingIndicator');
    const filterLoadingIndicator = document.getElementById('filterLoadingIndicator');
    const filterArea = document.getElementById('filterArea');
    const resultsArea = document.getElementById('resultsArea');
    const applyFiltersButton = document.getElementById('applyFiltersButton');
    const resetFiltersButton = document.getElementById('resetFiltersButton');
    const themeToggleButton = document.getElementById('themeToggleButton');
    const themeMetaTag = document.querySelector('meta[name="theme-color"]');

    // Filter Elements
    const regionFilter = document.getElementById('regionFilter');
    const districtFilter = document.getElementById('districtFilter');
    const territoryFilterContainer = document.getElementById('territoryFilterContainer');
    const storeFilterContainer = document.getElementById('storeFilterContainer');
    const fsmFilter = document.getElementById('fsmFilter');
    const channelFilter = document.getElementById('channelFilter');
    const subchannelFilter = document.getElementById('subchannelFilter');
    const dealerFilter = document.getElementById('dealerFilter');
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
            default: console.warn(`Unknown flag header encountered: ${header}`); return acc;
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
    // const secondaryChartCanvas = document.getElementById('secondaryChartCanvas').getContext('2d'); // Placeholder

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
    // let secondaryChartInstance = null; // Placeholder
    let storeOptions = [];
    let allPossibleStores = [];
    let currentSort = { column: 'Store', ascending: true };
    let selectedStoreRow = null;

    // --- Helper Functions ---
    const formatCurrency = (value) => isNaN(value) ? 'N/A' : CURRENCY_FORMAT.format(value);
    const formatPercent = (value) => isNaN(value) ? 'N/A' : PERCENT_FORMAT.format(value);
    const formatNumber = (value) => isNaN(value) ? 'N/A' : NUMBER_FORMAT.format(value);

    const parseNumber = (value) => {
        if (value === null || value === undefined || value === '') return NaN;
        if (typeof value === 'number') return value;
        if (typeof value === 'string') {
            value = value.replace(/[$,%]/g, '');
            const num = parseFloat(value);
            return isNaN(num) ? NaN : num;
        }
        return NaN;
    };
    const parsePercent = (value) => {
         if (value === null || value === undefined || value === '') return NaN;
         if (typeof value === 'number') return value;
         if (typeof value === 'string') {
            const num = parseFloat(value.replace('%', ''));
            return isNaN(num) ? NaN : num / 100;
         }
         return NaN;
    };
    // Updated safeGet default to null for easier type checking later if needed
    const safeGet = (obj, path, defaultValue = null) => {
        const value = obj ? obj[path] : undefined;
        // Return the value if it's not undefined or null, otherwise return the defaultValue
        return (value !== undefined && value !== null) ? value : defaultValue;
    };

    const isValidForAverage = (value) => {
         if (value === null || value === undefined || value === '') return false;
         return !isNaN(parseNumber(String(value).replace('%','')));
    };
    const getUniqueValues = (data, column) => {
        // Use safeGet with '' as default before putting into Set
        const values = new Set(data.map(item => safeGet(item, column, '')).filter(val => val !== ''));
        return [...Array.from(values).sort()];
    };
    const setOptions = (selectElement, options, disable = false) => {
        selectElement.innerHTML = '';
        const allOption = document.createElement('option');
        allOption.value = 'ALL';
        allOption.textContent = '-- ALL --';
        selectElement.appendChild(allOption);

        options.forEach(optionValue => {
            const option = document.createElement('option');
            option.value = optionValue;
            option.textContent = optionValue;
            option.title = optionValue;
            selectElement.appendChild(option);
        });
         selectElement.disabled = disable;
         selectElement.value = 'ALL';
    };

    const populateCheckboxMultiSelect = (containerElement, options, disable = false) => {
        if (!containerElement) return;
        containerElement.innerHTML = '';
        containerElement.scrollTop = 0;

        if (options.length === 0 && !disable) {
             containerElement.dataset.placeholder = "-- No matching options --";
        } else if (disable) {
            containerElement.dataset.placeholder = "-- Load File First --";
        } else {
             delete containerElement.dataset.placeholder;
             options.forEach(optionValue => {
                const label = document.createElement('label');
                label.className = 'cb-item';
                label.title = optionValue;

                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.value = optionValue;
                checkbox.name = containerElement.id;

                const text = document.createElement('span');
                text.textContent = optionValue;

                label.appendChild(checkbox);
                label.appendChild(text);
                containerElement.appendChild(label);
            });
        }

        containerElement.classList.toggle('disabled', disable);

        const selectAllBtn = document.getElementById(`${containerElement.id.replace('Container', '')}SelectAll`);
        const deselectAllBtn = document.getElementById(`${containerElement.id.replace('Container', '')}DeselectAll`);
        if (selectAllBtn) selectAllBtn.disabled = disable || options.length === 0;
        if (deselectAllBtn) deselectAllBtn.disabled = disable || options.length === 0;
    };


    const showLoading = (isLoading, isFiltering = false) => {
        if (isFiltering) {
            filterLoadingIndicator.style.display = isLoading ? 'flex' : 'none';
            applyFiltersButton.disabled = isLoading || rawData.length === 0;
            resetFiltersButton.disabled = isLoading || rawData.length === 0;
        } else {
            loadingIndicator.style.display = isLoading ? 'flex' : 'none';
            excelFileInput.disabled = isLoading;
        }
    };

    // --- Theme Handling ---
    const applyTheme = (theme) => {
        if (theme === 'light') {
            document.body.classList.add('light-mode');
            themeToggleButton.textContent = '🌙';
            themeToggleButton.title = 'Switch to Dark Mode';
            if (themeMetaTag) themeMetaTag.content = '#f4f4f8';
        } else {
            document.body.classList.remove('light-mode');
            themeToggleButton.textContent = '☀️';
            themeToggleButton.title = 'Switch to Light Mode';
            if (themeMetaTag) themeMetaTag.content = '#1e1e1e';
        }
        chartOptionsConfig = getChartOptions(theme);
        if (mainChartInstance) {
            mainChartInstance.options = chartOptionsConfig;
            mainChartInstance.update();
        }
    };

    const toggleTheme = () => {
        const currentTheme = document.body.classList.contains('light-mode') ? 'light' : 'dark';
        const newTheme = currentTheme === 'light' ? 'dark' : 'light';
        localStorage.setItem('dashboardTheme', newTheme);
        applyTheme(newTheme);
    };

    function getChartOptions(theme = 'dark') {
        const isLightMode = theme === 'light';
        const gridColor = isLightMode ? 'rgba(0, 0, 0, 0.1)' : 'rgba(224, 224, 224, 0.2)';
        const tickColor = isLightMode ? '#6c757d' : '#e0e0e0';
        const labelColor = isLightMode ? '#212529' : '#e0e0e0';
        const barGoodBg = isLightMode ? 'rgba(25, 135, 84, 0.6)' : 'rgba(75, 192, 192, 0.6)';
        const barGoodBorder = isLightMode ? 'rgba(25, 135, 84, 1)' : 'rgba(75, 192, 192, 1)';
        const barBadBg = isLightMode ? 'rgba(220, 53, 69, 0.6)' : 'rgba(255, 99, 132, 0.6)';
        const barBadBorder = isLightMode ? 'rgba(220, 53, 69, 1)' : 'rgba(255, 99, 132, 1)';
        const lineTargetColor = isLightMode ? 'rgba(255, 193, 7, 1)' : 'rgba(255, 206, 86, 1)';
        const lineTargetBg = isLightMode ? 'rgba(255, 193, 7, 0.2)' : 'rgba(255, 206, 86, 0.2)';

        return {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: { color: tickColor, callback: (value) => formatCurrency(value) },
                    grid: { color: gridColor }
                },
                x: {
                    ticks: { color: tickColor },
                    grid: { display: false }
                }
            },
            plugins: {
                legend: { labels: { color: labelColor } },
                tooltip: {
                    backgroundColor: isLightMode ? '#ffffff' : '#2c2c2c',
                    titleColor: isLightMode ? '#212529' : '#e0e0e0',
                    bodyColor: isLightMode ? '#212529' : '#e0e0e0',
                    callbacks: {
                         label: function(context) {
                            let label = context.dataset.label || '';
                            if (label) { label += ': '; }
                            if (context.parsed.y !== null) {
                                 if (context.dataset.type === 'line' || context.dataset.label.toLowerCase().includes('target')) {
                                      label += formatCurrency(context.parsed.y);
                                 } else {
                                      label += formatCurrency(context.parsed.y);
                                       const storeName = context.label;
                                       // Ensure filteredData is accessible or use logic to find the data point
                                       const storeData = filteredData.find(row => safeGet(row, 'Store', null) === storeName);
                                       if(storeData){
                                            const percentTarget = parsePercent(safeGet(storeData, '% Quarterly Revenue Target', 0));
                                            label += ` (${formatPercent(percentTarget)} of Qtr Target)`;
                                       }
                                 }
                            }
                            return label;
                        }
                    }
                }
            },
             onClick: (event, elements) => {
                 if (elements.length > 0) {
                     const index = elements[0].index;
                     const clickedLabel = mainChartInstance?.data?.labels?.[index];
                      if (!clickedLabel) return;
                     const storeData = filteredData.find(row => safeGet(row, 'Store', null) === clickedLabel);
                     if (storeData) {
                         showStoreDetails(storeData);
                         highlightTableRow(clickedLabel);
                     }
                 }
             },
             colorDefaults: { // Store colors here for updateCharts to access
                barGoodBg, barGoodBorder, barBadBg, barBadBorder, lineTargetColor, lineTargetBg
             }
        };
    }


    // --- Core Functions ---

    const handleFile = async (event) => {
        const file = event.target.files[0];
        if (!file) { statusDiv.textContent = 'No file selected.'; return; }

        statusDiv.textContent = 'Reading file...';
        showLoading(true);
        filterArea.style.display = 'none';
        resultsArea.style.display = 'none';
        resetUI(false);

        try {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data);
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: null });

            if (jsonData.length > 0) {
                const headers = Object.keys(jsonData[0]);
                const missingHeaders = REQUIRED_HEADERS.filter(h => !headers.includes(h));
                if (missingHeaders.length > 0) {
                    console.warn(`Warning: Missing expected columns: ${missingHeaders.join(', ')}.`);
                }
            } else { throw new Error("Excel sheet appears to be empty."); }

            rawData = jsonData;
            allPossibleStores = [...new Set(rawData.map(r => safeGet(r, 'Store', null)).filter(s => s !== null))] // Ensure not null before sort
                                 .sort()
                                 .map(s => ({ value: String(s), text: String(s) })); // Ensure values/text are strings

            statusDiv.textContent = `Loaded ${rawData.length} rows. Applying default filters...`;
            populateFilters(rawData);
            filterArea.style.display = 'block';
            resetFiltersButton.disabled = false;

            applyFilters();

        } catch (error) {
            console.error('Error processing file:', error);
            statusDiv.textContent = `Error: ${error.message}`;
            rawData = []; allPossibleStores = []; filteredData = [];
            resetUI(true);
        } finally {
            showLoading(false);
            excelFileInput.value = '';
        }
    };

    const populateFilters = (data) => {
        setOptions(regionFilter, getUniqueValues(data, 'REGION'));
        setOptions(districtFilter, getUniqueValues(data, 'DISTRICT'));
        setOptions(fsmFilter, getUniqueValues(data, 'FSM NAME'));
        setOptions(channelFilter, getUniqueValues(data, 'CHANNEL'));
        setOptions(subchannelFilter, getUniqueValues(data, 'SUB_CHANNEL'));
        setOptions(dealerFilter, getUniqueValues(data, 'DEALER_NAME'));

        // Use String conversion for territory options just in case
        const territoryOptions = getUniqueValues(data, 'Q2 Territory').map(String);
        populateCheckboxMultiSelect(territoryFilterContainer, territoryOptions, false);

        storeOptions = [...allPossibleStores]; // Use pre-stringified values
        populateCheckboxMultiSelect(storeFilterContainer, storeOptions.map(s => s.text), false);

        Object.values(flagFiltersCheckboxes).forEach(input => { if(input) input.disabled = false });
        storeSearch.disabled = false;
        applyFiltersButton.disabled = false;
        resetFiltersButton.disabled = false;
        addDependencyFilterListeners();
    };

    const addDependencyFilterListeners = () => {
        const handler = updateStoreFilterOptionsBasedOnHierarchy;
        [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(filter => {
            if (filter) { filter.removeEventListener('change', handler); filter.addEventListener('change', handler); }
        });
        Object.values(flagFiltersCheckboxes).forEach(input => {
             if (input) { input.removeEventListener('change', handler); input.addEventListener('change', handler); }
        });
        if (territoryFilterContainer) {
             territoryFilterContainer.removeEventListener('change', handler);
             territoryFilterContainer.addEventListener('change', handler);
        }
    };

     const updateStoreFilterOptionsBasedOnHierarchy = () => {
        if (rawData.length === 0) return;

        const selectedRegion = regionFilter.value;
        const selectedDistrict = districtFilter.value;
        const selectedTerritories = territoryFilterContainer
            ? Array.from(territoryFilterContainer.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value)
            : [];
        const selectedFsm = fsmFilter.value;
        const selectedChannel = channelFilter.value;
        const selectedSubchannel = subchannelFilter.value;
        const selectedDealer = dealerFilter.value;
        const selectedFlags = {};
        Object.entries(flagFiltersCheckboxes).forEach(([key, input]) => {
             if (input && input.checked) { selectedFlags[key] = true; }
         });

        const potentiallyValidStoresData = rawData.filter(row => {
             // Explicit String conversion for comparison consistency
            if (selectedRegion !== 'ALL' && String(safeGet(row, 'REGION', '')) !== selectedRegion) return false;
            if (selectedDistrict !== 'ALL' && String(safeGet(row, 'DISTRICT', '')) !== selectedDistrict) return false;
            if (selectedTerritories.length > 0 && !selectedTerritories.includes(String(safeGet(row, 'Q2 Territory', '')))) return false;
            if (selectedFsm !== 'ALL' && String(safeGet(row, 'FSM NAME', '')) !== selectedFsm) return false;
            if (selectedChannel !== 'ALL' && String(safeGet(row, 'CHANNEL', '')) !== selectedChannel) return false;
            if (selectedSubchannel !== 'ALL' && String(safeGet(row, 'SUB_CHANNEL', '')) !== selectedSubchannel) return false;
            if (selectedDealer !== 'ALL' && String(safeGet(row, 'DEALER_NAME', '')) !== selectedDealer) return false;

            for (const flag in selectedFlags) {
                 const flagValue = safeGet(row, flag, 'NO');
                 if (!(flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1')) {
                    return false;
                 }
            }
            return true;
        });

        // Get unique store names (ensure they are strings)
        const validStoreNames = new Set(potentiallyValidStoresData.map(row => String(safeGet(row, 'Store', ''))).filter(Boolean));
        storeOptions = Array.from(validStoreNames).sort().map(s => ({ value: s, text: s })); // Already strings

        const previouslyCheckedStores = storeFilterContainer
            ? new Set(Array.from(storeFilterContainer.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value))
            : new Set();

        populateCheckboxMultiSelect(storeFilterContainer, storeOptions.map(s => s.text), false);
        filterStoreOptions();

        if (storeFilterContainer) {
            Array.from(storeFilterContainer.querySelectorAll('input[type="checkbox"]')).forEach(checkbox => {
                checkbox.checked = previouslyCheckedStores.has(checkbox.value) && checkbox.closest('.cb-item').style.display !== 'none';
            });
        }
    };

    const filterStoreOptions = () => {
        if (!storeFilterContainer) return;
        const searchTerm = storeSearch.value.toLowerCase();
        let visibleCount = 0;

        Array.from(storeFilterContainer.querySelectorAll('.cb-item')).forEach(item => {
            const labelText = item.textContent.toLowerCase();
            const matches = labelText.includes(searchTerm);
            item.style.display = matches ? 'block' : 'none';
            if (matches) {
                visibleCount++;
            }
        });

        storeSelectAll.disabled = storeFilterContainer.classList.contains('disabled') || visibleCount === 0;
        storeDeselectAll.disabled = storeFilterContainer.classList.contains('disabled') || visibleCount === 0;

         if (storeFilterContainer.classList.contains('disabled')) {
            storeFilterContainer.dataset.placeholder = "-- Load File First --";
         } else if (storeOptions.length === 0 && rawData.length > 0) { // Check if rawData loaded but no options match hierarchy
             storeFilterContainer.dataset.placeholder = "-- No matching options --";
         } else if (visibleCount === 0 && searchTerm !== '') {
            storeFilterContainer.dataset.placeholder = "-- No stores match search --";
         } else if (visibleCount === 0 && storeOptions.length === 0) { // Handles case where storeOptions is initially empty
             storeFilterContainer.dataset.placeholder = "-- No matching options --";
         } else {
            delete storeFilterContainer.dataset.placeholder;
         }
    };


    // ** CORRECTED applyFilters with String Conversion **
    const applyFilters = () => {
        if (rawData.length === 0) { statusDiv.textContent = "Please load a file first."; return; }
        showLoading(true, true);
        resultsArea.style.display = 'none';

        setTimeout(() => {
            try {
                // Retrieve selected values (already strings from checkboxes/selects)
                const selectedRegion = regionFilter.value;
                const selectedDistrict = districtFilter.value;
                const selectedTerritories = territoryFilterContainer
                    ? Array.from(territoryFilterContainer.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value)
                    : [];
                const selectedFsm = fsmFilter.value;
                const selectedChannel = channelFilter.value;
                const selectedSubchannel = subchannelFilter.value;
                const selectedDealer = dealerFilter.value;
                const selectedStores = storeFilterContainer
                    ? Array.from(storeFilterContainer.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value)
                    : [];
                const selectedFlags = {};
                 Object.entries(flagFiltersCheckboxes).forEach(([key, input]) => {
                     if (input && input.checked) { selectedFlags[key] = true; }
                 });

                // Filter rawData
                filteredData = rawData.filter(row => {
                    // --- String Conversion Applied Here ---
                    if (selectedRegion !== 'ALL' && String(safeGet(row, 'REGION', '')) !== selectedRegion) return false;
                    if (selectedDistrict !== 'ALL' && String(safeGet(row, 'DISTRICT', '')) !== selectedDistrict) return false;
                    // Check includes using string values from row data
                    if (selectedTerritories.length > 0 && !selectedTerritories.includes(String(safeGet(row, 'Q2 Territory', '')))) return false;
                    if (selectedFsm !== 'ALL' && String(safeGet(row, 'FSM NAME', '')) !== selectedFsm) return false;
                    if (selectedChannel !== 'ALL' && String(safeGet(row, 'CHANNEL', '')) !== selectedChannel) return false;
                    if (selectedSubchannel !== 'ALL' && String(safeGet(row, 'SUB_CHANNEL', '')) !== selectedSubchannel) return false;
                    if (selectedDealer !== 'ALL' && String(safeGet(row, 'DEALER_NAME', '')) !== selectedDealer) return false;
                    // Check includes using string values from row data
                    if (selectedStores.length > 0 && !selectedStores.includes(String(safeGet(row, 'Store', '')))) return false;
                    // --- End String Conversion ---

                    // Flag logic remains the same
                    for (const flag in selectedFlags) {
                        const flagValue = safeGet(row, flag, 'NO'); // Use null default for safeGet
                        // Check multiple truthy representations
                        if (!(flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1')) {
                           return false;
                        }
                    }
                    return true; // Row passes all filters
                });

                // Update UI sections
                updateSummary(filteredData);
                updateCharts(filteredData);
                updateAttachRateTable(filteredData);

                // Show/hide store details
                if (filteredData.length === 1) {
                    showStoreDetails(filteredData[0]);
                    highlightTableRow(safeGet(filteredData[0], 'Store', null));
                 } else { hideStoreDetails(); }

                // Final UI updates
                statusDiv.textContent = `Displaying ${filteredData.length} of ${rawData.length} rows based on filters.`;
                resultsArea.style.display = 'block';
                exportCsvButton.disabled = filteredData.length === 0;

            } catch (error) {
                 console.error("Error applying filters:", error);
                 statusDiv.textContent = "Error applying filters. Check console for details.";
                 filteredData = []; resultsArea.style.display = 'none'; exportCsvButton.disabled = true;
                 updateSummary([]); updateCharts([]); updateAttachRateTable([]); hideStoreDetails();
             } finally { showLoading(false, true); }
        }, 10);
    };

    const resetFilters = () => {
         [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(sel => { if (sel) { sel.value = 'ALL'; } });
         if (territoryFilterContainer) territoryFilterContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
         if (storeFilterContainer) storeFilterContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
         if (storeSearch) storeSearch.value = '';
         Object.values(flagFiltersCheckboxes).forEach(input => { if(input) input.checked = false; });
         updateStoreFilterOptionsBasedOnHierarchy(); // This re-populates stores and applies search filter
         currentSort = { column: 'Store', ascending: true };
    };

     const handleResetFilters = () => {
         resetFilters();
         applyFilters();
     };

     const resetUI = (apply = true) => {
         resetFilters(); // This now also handles resetting checkbox containers via updateStoreFilterOptions...

         const disable = rawData.length === 0;
         [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter, storeSearch, territorySelectAll, territoryDeselectAll, storeSelectAll, storeDeselectAll, applyFiltersButton, resetFiltersButton, exportCsvButton].forEach(el => {if(el) el.disabled = disable});
         Object.values(flagFiltersCheckboxes).forEach(input => { if(input) input.disabled = disable });
         // Ensure containers are explicitly disabled/placeholder set if needed
         if (territoryFilterContainer) populateCheckboxMultiSelect(territoryFilterContainer, [], disable);
         if (storeFilterContainer) populateCheckboxMultiSelect(storeFilterContainer, [], disable);

         if (disable) { statusDiv.textContent = 'No file selected.'; }

         if (filterArea) filterArea.style.display = disable ? 'none' : 'block';
         if (resultsArea) resultsArea.style.display = 'none';
         if (mainChartInstance) { mainChartInstance.destroy(); mainChartInstance = null; }
         if (attachRateTableBody) attachRateTableBody.innerHTML = '';
         if (attachRateTableFooter) attachRateTableFooter.innerHTML = '';
         if (attachTableStatus) attachTableStatus.textContent = '';
         hideStoreDetails();
         updateSummary([]);

         const handler = updateStoreFilterOptionsBasedOnHierarchy;
          [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter, territoryFilterContainer].forEach(el => {
              if (el) el.removeEventListener('change', handler);
          });
          Object.values(flagFiltersCheckboxes).forEach(input => {
             if (input) input.removeEventListener('change', handler);
          });

         if (apply && !disable) { applyFilters(); }
     };

    const updateSummary = (data) => { /* ... (keep existing updateSummary logic) ... */
        const totalCount = data.length;

        const fieldsToClear = [revenueWithDFValue, qtdRevenueTargetValue, qtdGapValue, quarterlyRevenueTargetValue,
                        percentQuarterlyStoreTargetValue, revARValue, unitsWithDFValue, unitTargetValue,
                        unitAchievementValue, visitCountValue, trainingCountValue, retailModeConnectivityValue,
                        repSkillAchValue, vPmrAchValue, postTrainingScoreValue, eliteValue,
                        percentQuarterlyTerritoryTargetValue, territoryRevPercentValue, districtRevPercentValue, regionRevPercentValue];
        fieldsToClear.forEach(el => { if (el) el.textContent = 'N/A'; });
        [percentQuarterlyTerritoryTargetP, territoryRevPercentP, districtRevPercentP, regionRevPercentP].forEach(p => { if(p) p.style.display = 'none';});

        if (totalCount === 0) { return; } // Exit if filteredData is empty

        // ... (rest of calculations and DOM updates) ...
        const sumRevenue = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Revenue w/DF', 0)), 0);
        // ...
         revenueWithDFValue && (revenueWithDFValue.textContent = formatCurrency(sumRevenue));
         // ... etc ...
        updateContextualSummary(data);

     };

    const updateContextualSummary = (data) => { /* ... (keep existing updateContextualSummary logic) ... */ };

    const updateCharts = (data) => { /* ... (keep existing updateCharts logic, ensuring it uses chartOptionsConfig.colorDefaults) ... */
        if (mainChartInstance) { mainChartInstance.destroy(); mainChartInstance = null; }
        if (data.length === 0 || !mainChartCanvas) return;

        // ... sort data, prepare labels/datasets ...
        const sortedData = [...data].sort((a, b) => parseNumber(safeGet(b, 'Revenue w/DF', 0)) - parseNumber(safeGet(a, 'Revenue w/DF', 0)));
        const chartData = sortedData.slice(0, TOP_N_CHART);
        const labels = chartData.map(row => safeGet(row, 'Store', 'Unknown'));
        const revenueDataSet = chartData.map(row => parseNumber(safeGet(row, 'Revenue w/DF', 0)));
        const targetDataSet = chartData.map(row => parseNumber(safeGet(row, 'QTD Revenue Target', 0)));


        // Check if chartOptionsConfig and colorDefaults exist before destructuring
        const colors = chartOptionsConfig?.colorDefaults || {};
        const { barGoodBg = 'rgba(75, 192, 192, 0.6)', barGoodBorder = 'rgba(75, 192, 192, 1)', barBadBg = 'rgba(255, 99, 132, 0.6)', barBadBorder = 'rgba(255, 99, 132, 1)', lineTargetColor = 'rgba(255, 206, 86, 1)', lineTargetBg = 'rgba(255, 206, 86, 0.2)' } = colors;


        const backgroundColors = chartData.map((row, index) => (revenueDataSet[index] >= targetDataSet[index] ? barGoodBg : barBadBg));
        const borderColors = chartData.map((row, index) => (revenueDataSet[index] >= targetDataSet[index] ? barGoodBorder : barBadBorder));


        mainChartInstance = new Chart(mainChartCanvas, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [
                    { label: 'Total Revenue (incl. DF)', data: revenueDataSet, backgroundColor: backgroundColors, borderColor: borderColors, borderWidth: 1 },
                    { label: 'QTD Revenue Target', data: targetDataSet, type: 'line', borderColor: lineTargetColor, backgroundColor: lineTargetBg, fill: false, tension: 0.1, borderWidth: 2, pointRadius: 3, pointHoverRadius: 5 }
                ]
            },
            options: chartOptionsConfig // Use the main config object
        });
     };

    const updateAttachRateTable = (data) => { /* ... (keep existing updateAttachRateTable logic) ... */ };
    const handleSort = (event) => { /* ... (keep existing handleSort logic) ... */ };
    const updateSortArrows = () => { /* ... (keep existing updateSortArrows logic) ... */ };
    const showStoreDetails = (storeData) => { /* ... (keep existing showStoreDetails logic) ... */ };
    const hideStoreDetails = () => { /* ... (keep existing hideStoreDetails logic) ... */ };
     const highlightTableRow = (storeName) => { /* ... (keep existing highlightTableRow logic) ... */ };
    const exportData = () => { /* ... (keep existing exportData logic) ... */ };
    const getFilterSummary = () => { /* ... (keep existing getFilterSummary logic using checkbox containers) ... */ };
    const generateEmailBody = () => { /* ... (keep existing generateEmailBody logic) ... */ };
    const handleShareEmail = () => { /* ... (keep existing handleShareEmail logic) ... */ };
     const selectAllOptions = (containerElement) => { /* ... (keep existing selectAllOptions for checkboxes) ... */ };
     const deselectAllOptions = (containerElement) => { /* ... (keep existing deselectAllOptions for checkboxes) ... */ };


    // --- Event Listeners ---
    excelFileInput?.addEventListener('change', handleFile);
    applyFiltersButton?.addEventListener('click', applyFilters);
    resetFiltersButton?.addEventListener('click', handleResetFilters);
    themeToggleButton?.addEventListener('click', toggleTheme);
    storeSearch?.addEventListener('input', filterStoreOptions);
    exportCsvButton?.addEventListener('click', exportData);
    shareEmailButton?.addEventListener('click', handleShareEmail);
    closeStoreDetailsButton?.addEventListener('click', hideStoreDetails);
    territorySelectAll?.addEventListener('click', () => selectAllOptions(territoryFilterContainer));
    territoryDeselectAll?.addEventListener('click', () => deselectAllOptions(territoryFilterContainer));
    storeSelectAll?.addEventListener('click', () => selectAllOptions(storeFilterContainer));
    storeDeselectAll?.addEventListener('click', () => deselectAllOptions(storeFilterContainer));
    attachRateTable?.querySelector('thead')?.addEventListener('click', handleSort);

    // --- Initial Setup ---
    const savedTheme = localStorage.getItem('dashboardTheme') || 'dark';
    applyTheme(savedTheme); // Apply theme first
    resetUI(false); // Reset UI without applying filters initially

}); // End DOMContentLoaded

