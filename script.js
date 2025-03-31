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
    // Initialize chart options based on the initially applied theme
    let chartOptionsConfig; // Will be set after applying initial theme

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
         if (typeof value === 'number') return value; // Assume it's already decimal
         if (typeof value === 'string') {
            const num = parseFloat(value.replace('%', ''));
            return isNaN(num) ? NaN : num / 100; // Convert to decimal, NaN if fails
         }
         return NaN; // Default to NaN
    };
    const safeGet = (obj, path, defaultValue = 'N/A') => {
        const value = obj ? obj[path] : undefined;
        return (value !== undefined && value !== null) ? value : defaultValue;
    };
    const isValidForAverage = (value) => {
         if (value === null || value === undefined || value === '') return false;
         // Check if it parses to a number (handles both numbers and numeric strings)
         return !isNaN(parseNumber(String(value).replace('%',''))); // Check parseNumber after removing % just in case
    };
    const getUniqueValues = (data, column) => {
        // Use safeGet with '' as default to handle potential missing values gracefully
        const values = new Set(data.map(item => safeGet(item, column, '')).filter(val => val !== ''));
        return [...Array.from(values).sort()]; // Return sorted unique values only
    };
    const setOptions = (selectElement, options, disable = false) => { // For regular <select>
        selectElement.innerHTML = '';
        // Add the 'ALL' option manually for single selects
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
         selectElement.value = 'ALL'; // Default to ALL
    };

    const populateCheckboxMultiSelect = (containerElement, options, disable = false) => {
        if (!containerElement) return;
        containerElement.innerHTML = ''; // Clear existing checkboxes
        containerElement.scrollTop = 0; // Reset scroll position

        if (options.length === 0 && !disable) {
             containerElement.dataset.placeholder = "-- No matching options --";
        } else if (disable) {
            containerElement.dataset.placeholder = "-- Load File First --";
        } else {
             delete containerElement.dataset.placeholder; // Remove placeholder when populated
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
            themeToggleButton.textContent = '🌙'; // Moon icon for dark mode toggle
            themeToggleButton.title = 'Switch to Dark Mode';
            if (themeMetaTag) themeMetaTag.content = '#f4f4f8'; // Light background color from CSS
        } else {
            document.body.classList.remove('light-mode');
            themeToggleButton.textContent = '☀️'; // Sun icon for light mode toggle
            themeToggleButton.title = 'Switch to Light Mode';
            if (themeMetaTag) themeMetaTag.content = '#1e1e1e'; // Dark background color
        }
        // Update chart options based on theme
        chartOptionsConfig = getChartOptions(theme); // Update the global config
        // Force chart redraw if it exists to apply colors
        if (mainChartInstance) {
            // Merge new options, preserving data if possible, or just update fully
            // Note: Fully replacing options might be safer for complex changes
            mainChartInstance.options = chartOptionsConfig;
            mainChartInstance.update(); // Redraw the chart
        }
        // Add similar logic for secondaryChartInstance if/when implemented
    };

    const toggleTheme = () => {
        const currentTheme = document.body.classList.contains('light-mode') ? 'light' : 'dark';
        const newTheme = currentTheme === 'light' ? 'dark' : 'light';
        localStorage.setItem('dashboardTheme', newTheme);
        applyTheme(newTheme);
    };

    // ** CORRECTED getChartOptions **
    function getChartOptions(theme = 'dark') {
        const isLightMode = theme === 'light';
        // Define colors based on theme
        const gridColor = isLightMode ? 'rgba(0, 0, 0, 0.1)' : 'rgba(224, 224, 224, 0.2)';
        const tickColor = isLightMode ? '#6c757d' : '#e0e0e0';
        const labelColor = isLightMode ? '#212529' : '#e0e0e0';
        // Define other colors needed for datasets
        const barGoodBg = isLightMode ? 'rgba(25, 135, 84, 0.6)' : 'rgba(75, 192, 192, 0.6)';
        const barGoodBorder = isLightMode ? 'rgba(25, 135, 84, 1)' : 'rgba(75, 192, 192, 1)';
        const barBadBg = isLightMode ? 'rgba(220, 53, 69, 0.6)' : 'rgba(255, 99, 132, 0.6)';
        const barBadBorder = isLightMode ? 'rgba(220, 53, 69, 1)' : 'rgba(255, 99, 132, 1)';
        const lineTargetColor = isLightMode ? 'rgba(255, 193, 7, 1)' : 'rgba(255, 206, 86, 1)';
        const lineTargetBg = isLightMode ? 'rgba(255, 193, 7, 0.2)' : 'rgba(255, 206, 86, 0.2)';

        // Return the options object
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
             // Store theme-dependent colors directly within the returned object
             // so updateCharts can access them easily.
             colorDefaults: {
                barGoodBg, barGoodBorder, barBadBg, barBadBorder, lineTargetColor, lineTargetBg
             }
             // **REMOVED incorrect line from previous attempt:**
             // barGoodBg, barGoodBorder, barBadBg, barBadBorder, lineTargetColor, lineTargetBg // <- This was the error
        };
    }


    // --- Core Functions ---

    const handleFile = async (event) => {
        const file = event.target.files[0];
        if (!file) {
            statusDiv.textContent = 'No file selected.';
            return;
        }

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
            } else {
                throw new Error("Excel sheet appears to be empty.");
            }

            rawData = jsonData;
            allPossibleStores = [...new Set(rawData.map(r => safeGet(r, 'Store', null)).filter(s => s))]
                                 .sort()
                                 .map(s => ({ value: s, text: s }));

            statusDiv.textContent = `Loaded ${rawData.length} rows. Applying default filters...`;
            populateFilters(rawData);
            filterArea.style.display = 'block';
            resetFiltersButton.disabled = false;

            applyFilters(); // Auto-apply filters

        } catch (error) {
            console.error('Error processing file:', error);
            statusDiv.textContent = `Error: ${error.message}`;
            rawData = [];
            allPossibleStores = [];
            filteredData = [];
            resetUI(true); // Reset fully on error
        } finally {
            showLoading(false);
            excelFileInput.value = '';
        }
    };

    const populateFilters = (data) => {
        // Populate single-select dropdowns
        setOptions(regionFilter, getUniqueValues(data, 'REGION'));
        setOptions(districtFilter, getUniqueValues(data, 'DISTRICT'));
        setOptions(fsmFilter, getUniqueValues(data, 'FSM NAME'));
        setOptions(channelFilter, getUniqueValues(data, 'CHANNEL'));
        setOptions(subchannelFilter, getUniqueValues(data, 'SUB_CHANNEL'));
        setOptions(dealerFilter, getUniqueValues(data, 'DEALER_NAME'));

        // Populate multi-select checkbox containers
        const territoryOptions = getUniqueValues(data, 'Q2 Territory');
        populateCheckboxMultiSelect(territoryFilterContainer, territoryOptions, false);

        storeOptions = [...allPossibleStores];
        populateCheckboxMultiSelect(storeFilterContainer, storeOptions.map(s => s.text), false);

        // Enable flag filters and search
        Object.values(flagFiltersCheckboxes).forEach(input => { if(input) input.disabled = false });
        storeSearch.disabled = false;

        // Enable Apply/Reset buttons
        applyFiltersButton.disabled = false;
        resetFiltersButton.disabled = false;

        addDependencyFilterListeners();
    };

    const addDependencyFilterListeners = () => {
        const handler = updateStoreFilterOptionsBasedOnHierarchy;
        // Standard filters
        [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(filter => {
            if (filter) {
                filter.removeEventListener('change', handler);
                filter.addEventListener('change', handler);
            }
        });
        // Flag checkboxes
        Object.values(flagFiltersCheckboxes).forEach(input => {
             if (input) {
                 input.removeEventListener('change', handler);
                 input.addEventListener('change', handler);
             }
        });
        // Territory checkbox container
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
            if (selectedRegion !== 'ALL' && safeGet(row, 'REGION', null) !== selectedRegion) return false;
            if (selectedDistrict !== 'ALL' && safeGet(row, 'DISTRICT', null) !== selectedDistrict) return false;
            if (selectedTerritories.length > 0 && !selectedTerritories.includes(safeGet(row, 'Q2 Territory', null))) return false;
            if (selectedFsm !== 'ALL' && safeGet(row, 'FSM NAME', null) !== selectedFsm) return false;
            if (selectedChannel !== 'ALL' && safeGet(row, 'CHANNEL', null) !== selectedChannel) return false;
            if (selectedSubchannel !== 'ALL' && safeGet(row, 'SUB_CHANNEL', null) !== selectedSubchannel) return false;
            if (selectedDealer !== 'ALL' && safeGet(row, 'DEALER_NAME', null) !== selectedDealer) return false;

            for (const flag in selectedFlags) {
                 const flagValue = safeGet(row, flag, 'NO');
                 if (!(flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1')) {
                    return false;
                 }
            }
            return true;
        });

        const validStoreNames = new Set(potentiallyValidStoresData.map(row => safeGet(row, 'Store', null)).filter(Boolean));
        storeOptions = Array.from(validStoreNames).sort().map(s => ({ value: s, text: s }));

        const previouslyCheckedStores = storeFilterContainer
            ? new Set(Array.from(storeFilterContainer.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value))
            : new Set();

        populateCheckboxMultiSelect(storeFilterContainer, storeOptions.map(s => s.text), false);

        filterStoreOptions(); // Apply search term filter to the newly populated checkboxes

        // Re-check previously selected items that are still visible
        if (storeFilterContainer) {
            Array.from(storeFilterContainer.querySelectorAll('input[type="checkbox"]')).forEach(checkbox => {
                if (previouslyCheckedStores.has(checkbox.value) && checkbox.closest('.cb-item').style.display !== 'none') {
                    checkbox.checked = true;
                } else {
                    // Ensure checkboxes that shouldn't be checked (either not previously checked or now hidden by search) are unchecked
                     checkbox.checked = false;
                }
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

        // Update placeholder text based on visibility and state
         if (storeFilterContainer.classList.contains('disabled')) {
            storeFilterContainer.dataset.placeholder = "-- Load File First --";
         } else if (storeOptions.length === 0) {
             storeFilterContainer.dataset.placeholder = "-- No matching options --";
         } else if (visibleCount === 0 && searchTerm !== '') {
            storeFilterContainer.dataset.placeholder = "-- No stores match search --";
         } else {
            delete storeFilterContainer.dataset.placeholder;
         }
    };


    const applyFilters = () => {
        if (rawData.length === 0) { statusDiv.textContent = "Please load a file first."; return; }
        showLoading(true, true);
        resultsArea.style.display = 'none';

        setTimeout(() => {
            try {
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

                filteredData = rawData.filter(row => {
                    if (selectedRegion !== 'ALL' && safeGet(row, 'REGION', null) !== selectedRegion) return false;
                    if (selectedDistrict !== 'ALL' && safeGet(row, 'DISTRICT', null) !== selectedDistrict) return false;
                    if (selectedTerritories.length > 0 && !selectedTerritories.includes(safeGet(row, 'Q2 Territory', null))) return false;
                    if (selectedFsm !== 'ALL' && safeGet(row, 'FSM NAME', null) !== selectedFsm) return false;
                    if (selectedChannel !== 'ALL' && safeGet(row, 'CHANNEL', null) !== selectedChannel) return false;
                    if (selectedSubchannel !== 'ALL' && safeGet(row, 'SUB_CHANNEL', null) !== selectedSubchannel) return false;
                    if (selectedDealer !== 'ALL' && safeGet(row, 'DEALER_NAME', null) !== selectedDealer) return false;
                    if (selectedStores.length > 0 && !selectedStores.includes(safeGet(row, 'Store', null))) return false;

                    for (const flag in selectedFlags) {
                        const flagValue = safeGet(row, flag, 'NO');
                         if (!(flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1')) {
                            return false;
                         }
                    }
                    return true;
                });

                updateSummary(filteredData);
                updateCharts(filteredData);
                updateAttachRateTable(filteredData);

                if (filteredData.length === 1) {
                    showStoreDetails(filteredData[0]);
                    highlightTableRow(safeGet(filteredData[0], 'Store', null));
                 } else { hideStoreDetails(); }

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

         if (territoryFilterContainer) {
            territoryFilterContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
         }
         if (storeFilterContainer) {
            storeFilterContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
         }

         if (storeSearch) { storeSearch.value = ''; }
         Object.values(flagFiltersCheckboxes).forEach(input => { if(input) {input.checked = false;} });
         updateStoreFilterOptionsBasedOnHierarchy();
         currentSort = { column: 'Store', ascending: true };
    };

     const handleResetFilters = () => {
         resetFilters();
         applyFilters();
     };

     const resetUI = (apply = true) => {
         resetFilters();

         const disable = rawData.length === 0;
         [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter, storeSearch, territorySelectAll, territoryDeselectAll, storeSelectAll, storeDeselectAll, applyFiltersButton, resetFiltersButton, exportCsvButton].forEach(el => {if(el) el.disabled = disable});
         Object.values(flagFiltersCheckboxes).forEach(input => { if(input) input.disabled = disable });
         // Use populate function to handle disable state and placeholder for checkbox lists
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

    const updateSummary = (data) => {
        const totalCount = data.length;

        const fieldsToClear = [revenueWithDFValue, qtdRevenueTargetValue, qtdGapValue, quarterlyRevenueTargetValue,
                        percentQuarterlyStoreTargetValue, revARValue, unitsWithDFValue, unitTargetValue,
                        unitAchievementValue, visitCountValue, trainingCountValue, retailModeConnectivityValue,
                        repSkillAchValue, vPmrAchValue, postTrainingScoreValue, eliteValue,
                        percentQuarterlyTerritoryTargetValue, territoryRevPercentValue, districtRevPercentValue, regionRevPercentValue];
        fieldsToClear.forEach(el => { if (el) el.textContent = 'N/A'; });
        [percentQuarterlyTerritoryTargetP, territoryRevPercentP, districtRevPercentP, regionRevPercentP].forEach(p => { if(p) p.style.display = 'none';});

        if (totalCount === 0) { return; }

        // Calculate SUMS
        const sumRevenue = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Revenue w/DF', 0)), 0);
        const sumQtdTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'QTD Revenue Target', 0)), 0);
        const sumQuarterlyTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Quarterly Revenue Target', 0)), 0);
        const sumUnits = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Unit w/ DF', 0)), 0);
        const sumUnitTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Unit Target', 0)), 0);
        const sumVisits = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Visit count', 0)), 0);
        const sumTrainings = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Trainings', 0)), 0);

        // Calculate AVERAGES (excluding blanks/invalid values)
        let sumRevAR = 0, countRevAR = 0;
        let sumConnectivity = 0, countConnectivity = 0;
        let sumRepSkill = 0, countRepSkill = 0;
        let sumPmr = 0, countPmr = 0;
        let sumPostTraining = 0, countPostTraining = 0;
        let sumElite = 0, countElite = 0;

        data.forEach(row => {
            let val;
            val = safeGet(row, 'Rev AR%', null); if (isValidForAverage(val)) { sumRevAR += parsePercent(val); countRevAR++; }
            val = safeGet(row, 'Retail Mode Connectivity', null); if (isValidForAverage(val)) { sumConnectivity += parsePercent(val); countConnectivity++; }
            val = safeGet(row, 'Rep Skill Ach', null); if (isValidForAverage(val)) { sumRepSkill += parsePercent(val); countRepSkill++; }
            val = safeGet(row, '(V)PMR Ach', null); if (isValidForAverage(val)) { sumPmr += parsePercent(val); countPmr++; }
            val = safeGet(row, 'Post Training Score', null); if (isValidForAverage(val)) { sumPostTraining += parseNumber(val); countPostTraining++; }
            val = safeGet(row, 'Elite', null); if (isValidForAverage(val)) { sumElite += parsePercent(val); countElite++; }
        });

        const avgRevAR = countRevAR === 0 ? NaN : sumRevAR / countRevAR;
        const avgConnectivity = countConnectivity === 0 ? NaN : sumConnectivity / countConnectivity;
        const avgRepSkill = countRepSkill === 0 ? NaN : sumRepSkill / countRepSkill;
        const avgPmr = countPmr === 0 ? NaN : sumPmr / countPmr;
        const avgPostTraining = countPostTraining === 0 ? NaN : sumPostTraining / countPostTraining;
        const avgElite = countElite === 0 ? NaN : sumElite / countElite;

        const overallPercentStoreTarget = sumQuarterlyTarget === 0 ? 0 : sumRevenue / sumQuarterlyTarget;
        const overallUnitAchievement = sumUnitTarget === 0 ? 0 : sumUnits / sumUnitTarget;

        // Update DOM Elements (Simplified view, original logic remains)
        revenueWithDFValue && (revenueWithDFValue.textContent = formatCurrency(sumRevenue));
        // ... (rest of DOM updates for sums and averages) ...
        eliteValue && (eliteValue.textContent = formatPercent(avgElite));


        updateContextualSummary(data);
    };

    const updateContextualSummary = (data) => {
        [percentQuarterlyTerritoryTargetP, territoryRevPercentP, districtRevPercentP, regionRevPercentP].forEach(p => {if (p) p.style.display = 'none'});
        if (data.length === 0) return;

        const singleRegion = regionFilter.value !== 'ALL';
        const singleDistrict = districtFilter.value !== 'ALL';
        // Update check for single territory based on checkbox container
        const selectedTerritoriesCount = territoryFilterContainer
             ? territoryFilterContainer.querySelectorAll('input[type="checkbox"]:checked').length
             : 0;
        const singleTerritory = selectedTerritoriesCount === 1;


        const calculateAverageExcludeBlanks = (column) => {
            let sum = 0, count = 0;
            data.forEach(row => {
                const val = safeGet(row, column, null);
                if (isValidForAverage(val)) { sum += parsePercent(val); count++; }
            });
            return count === 0 ? NaN : sum / count;
        };

        const avgPercentTerritoryTarget = calculateAverageExcludeBlanks('%Quarterly Territory Rev Target');
        const avgTerritoryRevPercent = calculateAverageExcludeBlanks('Territory Rev%');
        const avgDistrictRevPercent = calculateAverageExcludeBlanks('District Rev%');
        const avgRegionRevPercent = calculateAverageExcludeBlanks('Region Rev%');

        if (percentQuarterlyTerritoryTargetValue) percentQuarterlyTerritoryTargetValue.textContent = formatPercent(avgPercentTerritoryTarget);
        if (percentQuarterlyTerritoryTargetP) percentQuarterlyTerritoryTargetP.style.display = 'block';

        if (singleTerritory || singleDistrict || singleRegion) {
             if (territoryRevPercentValue) territoryRevPercentValue.textContent = formatPercent(avgTerritoryRevPercent);
             if (territoryRevPercentP) territoryRevPercentP.style.display = 'block';
        }
        if (singleDistrict || singleRegion) {
             if (districtRevPercentValue) districtRevPercentValue.textContent = formatPercent(avgDistrictRevPercent);
             if (districtRevPercentP) districtRevPercentP.style.display = 'block';
        }
         if (singleRegion) {
             if (regionRevPercentValue) regionRevPercentValue.textContent = formatPercent(avgRegionRevPercent);
             if (regionRevPercentP) regionRevPercentP.style.display = 'block';
         }
    };

    const updateCharts = (data) => {
        if (mainChartInstance) {
            mainChartInstance.destroy();
            mainChartInstance = null;
        }

        if (data.length === 0 || !mainChartCanvas) return;

        const sortedData = [...data].sort((a, b) => parseNumber(safeGet(b, 'Revenue w/DF', 0)) - parseNumber(safeGet(a, 'Revenue w/DF', 0)));
        const chartData = sortedData.slice(0, TOP_N_CHART);

        const labels = chartData.map(row => safeGet(row, 'Store', 'Unknown'));
        const revenueDataSet = chartData.map(row => parseNumber(safeGet(row, 'Revenue w/DF', 0)));
        const targetDataSet = chartData.map(row => parseNumber(safeGet(row, 'QTD Revenue Target', 0)));

        // Access theme-specific colors from the global options config
        const { barGoodBg, barGoodBorder, barBadBg, barBadBorder, lineTargetColor, lineTargetBg } = chartOptionsConfig.colorDefaults;

        const backgroundColors = chartData.map((row, index) => {
             const revenue = revenueDataSet[index];
             const target = targetDataSet[index];
             return revenue >= target ? barGoodBg : barBadBg;
        });
        const borderColors = chartData.map((row, index) => {
             const revenue = revenueDataSet[index];
             const target = targetDataSet[index];
             return revenue >= target ? barGoodBorder : barBadBorder;
        });

        mainChartInstance = new Chart(mainChartCanvas, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [
                    {
                        label: 'Total Revenue (incl. DF)',
                        data: revenueDataSet,
                        backgroundColor: backgroundColors,
                        borderColor: borderColors,
                        borderWidth: 1
                    },
                    {
                        label: 'QTD Revenue Target',
                        data: targetDataSet,
                        type: 'line',
                        borderColor: lineTargetColor, // Use theme color
                        backgroundColor: lineTargetBg, // Use theme color
                        fill: false,
                        tension: 0.1,
                        borderWidth: 2,
                        pointRadius: 3,
                        pointHoverRadius: 5
                    }
                ]
            },
            options: chartOptionsConfig // Use the centrally managed options object
        });
    };

    const updateAttachRateTable = (data) => {
        if (!attachRateTableBody || !attachRateTableFooter) return;
        attachRateTableBody.innerHTML = '';
        attachRateTableFooter.innerHTML = '';

        if (data.length === 0) {
            if(attachTableStatus) attachTableStatus.textContent = 'No data to display based on filters.';
            return;
        }

        const sortedData = [...data].sort((a, b) => {
             let valA = safeGet(a, currentSort.column, null);
             let valB = safeGet(b, currentSort.column, null);
             if (valA === null && valB === null) return 0;
             if (valA === null) return currentSort.ascending ? -1 : 1;
             if (valB === null) return currentSort.ascending ? 1 : -1;
             const isPercentCol = currentSort.column.includes('Attach Rate') || currentSort.column.includes('%') || currentSort.column.includes('Target');
             const numA = isPercentCol ? parsePercent(valA) : parseNumber(valA);
             const numB = isPercentCol ? parsePercent(valB) : parseNumber(valB);
             if (!isNaN(numA) && !isNaN(numB)) { return currentSort.ascending ? numA - numB : numB - numA; }
             else { valA = String(valA).toLowerCase(); valB = String(valB).toLowerCase(); return currentSort.ascending ? valA.localeCompare(valB) : valB.localeCompare(valA); }
         });

        const averageMetrics = ['% Quarterly Revenue Target', 'Tablet Attach Rate', 'PC Attach Rate', 'NC Attach Rate', 'TWS Attach Rate', 'WW Attach Rate', 'ME Attach Rate', 'NCME Attach Rate'];
        const averages = {};
        averageMetrics.forEach(key => {
             let sum = 0, count = 0;
             data.forEach(row => {
                 const val = safeGet(row, key, null);
                 if (isValidForAverage(val)) { sum += parsePercent(val); count++; }
             });
             averages[key] = count === 0 ? NaN : sum / count;
         });

        sortedData.forEach(row => {
            const tr = document.createElement('tr');
            const storeName = safeGet(row, 'Store', null);
            if (storeName) {
                 tr.dataset.storeName = storeName;
                 tr.onclick = () => { showStoreDetails(row); highlightTableRow(storeName); };

                 const columns = [ /* ... column definitions ... */
                     { key: 'Store', format: (val) => val },
                     { key: '% Quarterly Revenue Target', format: formatPercent, highlight: true },
                     { key: 'Tablet Attach Rate', format: formatPercent, highlight: true },
                     { key: 'PC Attach Rate', format: formatPercent, highlight: true },
                     { key: 'NC Attach Rate', format: formatPercent, highlight: true },
                     { key: 'TWS Attach Rate', format: formatPercent, highlight: true },
                     { key: 'WW Attach Rate', format: formatPercent, highlight: true },
                     { key: 'ME Attach Rate', format: formatPercent, highlight: true },
                     { key: 'NCME Attach Rate', format: formatPercent, highlight: true },
                 ];

                 columns.forEach(col => {
                     const td = document.createElement('td');
                     const rawValue = safeGet(row, col.key, null);
                     const isPercentCol = col.key.includes('Attach Rate') || col.key.includes('%');
                     const numericValue = isPercentCol ? parsePercent(rawValue) : parseNumber(rawValue);
                     let formattedValue = (rawValue === null || (col.key !== 'Store' && isNaN(numericValue)))
                                         ? 'N/A'
                                         : (isPercentCol ? col.format(numericValue) : rawValue);
                     td.textContent = formattedValue;
                     td.title = `${col.key}: ${formattedValue}`;
                     if (col.highlight && averages[col.key] !== undefined && !isNaN(numericValue)) {
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
             averageMetrics.forEach(key => {
                 const td = footerRow.insertCell();
                 const avgValue = averages[key];
                 td.textContent = formatPercent(avgValue);
                 let validCount = data.filter(row => isValidForAverage(safeGet(row, key, null))).length;
                 td.title = `Average ${key}: ${formatPercent(avgValue)} (from ${validCount} stores)`;
                 td.style.textAlign = "right";
             });
         }

        if(attachTableStatus) attachTableStatus.textContent = `Showing ${sortedData.length} stores. Click row for details. Click headers to sort.`;
        updateSortArrows();
    };

    const handleSort = (event) => {
         const headerCell = event.target.closest('th');
         if (!headerCell || !headerCell.classList.contains('sortable')) return;
         const sortKey = headerCell.dataset.sort;
         if (!sortKey) return;
         if (currentSort.column === sortKey) { currentSort.ascending = !currentSort.ascending; }
         else { currentSort.column = sortKey; currentSort.ascending = true; }
         updateAttachRateTable(filteredData);
    };


    const updateSortArrows = () => {
        if (!attachRateTable) return;
        attachRateTable.querySelectorAll('th.sortable .sort-arrow').forEach(arrow => {
            arrow.className = 'sort-arrow';
            arrow.textContent = '';
        });
        const currentHeader = attachRateTable.querySelector(`th[data-sort="${currentSort.column}"] .sort-arrow`);
        if (currentHeader) {
            currentHeader.classList.add(currentSort.ascending ? 'asc' : 'desc');
            currentHeader.textContent = currentSort.ascending ? ' ▲' : ' ▼';
        }
    };

    const showStoreDetails = (storeData) => {
        if (!storeDetailsContent || !storeDetailsSection || !closeStoreDetailsButton) return;

        const addressParts = [safeGet(storeData, 'ADDRESS1', null), safeGet(storeData, 'CITY', null), safeGet(storeData, 'STATE', null), safeGet(storeData, 'ZIPCODE', null)].filter(Boolean);
        const addressString = addressParts.length > 0 ? addressParts.join(', ') : null;

        const latitude = parseFloat(safeGet(storeData, 'LATITUDE_ORG', NaN));
        const longitude = parseFloat(safeGet(storeData, 'LONGITUDE_ORG', NaN));
        let mapsLinkHtml = '';
        if (!isNaN(latitude) && !isNaN(longitude)) {
            const mapsUrl = `https://www.google.com/maps?q=${latitude},${longitude}`;
            mapsLinkHtml = `<p><a href="${mapsUrl}" target="_blank" title="Open in Google Maps">View on Google Maps</a></p>`;
        } else {
            mapsLinkHtml = `<p style="color: var(--text-muted-color); font-style: italic;">(Map coordinates not available)</p>`;
        }

        let flagSummaryHtml = FLAG_HEADERS.map(flag => {
            const flagValue = safeGet(storeData, flag);
            const isTrue = (flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1');
            return `<span title="${flag}" data-flag="${isTrue}">${flag.replace(/_/g, ' ')} ${isTrue ? '✔' : '✘'}</span>`;
        }).join(' | ');

        let detailsHtml = `<p><strong>Store:</strong> ${safeGet(storeData, 'Store')}</p>`;
        if (addressString) { detailsHtml += `<p><strong>Address:</strong> ${addressString}</p>`; }
        detailsHtml += mapsLinkHtml;
        detailsHtml += `<hr>`;
        detailsHtml += `<p><strong>IDs:</strong> Store: ${safeGet(storeData, 'STORE ID')} | Org: ${safeGet(storeData, 'ORG_STORE_ID')} | CV: ${safeGet(storeData, 'CV_STORE_ID')} | CinglePoint: ${safeGet(storeData, 'CINGLEPOINT_ID')}</p>`;
        detailsHtml += `<p><strong>Type:</strong> ${safeGet(storeData, 'STORE_TYPE_NAME')} | Nat Tier: ${safeGet(storeData, 'National_Tier')} | Merch Lvl: ${safeGet(storeData, 'Merchandising_Level')} | Comb Tier: ${safeGet(storeData, 'Combined_Tier')}</p>`;
        detailsHtml += `<hr>`;
        detailsHtml += `<p><strong>Hierarchy:</strong> ${safeGet(storeData, 'REGION')} > ${safeGet(storeData, 'DISTRICT')} > ${safeGet(storeData, 'Q2 Territory')}</p>`;
        detailsHtml += `<p><strong>FSM:</strong> ${safeGet(storeData, 'FSM NAME')}</p>`;
        detailsHtml += `<p><strong>Channel:</strong> ${safeGet(storeData, 'CHANNEL')} / ${safeGet(storeData, 'SUB_CHANNEL')}</p>`;
        detailsHtml += `<p><strong>Dealer:</strong> ${safeGet(storeData, 'DEALER_NAME')}</p>`;
        detailsHtml += `<hr>`;
        detailsHtml += `<p><strong>Visits:</strong> ${formatNumber(parseNumber(safeGet(storeData, 'Visit count', 0)))} | <strong>Trainings:</strong> ${formatNumber(parseNumber(safeGet(storeData, 'Trainings', 0)))}</p>`;
        detailsHtml += `<p><strong>Connectivity:</strong> ${formatPercent(parsePercent(safeGet(storeData, 'Retail Mode Connectivity', 0)))}</p>`;
        detailsHtml += `<hr>`;
        detailsHtml += `<p><strong>Flags:</strong> ${flagSummaryHtml}</p>`;

        storeDetailsContent.innerHTML = detailsHtml;
        storeDetailsSection.style.display = 'block';
        closeStoreDetailsButton.style.display = 'inline-block';
    };

    const hideStoreDetails = () => {
        if (!storeDetailsContent || !storeDetailsSection || !closeStoreDetailsButton) return;
        storeDetailsContent.innerHTML = 'Select a store from the table or chart for details, or apply filters resulting in a single store.';
        storeDetailsSection.style.display = 'none';
        closeStoreDetailsButton.style.display = 'none';
        highlightTableRow(null);
    };

     const highlightTableRow = (storeName) => {
         if (selectedStoreRow) { selectedStoreRow.classList.remove('selected-row'); }
        if (storeName && attachRateTableBody) {
             try {
                 selectedStoreRow = attachRateTableBody.querySelector(`tr[data-store-name="${CSS.escape(storeName)}"]`);
                 if (selectedStoreRow) { selectedStoreRow.classList.add('selected-row'); }
                 else { selectedStoreRow = null; }
             } catch (e) { console.error("Error selecting table row:", e); selectedStoreRow = null; }
         } else { selectedStoreRow = null; }
    };

    const exportData = () => { /* ... (keep existing exportData) ... */ };

    const getFilterSummary = () => {
        let summary = [];
        if (regionFilter?.value !== 'ALL') summary.push(`Region: ${regionFilter.value}`);
        if (districtFilter?.value !== 'ALL') summary.push(`District: ${districtFilter.value}`);
        const territories = territoryFilterContainer
            ? Array.from(territoryFilterContainer.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value)
            : [];
        if (territories.length > 0) summary.push(`Territories: ${territories.length === 1 ? territories[0] : territories.length + ' selected'}`);

        if (fsmFilter?.value !== 'ALL') summary.push(`FSM: ${fsmFilter.value}`);
        if (channelFilter?.value !== 'ALL') summary.push(`Channel: ${channelFilter.value}`);
        if (subchannelFilter?.value !== 'ALL') summary.push(`Subchannel: ${subchannelFilter.value}`);
        if (dealerFilter?.value !== 'ALL') summary.push(`Dealer: ${dealerFilter.value}`);
        const stores = storeFilterContainer
             ? Array.from(storeFilterContainer.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value)
             : [];
         if (stores.length > 0) summary.push(`Stores: ${stores.length === 1 ? stores[0] : stores.length + ' selected'}`);

         const flags = Object.entries(flagFiltersCheckboxes).filter(([key, input]) => input && input.checked).map(([key])=> key.replace(/_/g, ' '));
         if (flags.length > 0) summary.push(`Attributes: ${flags.join(', ')}`);
        return summary.length > 0 ? summary.join('; ') : 'None';
    };


    const generateEmailBody = () => { /* ... (keep existing generateEmailBody) ... */ };
    const handleShareEmail = () => { /* ... (keep existing handleShareEmail) ... */ };

     const selectAllOptions = (containerElement) => {
         if (!containerElement || !containerElement.classList.contains('checkbox-multiselect')) return;
         let changed = false;
         containerElement.querySelectorAll('input[type="checkbox"]').forEach(cb => {
             if (cb.closest('.cb-item').style.display !== 'none') { // Only check visible
                 if (!cb.checked) {
                     cb.checked = true;
                     changed = true;
                 }
             }
         });
          // Trigger change *once* after loop if any checkbox was changed
          if (changed) {
             containerElement.dispatchEvent(new Event('change', { bubbles: true }));
          }
    };

     const deselectAllOptions = (containerElement) => {
         if (!containerElement || !containerElement.classList.contains('checkbox-multiselect')) return;
          let changed = false;
         containerElement.querySelectorAll('input[type="checkbox"]').forEach(cb => {
              if (cb.closest('.cb-item').style.display !== 'none') { // Only uncheck visible
                 if (cb.checked) {
                     cb.checked = false;
                     changed = true;
                 }
             }
         });
         // Trigger change *once* after loop if any checkbox was changed
         if (changed) {
            containerElement.dispatchEvent(new Event('change', { bubbles: true }));
         }
    };


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
    // Apply theme *first* so getChartOptions has the correct context
    applyTheme(savedTheme);
    // Reset UI without triggering applyFilters (handleFile does this on load)
    resetUI(false);

}); // End DOMContentLoaded
