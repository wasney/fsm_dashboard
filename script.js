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
    // Chart colors can be updated dynamically based on theme if needed, setting defaults here
    let chartOptionsConfig = getChartOptions(); // Initialize chart options

    const TOP_N_CHART = 15; // Max items to show on the bar chart

    // --- DOM Elements ---
    const excelFileInput = document.getElementById('excelFile');
    const statusDiv = document.getElementById('status');
    const loadingIndicator = document.getElementById('loadingIndicator');
    const filterLoadingIndicator = document.getElementById('filterLoadingIndicator');
    const filterArea = document.getElementById('filterArea');
    const resultsArea = document.getElementById('resultsArea');
    const applyFiltersButton = document.getElementById('applyFiltersButton');
    const resetFiltersButton = document.getElementById('resetFiltersButton'); // New Reset Button
    const themeToggleButton = document.getElementById('themeToggleButton'); // New Theme Button
    const themeMetaTag = document.querySelector('meta[name="theme-color"]');

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

    // CORRECTED FLAG FILTER CHECKBOX MAPPING
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
    const safeGet = (obj, path, defaultValue = 'N/A') => {
        const value = obj ? obj[path] : undefined;
        return (value !== undefined && value !== null) ? value : defaultValue;
    };
    const isValidForAverage = (value) => {
         if (value === null || value === undefined || value === '') return false;
         return !isNaN(parseNumber(String(value).replace('%','')));
    };
    const getUniqueValues = (data, column) => {
        const values = new Set(data.map(item => safeGet(item, column, '')).filter(val => val !== ''));
        return ['ALL', ...Array.from(values).sort()];
    };
    const setOptions = (selectElement, options, disable = false) => {
        selectElement.innerHTML = '';
        options.forEach(optionValue => {
            const option = document.createElement('option');
            option.value = optionValue;
            option.textContent = optionValue === 'ALL' ? `-- ${optionValue} --` : optionValue;
            option.title = optionValue;
            selectElement.appendChild(option);
        });
         selectElement.disabled = disable;
         selectElement.value = 'ALL'; // Default to ALL
    };
    const setMultiSelectOptions = (selectElement, options, disable = false) => {
         selectElement.innerHTML = '';
         options.forEach(optionValue => {
             if (optionValue === 'ALL') return;
             const option = document.createElement('option');
             option.value = optionValue;
             option.textContent = optionValue;
             option.title = optionValue;
             selectElement.appendChild(option);
         });
         selectElement.disabled = disable;
         selectElement.selectedIndex = -1; // Start deselected
    };
    const showLoading = (isLoading, isFiltering = false) => {
        if (isFiltering) {
            filterLoadingIndicator.style.display = isLoading ? 'flex' : 'none';
            applyFiltersButton.disabled = isLoading || rawData.length === 0; // Also check rawData
            resetFiltersButton.disabled = isLoading || rawData.length === 0; // Also check rawData
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
        chartOptionsConfig = getChartOptions(theme);
        // Force chart redraw if it exists to apply colors
        if (mainChartInstance) {
            mainChartInstance.options = chartOptionsConfig; // Apply new options object
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

    // Function to get chart options based on the current theme
    function getChartOptions(theme = 'dark') {
        const isLightMode = theme === 'light';
        // Define colors based on theme
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
                    backgroundColor: isLightMode ? '#ffffff' : '#2c2c2c', // Adjust tooltip background
                    titleColor: isLightMode ? '#212529' : '#e0e0e0', // Adjust tooltip title color
                    bodyColor: isLightMode ? '#212529' : '#e0e0e0', // Adjust tooltip body color
                    callbacks: {
                         label: function(context) {
                            let label = context.dataset.label || '';
                            if (label) { label += ': '; }
                            if (context.parsed.y !== null) {
                                 if (context.dataset.type === 'line' || context.dataset.label.toLowerCase().includes('target')) {
                                      label += formatCurrency(context.parsed.y);
                                 } else {
                                      label += formatCurrency(context.parsed.y);
                                      // Add % Target achievement to tooltip for the bar if data exists
                                      // Accessing chartData correctly within callback context can be tricky,
                                      // Need to ensure it's accessible or passed correctly if needed.
                                      // For now, relying on parsed value.
                                      // Example: Accessing original data might require looking up via context.dataIndex in the source data array used to create the chart.
                                      const sourceData = mainChartInstance?.config?.data?.datasets[context.datasetIndex]?.data?.[context.dataIndex];
                                      // This is just an example, might need refinement based on how data is structured.

                                      // Let's try getting % from the original filteredData based on label/index
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
             // Keep existing onClick handler
             onClick: (event, elements) => {
                if (elements.length > 0) {
                    const index = elements[0].index;
                    const clickedLabel = mainChartInstance.data.labels[index]; // Get label from chart instance
                    // Find the original data row from filteredData
                    const storeData = filteredData.find(row => safeGet(row, 'Store', null) === clickedLabel);
                    if (storeData) {
                        showStoreDetails(storeData);
                        highlightTableRow(clickedLabel);
                    }
                }
            },
            // Theme specific colors need to be dynamically inserted into dataset properties in updateCharts
             // Store color definitions here for reference in updateCharts
             barGoodBg, barGoodBorder, barBadBg, barBadBorder, lineTargetColor, lineTargetBg
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
        resetUI(false); // Reset UI but don't trigger apply filters yet

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
                    console.warn(`Warning: Missing expected columns: ${missingHeaders.join(', ')}. Some features might not work correctly.`);
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
            resetFiltersButton.disabled = false; // Enable reset button

            // --- ITEM 3: Automatically apply filters after loading ---
            applyFilters();
            // Update status after filters are applied
            // applyFilters will update the status message appropriately

        } catch (error) {
            console.error('Error processing file:', error);
            statusDiv.textContent = `Error: ${error.message}`;
            rawData = [];
            allPossibleStores = [];
            filteredData = [];
            resetUI(true); // Reset UI fully on error
        } finally {
            showLoading(false);
            excelFileInput.value = '';
        }
    };

    const populateFilters = (data) => {
        // Populate top-level filters first
        setOptions(regionFilter, getUniqueValues(data, 'REGION'));
        setOptions(districtFilter, getUniqueValues(data, 'DISTRICT'));
        // Multi-select needs special handling for 'ALL' - remove it from options list
        const territoryOptions = getUniqueValues(data, 'Q2 Territory').filter(v => v !== 'ALL');
        setMultiSelectOptions(territoryFilter, territoryOptions);
        setOptions(fsmFilter, getUniqueValues(data, 'FSM NAME'));
        setOptions(channelFilter, getUniqueValues(data, 'CHANNEL'));
        setOptions(subchannelFilter, getUniqueValues(data, 'SUB_CHANNEL'));
        setOptions(dealerFilter, getUniqueValues(data, 'DEALER_NAME'));

        // Enable flag filters
        Object.values(flagFiltersCheckboxes).forEach(input => { if(input) input.disabled = false });

        // Initial population of store filter (shows all stores initially)
        storeOptions = [...allPossibleStores]; // Start with all stores
        setStoreFilterOptions(storeOptions, false); // Populate and enable

        // Enable multi-select buttons and search
        territorySelectAll.disabled = false;
        territoryDeselectAll.disabled = false;
        storeSelectAll.disabled = false;
        storeDeselectAll.disabled = false;
        storeSearch.disabled = false;

        applyFiltersButton.disabled = false; // Enable apply button
        resetFiltersButton.disabled = false; // Enable reset button

        addDependencyFilterListeners();
    };

    const addDependencyFilterListeners = () => {
        const handler = updateStoreFilterOptionsBasedOnHierarchy;
        [regionFilter, districtFilter, territoryFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(filter => {
            if (filter) {
                filter.removeEventListener('change', handler);
                filter.addEventListener('change', handler);
            }
        });
        Object.values(flagFiltersCheckboxes).forEach(input => {
             if (input) {
                 input.removeEventListener('change', handler);
                 input.addEventListener('change', handler);
             }
        });
    };

     const updateStoreFilterOptionsBasedOnHierarchy = () => {
        if (rawData.length === 0) return;

        const selectedRegion = regionFilter.value;
        const selectedDistrict = districtFilter.value;
        const selectedTerritories = Array.from(territoryFilter.selectedOptions).map(opt => opt.value);
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

        const previouslySelectedStores = new Set(Array.from(storeFilter.selectedOptions).map(opt => opt.value));

        setStoreFilterOptions(storeOptions, false);

        filterStoreOptions(); // Apply search term

        Array.from(storeFilter.options).forEach(option => {
            if (previouslySelectedStores.has(option.value)) {
                option.selected = true;
            }
        });
        if (storeFilter.selectedOptions.length === 0) {
             storeFilter.selectedIndex = -1;
        }
    };

    const setStoreFilterOptions = (optionsToShow, disable = true) => {
        const currentSearchTerm = storeSearch.value;
        storeFilter.innerHTML = '';
        optionsToShow.forEach(opt => {
             const option = document.createElement('option');
             option.value = opt.value;
             option.textContent = opt.text;
             option.title = opt.text;
             storeFilter.appendChild(option);
         });
         storeFilter.disabled = disable;
         storeSearch.disabled = disable;
         storeSelectAll.disabled = disable || optionsToShow.length === 0;
         storeDeselectAll.disabled = disable || optionsToShow.length === 0;
         storeSearch.value = currentSearchTerm;
    };

    const filterStoreOptions = () => {
        const searchTerm = storeSearch.value.toLowerCase();
        const filteredOptions = storeOptions.filter(opt => opt.text.toLowerCase().includes(searchTerm));
        const selectedValues = new Set(Array.from(storeFilter.selectedOptions).map(opt => opt.value));

        storeFilter.innerHTML = '';
        filteredOptions.forEach(opt => {
            const option = document.createElement('option');
            option.value = opt.value;
            option.textContent = opt.text;
            option.title = opt.text;
             if (selectedValues.has(opt.value)) {
                 option.selected = true;
             }
            storeFilter.appendChild(option);
        });

        storeSelectAll.disabled = storeFilter.disabled || filteredOptions.length === 0;
        storeDeselectAll.disabled = storeFilter.disabled || filteredOptions.length === 0;
    };


    const applyFilters = () => {
        if (rawData.length === 0) {
             statusDiv.textContent = "Please load a file first.";
             return; // Don't attempt to filter if no data
        }
        showLoading(true, true);
        resultsArea.style.display = 'none';

        setTimeout(() => {
            try {
                const selectedRegion = regionFilter.value;
                const selectedDistrict = districtFilter.value;
                const selectedTerritories = Array.from(territoryFilter.selectedOptions).map(opt => opt.value);
                const selectedFsm = fsmFilter.value;
                const selectedChannel = channelFilter.value;
                const selectedSubchannel = subchannelFilter.value;
                const selectedDealer = dealerFilter.value;
                const selectedStores = Array.from(storeFilter.selectedOptions).map(opt => opt.value);
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
                } else {
                    hideStoreDetails();
                }

                statusDiv.textContent = `Displaying ${filteredData.length} of ${rawData.length} rows based on filters.`;
                resultsArea.style.display = 'block';
                exportCsvButton.disabled = filteredData.length === 0;

            } catch (error) {
                console.error("Error applying filters:", error);
                statusDiv.textContent = "Error applying filters. Check console for details.";
                filteredData = [];
                resultsArea.style.display = 'none';
                 exportCsvButton.disabled = true;
                 updateSummary([]);
                 updateCharts([]);
                 updateAttachRateTable([]);
                 hideStoreDetails();

            } finally {
                 showLoading(false, true);
            }
        }, 10);
    };

    // --- ITEM 2: Modified resetFilters function ---
    const resetFilters = () => {
        // Reset dropdowns
         [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(sel => { if (sel) { sel.value = 'ALL'; } });
         // Reset multi-selects
         if (territoryFilter) { territoryFilter.selectedIndex = -1; } // Deselect all
         if (storeFilter) { storeFilter.selectedIndex = -1; } // Deselect all

         // Reset search input
         if (storeSearch) { storeSearch.value = ''; }

         // Reset checkboxes
         Object.values(flagFiltersCheckboxes).forEach(input => { if(input) {input.checked = false;} });

         // IMPORTANT: Update the dynamic store filter based on the reset hierarchy filters
         updateStoreFilterOptionsBasedOnHierarchy();
         // After updating store options, clear the search filter visually
         filterStoreOptions();
         // And ensure the store multi-select is visually deselected again
         if (storeFilter) { storeFilter.selectedIndex = -1; }

         // Reset sort state for the table
         currentSort = { column: 'Store', ascending: true };

         // Note: Enabling/disabling is handled by populateFilters initially
         // and showLoading during operations. No need to disable here.
    };

     // --- ITEM 2: Function to handle Reset button click ---
     const handleResetFilters = () => {
         resetFilters(); // Reset all filter inputs to default state
         applyFilters(); // Re-apply filters to show all data
     };

     // Function to reset the entire UI, optionally triggering apply filters
     const resetUI = (apply = true) => {
         resetFilters(); // Reset filter inputs first

         // Disable filters and buttons if no raw data
         const disable = rawData.length === 0;
         [regionFilter, districtFilter, territoryFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter, storeFilter, storeSearch, territorySelectAll, territoryDeselectAll, storeSelectAll, storeDeselectAll, applyFiltersButton, resetFiltersButton, exportCsvButton].forEach(el => {if(el) el.disabled = disable});
         Object.values(flagFiltersCheckboxes).forEach(input => { if(input) input.disabled = disable });

         if (disable) {
            if (territoryFilter) { setMultiSelectOptions(territoryFilter, [], true);} // Clear options and disable
            if (storeFilter) { setStoreFilterOptions([], true); } // Clear options and disable
             statusDiv.textContent = 'No file selected.';
         }

         if (filterArea) filterArea.style.display = disable ? 'none' : 'block';
         if (resultsArea) resultsArea.style.display = 'none';
         if (mainChartInstance) { mainChartInstance.destroy(); mainChartInstance = null; }
         if (attachRateTableBody) attachRateTableBody.innerHTML = '';
         if (attachRateTableFooter) attachRateTableFooter.innerHTML = '';
         if (attachTableStatus) attachTableStatus.textContent = '';
         hideStoreDetails();
         updateSummary([]); // Clear summary fields

         // Remove dependency listeners if they exist to prevent duplicates on reload
         const handler = updateStoreFilterOptionsBasedOnHierarchy;
         [regionFilter, districtFilter, territoryFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(filter => {
             if (filter) filter.removeEventListener('change', handler);
         });
          Object.values(flagFiltersCheckboxes).forEach(input => {
             if (input) input.removeEventListener('change', handler);
          });

         // Apply filters only if requested and data exists
         if (apply && !disable) {
             applyFilters();
         }
     };

    // --- UPDATED updateSummary to exclude blanks from averages ---
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

        // Calculate Overall Percentages
        const overallPercentStoreTarget = sumQuarterlyTarget === 0 ? 0 : sumRevenue / sumQuarterlyTarget;
        const overallUnitAchievement = sumUnitTarget === 0 ? 0 : sumUnits / sumUnitTarget;

        // Update DOM Elements
        revenueWithDFValue && (revenueWithDFValue.textContent = formatCurrency(sumRevenue));
        revenueWithDFValue && (revenueWithDFValue.title = `Sum of 'Revenue w/DF' for ${totalCount} filtered stores`);
        qtdRevenueTargetValue && (qtdRevenueTargetValue.textContent = formatCurrency(sumQtdTarget));
        qtdRevenueTargetValue && (qtdRevenueTargetValue.title = `Sum of 'QTD Revenue Target' for ${totalCount} filtered stores`);
        qtdGapValue && (qtdGapValue.textContent = formatCurrency(sumRevenue - sumQtdTarget));
        qtdGapValue && (qtdGapValue.title = `Calculated Gap (Total Revenue - QTD Target) for ${totalCount} filtered stores`);
        quarterlyRevenueTargetValue && (quarterlyRevenueTargetValue.textContent = formatCurrency(sumQuarterlyTarget));
        quarterlyRevenueTargetValue && (quarterlyRevenueTargetValue.title = `Sum of 'Quarterly Revenue Target' for ${totalCount} filtered stores`);
        unitsWithDFValue && (unitsWithDFValue.textContent = formatNumber(sumUnits));
        unitsWithDFValue && (unitsWithDFValue.title = `Sum of 'Unit w/ DF' for ${totalCount} filtered stores`);
        unitTargetValue && (unitTargetValue.textContent = formatNumber(sumUnitTarget));
        unitTargetValue && (unitTargetValue.title = `Sum of 'Unit Target' for ${totalCount} filtered stores`);
        visitCountValue && (visitCountValue.textContent = formatNumber(sumVisits));
        visitCountValue && (visitCountValue.title = `Sum of 'Visit count' for ${totalCount} filtered stores`);
        trainingCountValue && (trainingCountValue.textContent = formatNumber(sumTrainings));
        trainingCountValue && (trainingCountValue.title = `Sum of 'Trainings' for ${totalCount} filtered stores`);

        percentQuarterlyStoreTargetValue && (percentQuarterlyStoreTargetValue.textContent = formatPercent(overallPercentStoreTarget));
        percentQuarterlyStoreTargetValue && (percentQuarterlyStoreTargetValue.title = `Overall % Quarterly Target (Total Revenue / Total Quarterly Target)`);
        unitAchievementValue && (unitAchievementValue.textContent = formatPercent(overallUnitAchievement));
        unitAchievementValue && (unitAchievementValue.title = `Overall Unit Achievement % (Total Units / Total Unit Target)`);

        revARValue && (revARValue.textContent = formatPercent(avgRevAR));
        revARValue && (revARValue.title = `Average 'Rev AR%' across ${countRevAR} stores with data`);
        retailModeConnectivityValue && (retailModeConnectivityValue.textContent = formatPercent(avgConnectivity));
        retailModeConnectivityValue && (retailModeConnectivityValue.title = `Average 'Retail Mode Connectivity' across ${countConnectivity} stores with data`);
        repSkillAchValue && (repSkillAchValue.textContent = formatPercent(avgRepSkill));
        repSkillAchValue && (repSkillAchValue.title = `Average 'Rep Skill Ach' across ${countRepSkill} stores with data`);
        vPmrAchValue && (vPmrAchValue.textContent = formatPercent(avgPmr));
        vPmrAchValue && (vPmrAchValue.title = `Average '(V)PMR Ach' across ${countPmr} stores with data`);
        postTrainingScoreValue && (postTrainingScoreValue.textContent = formatNumber(avgPostTraining.toFixed(1)));
        postTrainingScoreValue && (postTrainingScoreValue.title = `Average 'Post Training Score' across ${countPostTraining} stores with data`);
        eliteValue && (eliteValue.textContent = formatPercent(avgElite));
        eliteValue && (eliteValue.title = `Average 'Elite' score % across ${countElite} stores with data`);

        updateContextualSummary(data);
    };

    const updateContextualSummary = (data) => {
        [percentQuarterlyTerritoryTargetP, territoryRevPercentP, districtRevPercentP, regionRevPercentP].forEach(p => {if (p) p.style.display = 'none'});
        if (data.length === 0) return;

        const singleRegion = regionFilter.value !== 'ALL';
        const singleDistrict = districtFilter.value !== 'ALL';
        const singleTerritory = territoryFilter.selectedOptions.length === 1;

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

    // Modified updateCharts to use theme-aware options and colors
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

        // Get theme-specific colors from the config
        const { barGoodBg, barGoodBorder, barBadBg, barBadBorder, lineTargetColor, lineTargetBg } = chartOptionsConfig;

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
            options: chartOptionsConfig // Use the centrally managed options
        });
    };

    // --- UPDATED updateAttachRateTable to exclude blanks from footer average ---
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

                 const columns = [
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
                     let formattedValue;
                     if (rawValue === null || (col.key !== 'Store' && isNaN(numericValue))) { formattedValue = 'N/A'; }
                     else { formattedValue = isPercentCol ? col.format(numericValue) : rawValue; }
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
                 let validCount = 0;
                 data.forEach(row => { if (isValidForAverage(safeGet(row, key, null))) validCount++; });
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
         updateAttachRateTable(filteredData); // Use current filtered data
    };


    const updateSortArrows = () => {
        if (!attachRateTable) return;
        attachRateTable.querySelectorAll('th.sortable .sort-arrow').forEach(arrow => {
            arrow.className = 'sort-arrow'; // Remove asc/desc classes
            arrow.textContent = '';
        });
        const currentHeader = attachRateTable.querySelector(`th[data-sort="${currentSort.column}"] .sort-arrow`);
        if (currentHeader) {
            currentHeader.classList.add(currentSort.ascending ? 'asc' : 'desc');
            currentHeader.textContent = currentSort.ascending ? ' ▲' : ' ▼';
        }
    };

    // Updated showStoreDetails to include more fields and Map Link
    const showStoreDetails = (storeData) => {
        if (!storeDetailsContent || !storeDetailsSection || !closeStoreDetailsButton) return;

        const addressParts = [safeGet(storeData, 'ADDRESS1', null), safeGet(storeData, 'CITY', null), safeGet(storeData, 'STATE', null), safeGet(storeData, 'ZIPCODE', null)].filter(Boolean);
        const addressString = addressParts.length > 0 ? addressParts.join(', ') : null;

        const latitude = parseFloat(safeGet(storeData, 'LATITUDE_ORG', NaN));
        const longitude = parseFloat(safeGet(storeData, 'LONGITUDE_ORG', NaN));
        let mapsLinkHtml = '';
        if (!isNaN(latitude) && !isNaN(longitude)) {
            // Using a standard Google Maps URL pattern
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

    const exportData = () => {
        if (filteredData.length === 0) { alert("No filtered data to export."); return; }
        try {
            if (!attachRateTable) throw new Error("Attach rate table not found.");
            const headers = Array.from(attachRateTable.querySelectorAll('thead th')).map(th => th.dataset.sort || th.textContent.replace(/ [▲▼]$/, '').trim());
            const dataToExport = filteredData.map(row => {
                return headers.map(header => {
                    let value = safeGet(row, header, '');
                    const isPercentLike = header.includes('%') || header.includes('Rate') || header.includes('Ach') || header.includes('Connectivity') || header.includes('Elite');
                    if (isPercentLike) { const numVal = parsePercent(value); value = isNaN(numVal) ? '' : numVal; }
                    else { if (typeof value === 'string' && (value.includes(',') || value.includes('"') || value.includes('\n'))) { value = `"${value.replace(/"/g, '""')}"`; } const numVal = parseNumber(value); if (!isPercentLike && !isNaN(numVal) && typeof value !== 'boolean') { value = numVal; } }
                    return value;
                });
             });
            let csvContent = "data:text/csv;charset=utf-8," + headers.join(",") + "\n" + dataToExport.map(e => e.join(",")).join("\n");
            const encodedUri = encodeURI(csvContent);
            const link = document.createElement("a");
            link.setAttribute("href", encodedUri); link.setAttribute("download", "fsm_dashboard_export.csv");
            document.body.appendChild(link); link.click(); document.body.removeChild(link);
        } catch (error) { console.error("Error exporting CSV:", error); alert("Error generating CSV export. See console for details."); }
    };

    const generateEmailBody = () => {
        if (filteredData.length === 0) { return "No data available based on current filters."; }
        let body = "FSM Dashboard Summary:\n";
        body += "---------------------------------\n";
        body += `Filters Applied: ${getFilterSummary()}\n`;
        body += `Stores Found: ${filteredData.length}\n`;
        body += "---------------------------------\n\n";
        body += "Performance Summary:\n";
        body += `- Total Revenue (incl. DF): ${revenueWithDFValue?.textContent || 'N/A'}\n`;
        body += `- QTD Revenue Target: ${qtdRevenueTargetValue?.textContent || 'N/A'}\n`;
        body += `- QTD Gap: ${qtdGapValue?.textContent || 'N/A'}\n`;
        body += `- % Store Quarterly Target: ${percentQuarterlyStoreTargetValue?.textContent || 'N/A'}\n`;
        body += `- Rev AR%: ${revARValue?.textContent || 'N/A'}\n`;
        body += `- Total Units (incl. DF): ${unitsWithDFValue?.textContent || 'N/A'}\n`;
        body += `- Unit Achievement %: ${unitAchievementValue?.textContent || 'N/A'}\n`;
        body += `- Total Visits: ${visitCountValue?.textContent || 'N/A'}\n`;
        body += `- Avg. Connectivity: ${retailModeConnectivityValue?.textContent || 'N/A'}\n\n`;
        body += "Mysteryshop & Training (Avg*):\n";
        body += `- Rep Skill Ach: ${repSkillAchValue?.textContent || 'N/A'}\n`;
        body += `- (V)PMR Ach: ${vPmrAchValue?.textContent || 'N/A'}\n`;
        body += `- Post Training Score: ${postTrainingScoreValue?.textContent || 'N/A'}\n`;
        body += `- Elite Score %: ${eliteValue?.textContent || 'N/A'}\n\n`;
        body += "*Averages calculated only using stores with valid data for each metric.\n\n";

        const sortedFilteredData = [...filteredData].sort((a, b) => {
            let valA = safeGet(a, currentSort.column, null); let valB = safeGet(b, currentSort.column, null);
            if (valA === null && valB === null) return 0; if (valA === null) return currentSort.ascending ? -1 : 1; if (valB === null) return currentSort.ascending ? 1 : -1;
            const isPercentCol = currentSort.column.includes('Attach Rate') || currentSort.column.includes('%') || currentSort.column.includes('Target');
            const numA = isPercentCol ? parsePercent(valA) : parseNumber(valA); const numB = isPercentCol ? parsePercent(valB) : parseNumber(valB);
            if (!isNaN(numA) && !isNaN(numB)) { return currentSort.ascending ? numA - numB : numB - numA; }
            else { valA = String(valA).toLowerCase(); valB = String(valB).toLowerCase(); return currentSort.ascending ? valA.localeCompare(valB) : valB.localeCompare(valA); }
        });
        const topStores = sortedFilteredData.slice(0, 5);
        if (topStores.length > 0) {
            body += `Top ${topStores.length} Stores (Sorted by ${currentSort.column} ${currentSort.ascending ? 'ASC' : 'DESC'}):\n`;
            topStores.forEach((store, index) => {
                 body += `${index + 1}. ${safeGet(store, 'Store', 'N/A')} (% Qtr Rev: ${formatPercent(parsePercent(safeGet(store, '% Quarterly Revenue Target', 0)))}, NCME: ${formatPercent(parsePercent(safeGet(store, 'NCME Attach Rate', 0)))}) \n`;
            });
             body += "\n";
        }
        body += "---------------------------------\n";
        body += "Generated by FSM Dashboard\n";
        return body;
    };

    const getFilterSummary = () => {
        let summary = [];
        if (regionFilter?.value !== 'ALL') summary.push(`Region: ${regionFilter.value}`);
        if (districtFilter?.value !== 'ALL') summary.push(`District: ${districtFilter.value}`);
        const territories = territoryFilter ? Array.from(territoryFilter.selectedOptions).map(o => o.value) : [];
        if (territories.length > 0) summary.push(`Territories: ${territories.length === 1 ? territories[0] : territories.length + ' selected'}`);
        if (fsmFilter?.value !== 'ALL') summary.push(`FSM: ${fsmFilter.value}`);
        if (channelFilter?.value !== 'ALL') summary.push(`Channel: ${channelFilter.value}`);
        if (subchannelFilter?.value !== 'ALL') summary.push(`Subchannel: ${subchannelFilter.value}`);
        if (dealerFilter?.value !== 'ALL') summary.push(`Dealer: ${dealerFilter.value}`);
        const stores = storeFilter ? Array.from(storeFilter.selectedOptions).map(o => o.value) : [];
         if (stores.length > 0) summary.push(`Stores: ${stores.length === 1 ? stores[0] : stores.length + ' selected'}`);
         const flags = Object.entries(flagFiltersCheckboxes).filter(([key, input]) => input && input.checked).map(([key])=> key.replace(/_/g, ' '));
         if (flags.length > 0) summary.push(`Attributes: ${flags.join(', ')}`);
        return summary.length > 0 ? summary.join('; ') : 'None';
    };

    const handleShareEmail = () => {
        if (!emailRecipientInput || !shareStatus) return;
        const recipient = emailRecipientInput.value;
        if (!recipient || !/\S+@\S+\.\S+/.test(recipient)) { shareStatus.textContent = "Please enter a valid recipient email address."; return; }
        try {
            const subject = `FSM Dashboard Summary - ${new Date().toLocaleDateString()}`;
            const body = generateEmailBody();
            const mailtoLink = `mailto:${recipient}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
             if (mailtoLink.length > 2000) { shareStatus.textContent = "Generated email body is too long for mailto link. Try applying more filters."; console.warn("Mailto link length exceeds 2000 characters:", mailtoLink.length); return; }
            window.location.href = mailtoLink;
            shareStatus.textContent = "Email client should open. Please review and send.";
        } catch (error) { console.error("Error generating mailto link:", error); shareStatus.textContent = "Error generating email content."; }
    };

     const selectAllOptions = (selectElement) => {
         if (!selectElement) return;
         Array.from(selectElement.options).forEach(option => option.selected = true);
    };

     const deselectAllOptions = (selectElement) => {
         if (!selectElement) return;
         selectElement.selectedIndex = -1; // Deselects all
    };


    // --- Event Listeners ---
    excelFileInput?.addEventListener('change', handleFile);
    applyFiltersButton?.addEventListener('click', applyFilters);
    resetFiltersButton?.addEventListener('click', handleResetFilters); // Listener for Reset button
    themeToggleButton?.addEventListener('click', toggleTheme); // Listener for Theme button
    storeSearch?.addEventListener('input', filterStoreOptions);
    exportCsvButton?.addEventListener('click', exportData);
    shareEmailButton?.addEventListener('click', handleShareEmail);
    closeStoreDetailsButton?.addEventListener('click', hideStoreDetails);

    territorySelectAll?.addEventListener('click', () => { selectAllOptions(territoryFilter); updateStoreFilterOptionsBasedOnHierarchy(); });
    territoryDeselectAll?.addEventListener('click', () => { deselectAllOptions(territoryFilter); updateStoreFilterOptionsBasedOnHierarchy(); });
    storeSelectAll?.addEventListener('click', () => selectAllOptions(storeFilter));
    storeDeselectAll?.addEventListener('click', () => deselectAllOptions(storeFilter));

    attachRateTable?.querySelector('thead')?.addEventListener('click', handleSort);

    // --- Initial Setup ---
    // Apply saved theme preference on load
    const savedTheme = localStorage.getItem('dashboardTheme') || 'dark'; // Default to dark
    applyTheme(savedTheme);
    // Initial UI reset (don't apply filters automatically here, handleFile does it)
    resetUI(false);

}); // End DOMContentLoaded

