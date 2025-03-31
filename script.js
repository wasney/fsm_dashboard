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
    const DEBUG_ROW_COUNT = 5; // How many rows to log details for in the filter

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

    // Summary Elements (remain unchanged)
    const revenueWithDFValue = document.getElementById('revenueWithDFValue');
    // ... rest of summary elements

    // Table Elements (remain unchanged)
    const attachRateTableBody = document.getElementById('attachRateTableBody');
    // ... rest of table elements

    // Chart Elements (remain unchanged)
    const mainChartCanvas = document.getElementById('mainChartCanvas').getContext('2d');
    // ...

    // Store Details Elements (remain unchanged)
    const storeDetailsSection = document.getElementById('storeDetailsSection');
    // ...

    // Share Elements (remain unchanged)
    const emailRecipientInput = document.getElementById('emailRecipient');
    // ...

    // --- Global State ---
    let rawData = [];
    let filteredData = [];
    let mainChartInstance = null;
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
    // ** Reverted safeGet default back to empty string '' **
    const safeGet = (obj, path, defaultValue = '') => {
        const value = obj ? obj[path] : undefined;
        return (value !== undefined && value !== null) ? value : defaultValue;
    };

    const isValidForAverage = (value) => {
         if (value === null || value === undefined || value === '') return false;
         return !isNaN(parseNumber(String(value).replace('%','')));
    };
    const getUniqueValues = (data, column) => {
        const values = new Set(data.map(item => safeGet(item, column, '')).filter(val => val !== ''));
        return [...Array.from(values).sort()];
    };
    const setOptions = (selectElement, options, disable = false) => { /* ... (unchanged) ... */
        selectElement.innerHTML = '';
        const allOption = document.createElement('option');
        allOption.value = 'ALL';
        allOption.textContent = '-- ALL --';
        selectElement.appendChild(allOption);
        options.forEach(optionValue => { /* ... */ });
         selectElement.disabled = disable;
         selectElement.value = 'ALL';
     };

    const populateCheckboxMultiSelect = (containerElement, options, disable = false) => { /* ... (unchanged) ... */
        if (!containerElement) return;
        containerElement.innerHTML = '';
        containerElement.scrollTop = 0;
        if (options.length === 0 && !disable) containerElement.dataset.placeholder = "-- No matching options --";
        else if (disable) containerElement.dataset.placeholder = "-- Load File First --";
        else {
             delete containerElement.dataset.placeholder;
             options.forEach(optionValue => { /* ... create labels/checkboxes ... */ });
        }
        containerElement.classList.toggle('disabled', disable);
        const selectAllBtn = document.getElementById(`${containerElement.id.replace('Container', '')}SelectAll`);
        const deselectAllBtn = document.getElementById(`${containerElement.id.replace('Container', '')}DeselectAll`);
        if (selectAllBtn) selectAllBtn.disabled = disable || options.length === 0;
        if (deselectAllBtn) deselectAllBtn.disabled = disable || options.length === 0;
    };


    const showLoading = (isLoading, isFiltering = false) => { /* ... (unchanged) ... */ };

    // --- Theme Handling ---
    const applyTheme = (theme) => { /* ... (unchanged) ... */ };
    const toggleTheme = () => { /* ... (unchanged) ... */ };
    function getChartOptions(theme = 'dark') { /* ... (unchanged, includes colorDefaults) ... */ }


    // --- Core Functions ---

    const handleFile = async (event) => { /* ... (unchanged, ensures String conversion for allPossibleStores) ... */
        const file = event.target.files[0];
        if (!file) { statusDiv.textContent = 'No file selected.'; return; }

        statusDiv.textContent = 'Reading file...';
        showLoading(true); filterArea.style.display = 'none'; resultsArea.style.display = 'none'; resetUI(false);

        try {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data);
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: null }); // Use null default here

            if (jsonData.length > 0) { /* header validation */ }
            else { throw new Error("Excel sheet appears to be empty."); }

            rawData = jsonData;
            // Ensure store names/values are strings for consistent comparison later
            allPossibleStores = [...new Set(rawData.map(r => safeGet(r, 'Store', '')).filter(s => s !== ''))]
                                 .sort()
                                 .map(s => ({ value: String(s), text: String(s) }));

            statusDiv.textContent = `Loaded ${rawData.length} rows. Applying default filters...`;
            populateFilters(rawData);
            filterArea.style.display = 'block'; resetFiltersButton.disabled = false;
            applyFilters();

        } catch (error) { /* error handling */ }
        finally { /* finally block */ }
    };

    const populateFilters = (data) => { /* ... (unchanged, uses String map for territory options) ... */
        setOptions(regionFilter, getUniqueValues(data, 'REGION'));
        setOptions(districtFilter, getUniqueValues(data, 'DISTRICT'));
        setOptions(fsmFilter, getUniqueValues(data, 'FSM NAME'));
        setOptions(channelFilter, getUniqueValues(data, 'CHANNEL'));
        setOptions(subchannelFilter, getUniqueValues(data, 'SUB_CHANNEL'));
        setOptions(dealerFilter, getUniqueValues(data, 'DEALER_NAME'));
        // Ensure territory options are strings
        const territoryOptions = getUniqueValues(data, 'Q2 Territory').map(String);
        populateCheckboxMultiSelect(territoryFilterContainer, territoryOptions, false);
        storeOptions = [...allPossibleStores];
        populateCheckboxMultiSelect(storeFilterContainer, storeOptions.map(s => s.text), false);
        Object.values(flagFiltersCheckboxes).forEach(input => { if(input) input.disabled = false });
        storeSearch.disabled = false;
        applyFiltersButton.disabled = false; resetFiltersButton.disabled = false;
        addDependencyFilterListeners();
    };

    const addDependencyFilterListeners = () => { /* ... (unchanged) ... */ };

     const updateStoreFilterOptionsBasedOnHierarchy = () => { /* ... (unchanged, uses String comparisons) ... */
        if (rawData.length === 0) return;
        const selectedRegion = regionFilter.value;
        const selectedDistrict = districtFilter.value;
        const selectedTerritories = territoryFilterContainer ? Array.from(territoryFilterContainer.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value) : [];
        // ... other filters ...
        const potentiallyValidStoresData = rawData.filter(row => {
            if (selectedRegion !== 'ALL' && String(safeGet(row, 'REGION', '')) !== selectedRegion) return false;
            if (selectedDistrict !== 'ALL' && String(safeGet(row, 'DISTRICT', '')) !== selectedDistrict) return false;
            if (selectedTerritories.length > 0 && !selectedTerritories.includes(String(safeGet(row, 'Q2 Territory', '')))) return false;
            // ... rest of filter conditions ...
            return true;
        });
        const validStoreNames = new Set(potentiallyValidStoresData.map(row => String(safeGet(row, 'Store', ''))).filter(Boolean));
        storeOptions = Array.from(validStoreNames).sort().map(s => ({ value: s, text: s }));
        const previouslyCheckedStores = storeFilterContainer ? new Set(Array.from(storeFilterContainer.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value)) : new Set();
        populateCheckboxMultiSelect(storeFilterContainer, storeOptions.map(s => s.text), false);
        filterStoreOptions();
        if (storeFilterContainer) { /* re-check logic */ }
     };

    const filterStoreOptions = () => { /* ... (unchanged) ... */ };


    // ** ADDED DEBUGGING LOGS to applyFilters **
    const applyFilters = () => {
        if (rawData.length === 0) { statusDiv.textContent = "Please load a file first."; return; }
        showLoading(true, true);
        resultsArea.style.display = 'none';

        // --- DEBUGGING START ---
        console.log("--- Applying Filters ---");
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

         console.log("Selected Region:", selectedRegion);
         console.log("Selected District:", selectedDistrict);
         console.log("Selected Territories:", selectedTerritories);
         console.log("Selected FSM:", selectedFsm);
         console.log("Selected Channel:", selectedChannel);
         console.log("Selected Subchannel:", selectedSubchannel);
         console.log("Selected Dealer:", selectedDealer);
         console.log("Selected Stores:", selectedStores);
         console.log("Selected Flags:", selectedFlags);
         console.log("Raw Data Rows:", rawData.length);
        // --- DEBUGGING END ---


        setTimeout(() => {
            try {
                // Filter rawData
                filteredData = rawData.filter((row, index) => { // Added index for logging limit
                    // Get stringified row values (using safeGet default '')
                    const rowRegion = String(safeGet(row, 'REGION', ''));
                    const rowDistrict = String(safeGet(row, 'DISTRICT', ''));
                    const rowTerritory = String(safeGet(row, 'Q2 Territory', ''));
                    const rowFsm = String(safeGet(row, 'FSM NAME', ''));
                    const rowChannel = String(safeGet(row, 'CHANNEL', ''));
                    const rowSubchannel = String(safeGet(row, 'SUB_CHANNEL', ''));
                    const rowDealer = String(safeGet(row, 'DEALER_NAME', ''));
                    const rowStore = String(safeGet(row, 'Store', ''));

                    // Conditions
                    const regionMatch = selectedRegion === 'ALL' || rowRegion === selectedRegion;
                    const districtMatch = selectedDistrict === 'ALL' || rowDistrict === selectedDistrict;
                    const territoryMatch = selectedTerritories.length === 0 || selectedTerritories.includes(rowTerritory);
                    const fsmMatch = selectedFsm === 'ALL' || rowFsm === selectedFsm;
                    const channelMatch = selectedChannel === 'ALL' || rowChannel === selectedChannel;
                    const subchannelMatch = selectedSubchannel === 'ALL' || rowSubchannel === selectedSubchannel;
                    const dealerMatch = selectedDealer === 'ALL' || rowDealer === selectedDealer;
                    const storeMatch = selectedStores.length === 0 || selectedStores.includes(rowStore);

                    let flagsMatch = true;
                    for (const flag in selectedFlags) {
                        const flagValue = safeGet(row, flag, 'NO');
                        const isTrue = (flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1');
                        if (!isTrue) {
                            flagsMatch = false;
                            break;
                        }
                    }

                    const finalMatch = regionMatch && districtMatch && territoryMatch && fsmMatch && channelMatch && subchannelMatch && dealerMatch && storeMatch && flagsMatch;

                    // --- DEBUGGING START (Log details for first few rows) ---
                    if (index < DEBUG_ROW_COUNT) {
                         console.log(`--- Row ${index} Data ---`);
                         console.log(` Store: ${rowStore}, Territory: ${rowTerritory}, Region: ${rowRegion}, District: ${rowDistrict}`);
                         console.log(` Conditions Met: Region=${regionMatch}, District=${districtMatch}, Territory=${territoryMatch}, FSM=${fsmMatch}, Channel=${channelMatch}, Subchannel=${subchannelMatch}, Dealer=${dealerMatch}, Store=${storeMatch}, Flags=${flagsMatch}`);
                         console.log(` Final Match for Row ${index}:`, finalMatch);
                    }
                    // --- DEBUGGING END ---

                    return finalMatch; // Row passes all filters
                });

                 // --- DEBUGGING START ---
                 console.log("Filtered Data Rows:", filteredData.length);
                 if (filteredData.length === 0 && rawData.length > 0) {
                     console.warn("Filtering resulted in zero rows. Check filter logic and data consistency.");
                 }
                 // --- DEBUGGING END ---


                // Update UI sections
                updateSummary(filteredData);
                updateCharts(filteredData);
                updateAttachRateTable(filteredData);

                // Show/hide store details
                if (filteredData.length === 1) { /* ... */ }
                else { hideStoreDetails(); }

                // Final UI updates
                statusDiv.textContent = `Displaying ${filteredData.length} of ${rawData.length} rows based on filters.`;
                resultsArea.style.display = 'block';
                exportCsvButton.disabled = filteredData.length === 0;

            } catch (error) { /* ... error handling ... */ }
            finally { showLoading(false, true); }
        }, 10);
    };

    const resetFilters = () => { /* ... (unchanged) ... */ };
     const handleResetFilters = () => { /* ... (unchanged) ... */ };
     const resetUI = (apply = true) => { /* ... (unchanged) ... */ };
    const updateSummary = (data) => { /* ... (unchanged) ... */ };
    const updateContextualSummary = (data) => { /* ... (unchanged) ... */ };
    const updateCharts = (data) => { /* ... (unchanged) ... */ };
    const updateAttachRateTable = (data) => { /* ... (unchanged) ... */ };
    const handleSort = (event) => { /* ... (unchanged) ... */ };
    const updateSortArrows = () => { /* ... (unchanged) ... */ };
    const showStoreDetails = (storeData) => { /* ... (unchanged) ... */ };
    const hideStoreDetails = () => { /* ... (unchanged) ... */ };
     const highlightTableRow = (storeName) => { /* ... (unchanged) ... */ };
    const exportData = () => { /* ... (unchanged) ... */ };
    const getFilterSummary = () => { /* ... (unchanged) ... */ };
    const generateEmailBody = () => { /* ... (unchanged) ... */ };
    const handleShareEmail = () => { /* ... (unchanged) ... */ };
     const selectAllOptions = (containerElement) => { /* ... (unchanged) ... */ };
     const deselectAllOptions = (containerElement) => { /* ... (unchanged) ... */ };

    // --- Event Listeners ---
    // ... (remain unchanged) ...

    // --- Initial Setup ---
    const savedTheme = localStorage.getItem('dashboardTheme') || 'dark';
    applyTheme(savedTheme);
    resetUI(false);

}); // End DOMContentLoaded
