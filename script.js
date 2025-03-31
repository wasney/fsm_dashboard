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
    const resetFiltersButton = document.getElementById('resetFiltersButton');
    const themeToggleButton = document.getElementById('themeToggleButton');
    const themeMetaTag = document.querySelector('meta[name="theme-color"]');

    // Filter Elements
    const regionFilter = document.getElementById('regionFilter');
    const districtFilter = document.getElementById('districtFilter');
    // ** NEW: References to checkbox containers **
    const territoryFilterContainer = document.getElementById('territoryFilterContainer');
    const storeFilterContainer = document.getElementById('storeFilterContainer');
    // Other filters remain the same
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

    // Select/Deselect buttons remain the same
    const territorySelectAll = document.getElementById('territorySelectAll');
    const territoryDeselectAll = document.getElementById('territoryDeselectAll');
    const storeSelectAll = document.getElementById('storeSelectAll');
    const storeDeselectAll = document.getElementById('storeDeselectAll');

    // Summary Elements (No changes needed)
    const revenueWithDFValue = document.getElementById('revenueWithDFValue');
    const qtdRevenueTargetValue = document.getElementById('qtdRevenueTargetValue');
    // ... (rest of summary elements) ...
    const eliteValue = document.getElementById('eliteValue');
    const percentQuarterlyTerritoryTargetP = document.getElementById('percentQuarterlyTerritoryTargetP');
    // ... (rest of contextual summary elements) ...
    const regionRevPercentValue = document.getElementById('regionRevPercentValue');


    // Table Elements (No changes needed)
    const attachRateTableBody = document.getElementById('attachRateTableBody');
    // ... (rest of table elements) ...
    const exportCsvButton = document.getElementById('exportCsvButton');

    // Chart Elements (No changes needed)
    const mainChartCanvas = document.getElementById('mainChartCanvas').getContext('2d');
    // const secondaryChartCanvas = document.getElementById('secondaryChartCanvas').getContext('2d');

    // Store Details Elements (No changes needed)
    const storeDetailsSection = document.getElementById('storeDetailsSection');
    // ... (rest of store detail elements) ...
    const closeStoreDetailsButton = document.getElementById('closeStoreDetailsButton');

    // Share Elements (No changes needed)
    const emailRecipientInput = document.getElementById('emailRecipient');
    // ... (rest of share elements) ...
    const shareStatus = document.getElementById('shareStatus');

    // --- Global State ---
    let rawData = [];
    let filteredData = [];
    let mainChartInstance = null;
    // let secondaryChartInstance = null;
    // storeOptions now holds {value, text} for *available* stores based on hierarchy/search
    let storeOptions = [];
    // allPossibleStores holds {value, text} for *all* stores from the initial file load
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
        // Keep 'ALL' for single-select dropdowns, but maybe not strictly needed for checkbox logic?
        // Let's return without 'ALL' for potential use in checkbox lists.
        // return ['ALL', ...Array.from(values).sort()];
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

    // ** NEW: Function to populate checkbox containers **
    const populateCheckboxMultiSelect = (containerElement, options, disable = false) => {
        if (!containerElement) return;
        containerElement.innerHTML = ''; // Clear existing checkboxes
        containerElement.scrollTop = 0; // Reset scroll position

        if (options.length === 0 && !disable) {
             containerElement.dataset.placeholder = "-- No matching options --";
        } else if (disable) {
            containerElement.dataset.placeholder = "-- Load File First --";
        } else {
            // Remove placeholder when populated
             delete containerElement.dataset.placeholder;
             options.forEach(optionValue => {
                const label = document.createElement('label');
                label.className = 'cb-item';
                label.title = optionValue; // Add tooltip

                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.value = optionValue;
                checkbox.name = containerElement.id; // Group checkboxes logically

                const text = document.createElement('span');
                text.textContent = optionValue;

                label.appendChild(checkbox);
                label.appendChild(text);
                containerElement.appendChild(label);
            });
        }

        // Handle disabled state for the container
        containerElement.classList.toggle('disabled', disable);

        // Disable Select/Deselect buttons based on container state
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

    // --- Theme Handling (No changes needed from previous step) ---
    const applyTheme = (theme) => {
        // ... (keep existing applyTheme function) ...
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
    const toggleTheme = () => { /* ... (keep existing toggleTheme function) ... */
        const currentTheme = document.body.classList.contains('light-mode') ? 'light' : 'dark';
        const newTheme = currentTheme === 'light' ? 'dark' : 'light';
        localStorage.setItem('dashboardTheme', newTheme);
        applyTheme(newTheme);
    };
    function getChartOptions(theme = 'dark') { /* ... (keep existing getChartOptions function) ... */
        const isLightMode = theme === 'light';
        const gridColor = isLightMode ? 'rgba(0, 0, 0, 0.1)' : 'rgba(224, 224, 224, 0.2)';
        // ...(rest of the color definitions)...
        const lineTargetBg = isLightMode ? 'rgba(255, 193, 7, 0.2)' : 'rgba(255, 206, 86, 0.2)';

        return {
            // ...(rest of the chart options)...
             onClick: (event, elements) => {
                 if (elements.length > 0) {
                     const index = elements[0].index;
                     // Ensure mainChartInstance exists and has data before accessing
                     const clickedLabel = mainChartInstance?.data?.labels?.[index];
                     if (!clickedLabel) return; // Exit if label can't be found

                     const storeData = filteredData.find(row => safeGet(row, 'Store', null) === clickedLabel);
                     if (storeData) {
                         showStoreDetails(storeData);
                         highlightTableRow(clickedLabel);
                     }
                 }
             },
            barGoodBg, barGoodBorder, barBadBg, barBadBorder, lineTargetColor, lineTargetBg
        };
    }


    // --- Core Functions ---

    const handleFile = async (event) => { /* ... (keep existing handleFile function structure) ... */
        const file = event.target.files[0];
        if (!file) { statusDiv.textContent = 'No file selected.'; return; }

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

            if (jsonData.length > 0) { /* ... (header validation) ... */
                const headers = Object.keys(jsonData[0]);
                const missingHeaders = REQUIRED_HEADERS.filter(h => !headers.includes(h));
                if (missingHeaders.length > 0) {
                    console.warn(`Warning: Missing expected columns: ${missingHeaders.join(', ')}.`);
                }
            } else { throw new Error("Excel sheet appears to be empty."); }

            rawData = jsonData;
            allPossibleStores = [...new Set(rawData.map(r => safeGet(r, 'Store', null)).filter(s => s))]
                                 .sort()
                                 .map(s => ({ value: s, text: s }));

            statusDiv.textContent = `Loaded ${rawData.length} rows. Applying default filters...`;
            populateFilters(rawData);
            filterArea.style.display = 'block';
            resetFiltersButton.disabled = false;

            applyFilters(); // Auto-apply filters

        } catch (error) { /* ... (error handling) ... */
             console.error('Error processing file:', error);
             statusDiv.textContent = `Error: ${error.message}`;
             rawData = []; allPossibleStores = []; filteredData = [];
             resetUI(true); // Reset fully on error
         } finally { /* ... (finally block) ... */
             showLoading(false);
             excelFileInput.value = '';
         }
    };

    // ** MODIFIED: Populate filters using new checkbox function **
    const populateFilters = (data) => {
        // Populate single-select dropdowns
        setOptions(regionFilter, getUniqueValues(data, 'REGION'));
        setOptions(districtFilter, getUniqueValues(data, 'DISTRICT'));
        setOptions(fsmFilter, getUniqueValues(data, 'FSM NAME'));
        setOptions(channelFilter, getUniqueValues(data, 'CHANNEL'));
        setOptions(subchannelFilter, getUniqueValues(data, 'SUB_CHANNEL'));
        setOptions(dealerFilter, getUniqueValues(data, 'DEALER_NAME'));

        // Populate multi-select checkbox containers
        const territoryOptions = getUniqueValues(data, 'Q2 Territory'); // Get unique values (no 'ALL')
        populateCheckboxMultiSelect(territoryFilterContainer, territoryOptions, false); // Populate and enable

        storeOptions = [...allPossibleStores]; // Start with all stores available
        populateCheckboxMultiSelect(storeFilterContainer, storeOptions.map(s => s.text), false); // Populate and enable stores

        // Enable flag filters and search
        Object.values(flagFiltersCheckboxes).forEach(input => { if(input) input.disabled = false });
        storeSearch.disabled = false;

        // Enable Apply/Reset buttons
        applyFiltersButton.disabled = false;
        resetFiltersButton.disabled = false;

        addDependencyFilterListeners();
    };

    // ** MODIFIED: Add listeners for checkbox containers too **
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
        // ** NEW: Listen for changes within the Territory checkbox container **
        if (territoryFilterContainer) {
             territoryFilterContainer.removeEventListener('change', handler); // Use container for event delegation
             territoryFilterContainer.addEventListener('change', handler);
        }
        // Note: Store filter changes *don't* trigger hierarchy updates, so no listener needed on storeFilterContainer itself for this handler.
    };

    // ** MODIFIED: Update store options based on hierarchy, populating checkboxes **
     const updateStoreFilterOptionsBasedOnHierarchy = () => {
        if (rawData.length === 0) return;

        const selectedRegion = regionFilter.value;
        const selectedDistrict = districtFilter.value;
        // ** NEW: Get selected territories from checkboxes **
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
            // Keep existing filter logic...
            if (selectedRegion !== 'ALL' && safeGet(row, 'REGION', null) !== selectedRegion) return false;
            if (selectedDistrict !== 'ALL' && safeGet(row, 'DISTRICT', null) !== selectedDistrict) return false;
            // ** NEW: Check against selected territory checkboxes **
            if (selectedTerritories.length > 0 && !selectedTerritories.includes(safeGet(row, 'Q2 Territory', null))) return false;
            if (selectedFsm !== 'ALL' && safeGet(row, 'FSM NAME', null) !== selectedFsm) return false;
            if (selectedChannel !== 'ALL' && safeGet(row, 'CHANNEL', null) !== selectedChannel) return false;
            if (selectedSubchannel !== 'ALL' && safeGet(row, 'SUB_CHANNEL', null) !== selectedSubchannel) return false;
            if (selectedDealer !== 'ALL' && safeGet(row, 'DEALER_NAME', null) !== selectedDealer) return false;

            for (const flag in selectedFlags) { /* ... keep flag logic ... */
                const flagValue = safeGet(row, flag, 'NO');
                 if (!(flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1')) {
                    return false;
                 }
             }
            return true;
        });

        const validStoreNames = new Set(potentiallyValidStoresData.map(row => safeGet(row, 'Store', null)).filter(Boolean));
        // Update the global storeOptions array used by search/population
        storeOptions = Array.from(validStoreNames).sort().map(s => ({ value: s, text: s }));

        // ** NEW: Get currently checked stores BEFORE repopulating **
        const previouslyCheckedStores = storeFilterContainer
            ? new Set(Array.from(storeFilterContainer.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value))
            : new Set();

        // ** NEW: Populate store checkbox container (this replaces setStoreFilterOptions) **
        populateCheckboxMultiSelect(storeFilterContainer, storeOptions.map(s => s.text), false); // Update the checkboxes

        // ** NEW: Re-apply the search filter if there was text in the search box **
        filterStoreOptions(); // This updates the *visible* checkboxes based on search term

        // ** NEW: Attempt to re-check previously checked stores *if* they are still visible **
        if (storeFilterContainer) {
            Array.from(storeFilterContainer.querySelectorAll('input[type="checkbox"]')).forEach(checkbox => {
                // Check if this checkbox's value was previously checked AND if its parent label is currently visible (not hidden by search)
                if (previouslyCheckedStores.has(checkbox.value) && checkbox.closest('.cb-item').style.display !== 'none') {
                    checkbox.checked = true;
                } else {
                    checkbox.checked = false; // Ensure others are unchecked
                }
            });
        }
    };

    // ** REMOVED setStoreFilterOptions function (replaced by populateCheckboxMultiSelect) **

    // ** MODIFIED: Filter checkboxes based on search term (hides/shows labels) **
    const filterStoreOptions = () => {
        if (!storeFilterContainer) return;
        const searchTerm = storeSearch.value.toLowerCase();
        let visibleCount = 0;

        Array.from(storeFilterContainer.querySelectorAll('.cb-item')).forEach(item => {
            const labelText = item.textContent.toLowerCase();
            const matches = labelText.includes(searchTerm);
            item.style.display = matches ? 'block' : 'none'; // Show or hide the label
            if (matches) {
                visibleCount++;
            }
        });

        // Update Select/Deselect All buttons based on *visible* options
        storeSelectAll.disabled = storeFilterContainer.classList.contains('disabled') || visibleCount === 0;
        storeDeselectAll.disabled = storeFilterContainer.classList.contains('disabled') || visibleCount === 0;

        // Update placeholder if search yields no results
        if (visibleCount === 0 && searchTerm !== '' && !storeFilterContainer.classList.contains('disabled')) {
            storeFilterContainer.dataset.placeholder = "-- No stores match search --";
        } else if (visibleCount > 0 || storeFilterContainer.classList.contains('disabled')) {
             delete storeFilterContainer.dataset.placeholder; // Remove placeholder if items are visible or container is disabled
        } else if (searchTerm === '' && storeOptions.length === 0 && !storeFilterContainer.classList.contains('disabled')) {
             storeFilterContainer.dataset.placeholder = "-- No matching options --"; // Show no options placeholder
        } else {
             delete storeFilterContainer.dataset.placeholder;
        }

    };


    // ** MODIFIED: Apply filters using selected checkboxes **
    const applyFilters = () => {
        if (rawData.length === 0) { statusDiv.textContent = "Please load a file first."; return; }
        showLoading(true, true);
        resultsArea.style.display = 'none';

        setTimeout(() => {
            try {
                const selectedRegion = regionFilter.value;
                const selectedDistrict = districtFilter.value;
                // ** NEW: Get values from Territory checkboxes **
                const selectedTerritories = territoryFilterContainer
                    ? Array.from(territoryFilterContainer.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value)
                    : [];
                const selectedFsm = fsmFilter.value;
                const selectedChannel = channelFilter.value;
                const selectedSubchannel = subchannelFilter.value;
                const selectedDealer = dealerFilter.value;
                // ** NEW: Get values from Store checkboxes **
                const selectedStores = storeFilterContainer
                    ? Array.from(storeFilterContainer.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value)
                    : [];
                const selectedFlags = {};
                 Object.entries(flagFiltersCheckboxes).forEach(([key, input]) => {
                     if (input && input.checked) { selectedFlags[key] = true; }
                 });

                filteredData = rawData.filter(row => {
                    // ** Update filter logic to use checkbox arrays **
                    if (selectedRegion !== 'ALL' && safeGet(row, 'REGION', null) !== selectedRegion) return false;
                    if (selectedDistrict !== 'ALL' && safeGet(row, 'DISTRICT', null) !== selectedDistrict) return false;
                    if (selectedTerritories.length > 0 && !selectedTerritories.includes(safeGet(row, 'Q2 Territory', null))) return false;
                    if (selectedFsm !== 'ALL' && safeGet(row, 'FSM NAME', null) !== selectedFsm) return false;
                    if (selectedChannel !== 'ALL' && safeGet(row, 'CHANNEL', null) !== selectedChannel) return false;
                    if (selectedSubchannel !== 'ALL' && safeGet(row, 'SUB_CHANNEL', null) !== selectedSubchannel) return false;
                    if (selectedDealer !== 'ALL' && safeGet(row, 'DEALER_NAME', null) !== selectedDealer) return false;
                    if (selectedStores.length > 0 && !selectedStores.includes(safeGet(row, 'Store', null))) return false;

                    for (const flag in selectedFlags) { /* ... keep flag logic ... */
                         const flagValue = safeGet(row, flag, 'NO');
                         if (!(flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1')) {
                            return false;
                         }
                     }
                    return true;
                });

                // Update UI (no changes needed here)
                updateSummary(filteredData);
                updateCharts(filteredData);
                updateAttachRateTable(filteredData);

                if (filteredData.length === 1) { /* ... (show details logic) ... */
                    showStoreDetails(filteredData[0]);
                    highlightTableRow(safeGet(filteredData[0], 'Store', null));
                 } else { hideStoreDetails(); }

                statusDiv.textContent = `Displaying ${filteredData.length} of ${rawData.length} rows based on filters.`;
                resultsArea.style.display = 'block';
                exportCsvButton.disabled = filteredData.length === 0;

            } catch (error) { /* ... (error handling) ... */
                 console.error("Error applying filters:", error);
                 statusDiv.textContent = "Error applying filters. Check console for details.";
                 filteredData = []; resultsArea.style.display = 'none'; exportCsvButton.disabled = true;
                 updateSummary([]); updateCharts([]); updateAttachRateTable([]); hideStoreDetails();
             } finally { showLoading(false, true); }
        }, 10);
    };

    // ** MODIFIED: Reset function to handle checkboxes **
    const resetFilters = () => {
        // Reset dropdowns
         [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(sel => { if (sel) { sel.value = 'ALL'; } });

         // ** NEW: Uncheck all checkboxes in containers **
         if (territoryFilterContainer) {
            territoryFilterContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
         }
         if (storeFilterContainer) {
            storeFilterContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
         }

         // Reset search input
         if (storeSearch) { storeSearch.value = ''; }

         // Reset flag checkboxes
         Object.values(flagFiltersCheckboxes).forEach(input => { if(input) {input.checked = false;} });

         // IMPORTANT: Update the dynamic store filter based on the reset hierarchy filters
         updateStoreFilterOptionsBasedOnHierarchy();
         // After updating store options, clear the search filter visually
         // filterStoreOptions(); // updateStoreFilterOptions... calls this now

         // Reset sort state for the table
         currentSort = { column: 'Store', ascending: true };
    };

     const handleResetFilters = () => {
         resetFilters();
         applyFilters();
     };

     const resetUI = (apply = true) => { /* ... (keep existing resetUI, ensuring it disables containers if needed) ... */
         resetFilters();

         const disable = rawData.length === 0;
         // Disable standard elements
         [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter, storeSearch, territorySelectAll, territoryDeselectAll, storeSelectAll, storeDeselectAll, applyFiltersButton, resetFiltersButton, exportCsvButton].forEach(el => {if(el) el.disabled = disable});
         Object.values(flagFiltersCheckboxes).forEach(input => { if(input) input.disabled = disable });
         // Disable checkbox containers
         if (territoryFilterContainer) populateCheckboxMultiSelect(territoryFilterContainer, [], disable);
         if (storeFilterContainer) populateCheckboxMultiSelect(storeFilterContainer, [], disable); // Use the population function to handle disable/placeholder


         if (disable) { statusDiv.textContent = 'No file selected.'; }

         if (filterArea) filterArea.style.display = disable ? 'none' : 'block';
         if (resultsArea) resultsArea.style.display = 'none';
         // ...(clear chart, table, details, summary)...
         if (mainChartInstance) { mainChartInstance.destroy(); mainChartInstance = null; }
         if (attachRateTableBody) attachRateTableBody.innerHTML = '';
         if (attachRateTableFooter) attachRateTableFooter.innerHTML = '';
         if (attachTableStatus) attachTableStatus.textContent = '';
         hideStoreDetails();
         updateSummary([]);

         // Remove dependency listeners
         const handler = updateStoreFilterOptionsBasedOnHierarchy;
          [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter, territoryFilterContainer].forEach(el => { // Include territory container
              if (el) el.removeEventListener('change', handler);
          });
          Object.values(flagFiltersCheckboxes).forEach(input => {
             if (input) input.removeEventListener('change', handler);
          });


         if (apply && !disable) { applyFilters(); }
     };

    // --- UPDATED updateSummary (No changes needed here) ---
    const updateSummary = (data) => { /* ... (keep existing updateSummary) ... */ };

    // --- UPDATED updateContextualSummary (No changes needed here) ---
    const updateContextualSummary = (data) => { /* ... (keep existing updateContextualSummary) ... */ };

    // --- UPDATED updateCharts (No changes needed here, theme handled separately) ---
    const updateCharts = (data) => { /* ... (keep existing updateCharts) ... */ };

    // --- UPDATED updateAttachRateTable (No changes needed here) ---
    const updateAttachRateTable = (data) => { /* ... (keep existing updateAttachRateTable) ... */ };

    // --- UPDATED handleSort (No changes needed here) ---
    const handleSort = (event) => { /* ... (keep existing handleSort) ... */ };

    // --- UPDATED updateSortArrows (No changes needed here) ---
    const updateSortArrows = () => { /* ... (keep existing updateSortArrows) ... */ };

    // --- UPDATED showStoreDetails (No changes needed here) ---
    const showStoreDetails = (storeData) => { /* ... (keep existing showStoreDetails) ... */ };

    // --- UPDATED hideStoreDetails (No changes needed here) ---
    const hideStoreDetails = () => { /* ... (keep existing hideStoreDetails) ... */ };

     // --- UPDATED highlightTableRow (No changes needed here) ---
     const highlightTableRow = (storeName) => { /* ... (keep existing highlightTableRow) ... */ };

    // --- UPDATED exportData (No changes needed here) ---
    const exportData = () => { /* ... (keep existing exportData) ... */ };

    // ** MODIFIED: getFilterSummary to read from checkboxes **
    const getFilterSummary = () => {
        let summary = [];
        if (regionFilter?.value !== 'ALL') summary.push(`Region: ${regionFilter.value}`);
        if (districtFilter?.value !== 'ALL') summary.push(`District: ${districtFilter.value}`);
        // ** NEW: Read from territory checkboxes **
        const territories = territoryFilterContainer
            ? Array.from(territoryFilterContainer.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value)
            : [];
        if (territories.length > 0) summary.push(`Territories: ${territories.length === 1 ? territories[0] : territories.length + ' selected'}`);

        if (fsmFilter?.value !== 'ALL') summary.push(`FSM: ${fsmFilter.value}`);
        if (channelFilter?.value !== 'ALL') summary.push(`Channel: ${channelFilter.value}`);
        if (subchannelFilter?.value !== 'ALL') summary.push(`Subchannel: ${subchannelFilter.value}`);
        if (dealerFilter?.value !== 'ALL') summary.push(`Dealer: ${dealerFilter.value}`);
        // ** NEW: Read from store checkboxes **
        const stores = storeFilterContainer
             ? Array.from(storeFilterContainer.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value)
             : [];
         if (stores.length > 0) summary.push(`Stores: ${stores.length === 1 ? stores[0] : stores.length + ' selected'}`);

         const flags = Object.entries(flagFiltersCheckboxes).filter(([key, input]) => input && input.checked).map(([key])=> key.replace(/_/g, ' '));
         if (flags.length > 0) summary.push(`Attributes: ${flags.join(', ')}`);
        return summary.length > 0 ? summary.join('; ') : 'None';
    };

    // --- UPDATED generateEmailBody (uses getFilterSummary, no direct changes needed) ---
    const generateEmailBody = () => { /* ... (keep existing generateEmailBody) ... */ };

    // --- UPDATED handleShareEmail (No changes needed here) ---
    const handleShareEmail = () => { /* ... (keep existing handleShareEmail) ... */ };

     // ** MODIFIED: Select/Deselect All for Checkbox Containers **
     const selectAllOptions = (containerElementOrSelect) => {
         if (!containerElementOrSelect) return;
         // Check if it's one of our checkbox containers
         if (containerElementOrSelect.classList.contains('checkbox-multiselect')) {
             containerElementOrSelect.querySelectorAll('input[type="checkbox"]').forEach(cb => {
                 // Only check visible checkboxes (respecting search filter)
                 if (cb.closest('.cb-item').style.display !== 'none') {
                     cb.checked = true;
                 }
             });
             // Trigger change event manually on the container if needed for dependency updates
             containerElementOrSelect.dispatchEvent(new Event('change', { bubbles: true }));
         } else if (containerElementOrSelect.tagName === 'SELECT') { // Handle original select elements if any remain
              Array.from(containerElementOrSelect.options).forEach(option => option.selected = true);
              containerElementOrSelect.dispatchEvent(new Event('change', { bubbles: true }));
         }
    };

     const deselectAllOptions = (containerElementOrSelect) => {
         if (!containerElementOrSelect) return;
         if (containerElementOrSelect.classList.contains('checkbox-multiselect')) {
             containerElementOrSelect.querySelectorAll('input[type="checkbox"]').forEach(cb => {
                 // Only deselect visible checkboxes
                  if (cb.closest('.cb-item').style.display !== 'none') {
                      cb.checked = false;
                  }
             });
             containerElementOrSelect.dispatchEvent(new Event('change', { bubbles: true }));
         } else if (containerElementOrSelect.tagName === 'SELECT') {
              containerElementOrSelect.selectedIndex = -1;
              containerElementOrSelect.dispatchEvent(new Event('change', { bubbles: true }));
         }
    };


    // --- Event Listeners ---
    excelFileInput?.addEventListener('change', handleFile);
    applyFiltersButton?.addEventListener('click', applyFilters);
    resetFiltersButton?.addEventListener('click', handleResetFilters);
    themeToggleButton?.addEventListener('click', toggleTheme);
    storeSearch?.addEventListener('input', filterStoreOptions); // Keep listener on search input
    exportCsvButton?.addEventListener('click', exportData);
    shareEmailButton?.addEventListener('click', handleShareEmail);
    closeStoreDetailsButton?.addEventListener('click', hideStoreDetails);

    // ** MODIFIED: Point Select/Deselect to new checkbox containers **
    territorySelectAll?.addEventListener('click', () => { selectAllOptions(territoryFilterContainer); /* Dependency handled by container listener */ });
    territoryDeselectAll?.addEventListener('click', () => { deselectAllOptions(territoryFilterContainer); /* Dependency handled by container listener */ });
    storeSelectAll?.addEventListener('click', () => selectAllOptions(storeFilterContainer)); // No hierarchy update needed here
    storeDeselectAll?.addEventListener('click', () => deselectAllOptions(storeFilterContainer)); // No hierarchy update needed here

    attachRateTable?.querySelector('thead')?.addEventListener('click', handleSort);

    // --- Initial Setup ---
    const savedTheme = localStorage.getItem('dashboardTheme') || 'dark';
    applyTheme(savedTheme);
    resetUI(false);

}); // End DOMContentLoaded
