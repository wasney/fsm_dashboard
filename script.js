/* Generated: 2025-04-01 12:57:45 AM EDT - Re-attempt Attach Rate table filtering: Hide rows with all 0% rates *during* row creation, preserving original data for sorting/averages. */
document.addEventListener('DOMContentLoaded', () => {
    // --- Configuration ---
    const REQUIRED_HEADERS = [ // Add all essential headers needed for calculations/display
        'Store', 'REGION', 'DISTRICT', 'Q2 Territory', 'FSM NAME', 'CHANNEL',
        'SUB_CHANNEL', 'DEALER_NAME', 'Revenue w/DF', 'QTD Revenue Target',
        'Quarterly Revenue Target', 'QTD Gap', '% Quarterly Revenue Target', //'Rev AR%', // Rev AR% is calculated
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
    const ATTACH_RATE_KEYS = [ // Keys used for the attach rate table filtering and averages
         'Tablet Attach Rate', 'PC Attach Rate', 'NC Attach Rate',
         'TWS Attach Rate', 'WW Attach Rate', 'ME Attach Rate', 'NCME Attach Rate'
    ];
    const CURRENCY_FORMAT = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' });
    const PERCENT_FORMAT = new Intl.NumberFormat('en-US', { style: 'percent', minimumFractionDigits: 1, maximumFractionDigits: 1 });
    const NUMBER_FORMAT = new Intl.NumberFormat('en-US');
    const CHART_COLORS = ['#58a6ff', '#ffb758', '#86dc86', '#ff7f7f', '#b796e6', '#ffda8a', '#8ad7ff', '#ff9ba6'];
    const TOP_N_CHART = 15; // Max items to show on the bar chart
    const TOP_N_TABLES = 5; // Items for Top/Bottom 5 tables

    // --- DOM Elements ---
    const excelFileInput = document.getElementById('excelFile');
    const statusDiv = document.getElementById('status'); // Reference to the moved status element
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

    // CORRECTED FLAG FILTER CHECKBOX MAPPING
    const flagFiltersCheckboxes = FLAG_HEADERS.reduce((acc, header) => {
        let expectedId = '';
        // Manually map headers to their specific HTML IDs
        switch (header) {
            case 'SUPER STORE':       expectedId = 'superStoreFilter'; break;
            case 'GOLDEN RHINO':      expectedId = 'goldenRhinoFilter'; break;
            case 'GCE':               expectedId = 'gceFilter'; break;
            case 'AI_Zone':           expectedId = 'aiZoneFilter'; break; // Match exact case from HTML
            case 'Hispanic_Market':   expectedId = 'hispanicMarketFilter'; break; // Match exact case from HTML
            case 'EV ROUTE':          expectedId = 'evRouteFilter'; break;
            default:
                console.warn(`Unknown flag header encountered: ${header}`);
                return acc; // Skip unknown headers
        }

        const element = document.getElementById(expectedId);
        if (element) {
            acc[header] = element; // Store the element reference using the original header name as key
        } else {
            // This warning should now only appear if the HTML ID genuinely doesn't exist or HTML is not loaded
            console.warn(`Flag filter checkbox not found for ID: ${expectedId} (Header: ${header})`);
        }
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
    const revARValue = document.getElementById('revARValue'); // This now represents calculated Rev % QTD Target
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
    // Contextual Summary Elements & Paragraphs
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

    // ** NEW ** Top/Bottom 5 Elements
    const topBottomSection = document.getElementById('topBottomSection');
    const top5TableBody = document.getElementById('top5TableBody');
    const bottom5TableBody = document.getElementById('bottom5TableBody');

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
    let filteredData = []; // Data after *all* filters (including store) are applied
    let mainChartInstance = null;
    // let secondaryChartInstance = null; // Placeholder
    let storeOptions = []; // Holds the *currently available* store options {value, text} based on other filters
    let allPossibleStores = []; // Holds *all* store options {value, text} from the initial file load
    let currentSort = { column: 'Store', ascending: true }; // For Attach Rate Table
    let selectedStoreRow = null; // Tracks the currently highlighted row across *all* tables

    // --- Helper Functions ---
    const formatCurrency = (value) => isNaN(value) ? 'N/A' : CURRENCY_FORMAT.format(value);
    const formatPercent = (value) => isNaN(value) ? 'N/A' : PERCENT_FORMAT.format(value);
    const formatNumber = (value) => isNaN(value) ? 'N/A' : NUMBER_FORMAT.format(value);

    // Enhanced parsing to handle potentially null/empty values better
    const parseNumber = (value) => {
        if (value === null || value === undefined || value === '') return NaN; // Treat blanks/nulls as NaN
        if (typeof value === 'number') return value;
        if (typeof value === 'string') {
            value = value.replace(/[$,%]/g, '');
            const num = parseFloat(value);
            return isNaN(num) ? NaN : num; // Return NaN if parsing fails
        }
        return NaN; // Default to NaN for other types
    };
    const parsePercent = (value) => {
         if (value === null || value === undefined || value === '') return NaN; // Treat blanks/nulls as NaN
         if (typeof value === 'number') return value; // Assume it's already decimal
         if (typeof value === 'string') {
            const num = parseFloat(value.replace('%', ''));
            return isNaN(num) ? NaN : num / 100; // Convert to decimal, NaN if fails
         }
         return NaN; // Default to NaN
    };
    const safeGet = (obj, path, defaultValue = 'N/A') => {
        // Special handling for 0: if defaultValue is explicitly null, return 0 as 0, not null.
        if (defaultValue === null && obj && obj[path] === 0) {
             return 0;
        }
        const value = obj ? obj[path] : undefined;
        return (value !== undefined && value !== null) ? value : defaultValue;
    };
    // Helper to check if a value is valid for averaging (not null, not empty, parses to number)
    const isValidForAverage = (value) => {
         if (value === null || value === undefined || value === '') return false;
         return !isNaN(parseNumber(String(value).replace('%','')));
    };
    // Helper to check if a value is valid for averaging AND non-zero
    const isValidAndNonZeroForAverage = (value) => {
         if (!isValidForAverage(value)) return false;
         const num = parseNumber(String(value).replace('%',''));
         return num !== 0;
    };
    // Function to calculate QTD Gap for sorting/display
    const calculateQtdGap = (row) => {
         const revenue = parseNumber(safeGet(row, 'Revenue w/DF', 0));
         const target = parseNumber(safeGet(row, 'QTD Revenue Target', 0));
         if (isNaN(revenue) || isNaN(target)) { return Infinity; } // Treat undefined gaps as highest for ascending sort
         return revenue - target;
    };
    // Function to calculate Rev AR% for display in Top/Bottom tables
    const calculateRevARPercent = (row) => {
        const revenue = parseNumber(safeGet(row, 'Revenue w/DF', 0));
        const target = parseNumber(safeGet(row, 'QTD Revenue Target', 0));
        if (isNaN(revenue) || isNaN(target) || target === 0) { return NaN; }
        return revenue / target;
    };
    // Function to calculate Unit Achievement % for display in Top/Bottom tables
    const calculateUnitAchievementPercent = (row) => {
        const units = parseNumber(safeGet(row, 'Unit w/ DF', 0));
        const target = parseNumber(safeGet(row, 'Unit Target', 0));
        if (isNaN(units) || isNaN(target) || target === 0) { return NaN; }
        return units / target;
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
    };
    const showLoading = (isLoading, isFiltering = false) => {
        if (isFiltering) {
            filterLoadingIndicator.style.display = isLoading ? 'flex' : 'none';
            applyFiltersButton.disabled = isLoading;
        } else {
            loadingIndicator.style.display = isLoading ? 'flex' : 'none';
            excelFileInput.disabled = isLoading;
        }
    };

    // --- Core Functions ---

    const handleFile = async (event) => {
        const file = event.target.files[0];
        if (!file) { statusDiv.textContent = 'No file selected.'; return; }
        statusDiv.textContent = 'Reading file...';
        showLoading(true);
        filterArea.style.display = 'none';
        resultsArea.style.display = 'none';
        resetFilters();
        try {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data);
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: null });
            if (jsonData.length > 0) {
                const headers = Object.keys(jsonData[0]);
                const missingHeaders = REQUIRED_HEADERS.filter(h => !headers.includes(h));
                if (missingHeaders.length > 0) { console.warn(`Warning: Missing expected columns: ${missingHeaders.join(', ')}.`); }
            } else { throw new Error("Excel sheet appears to be empty."); }
            rawData = jsonData;
            allPossibleStores = [...new Set(rawData.map(r => safeGet(r, 'Store', null)).filter(s => s))].sort().map(s => ({ value: s, text: s }));
            statusDiv.textContent = `Loaded ${rawData.length} rows. Adjust filters and click 'Apply Filters'.`;
            populateFilters(rawData);
            filterArea.style.display = 'block';
        } catch (error) {
            console.error('Error processing file:', error);
            statusDiv.textContent = `Error: ${error.message}`;
            rawData = []; allPossibleStores = []; filteredData = []; resetUI();
        } finally {
            showLoading(false); excelFileInput.value = '';
        }
    };

    const populateFilters = (data) => {
        setOptions(regionFilter, getUniqueValues(data, 'REGION'));
        setOptions(districtFilter, getUniqueValues(data, 'DISTRICT'));
        setMultiSelectOptions(territoryFilter, getUniqueValues(data, 'Q2 Territory').slice(1));
        setOptions(fsmFilter, getUniqueValues(data, 'FSM NAME'));
        setOptions(channelFilter, getUniqueValues(data, 'CHANNEL'));
        setOptions(subchannelFilter, getUniqueValues(data, 'SUB_CHANNEL'));
        setOptions(dealerFilter, getUniqueValues(data, 'DEALER_NAME'));
        Object.values(flagFiltersCheckboxes).forEach(input => { if(input) input.disabled = false });
        storeOptions = [...allPossibleStores];
        setStoreFilterOptions(storeOptions, false);
        territorySelectAll.disabled = false; territoryDeselectAll.disabled = false;
        storeSelectAll.disabled = false; storeDeselectAll.disabled = false;
        storeSearch.disabled = false; applyFiltersButton.disabled = false;
        addDependencyFilterListeners();
    };

    const addDependencyFilterListeners = () => {
        const handler = updateStoreFilterOptionsBasedOnHierarchy;
        [regionFilter, districtFilter, territoryFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(filter => {
            if (filter) { filter.removeEventListener('change', handler); filter.addEventListener('change', handler); }
        });
        Object.values(flagFiltersCheckboxes).forEach(input => {
             if (input) { input.removeEventListener('change', handler); input.addEventListener('change', handler); }
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
        const selectedFlags = {}; Object.entries(flagFiltersCheckboxes).forEach(([key, input]) => { if (input && input.checked) selectedFlags[key] = true; });
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
                if (!(flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1')) return false;
            }
            return true;
        });
        const validStoreNames = new Set(potentiallyValidStoresData.map(row => safeGet(row, 'Store', null)).filter(Boolean));
        storeOptions = Array.from(validStoreNames).sort().map(s => ({ value: s, text: s }));
        const previouslySelectedStores = new Set(Array.from(storeFilter.selectedOptions).map(opt => opt.value));
        setStoreFilterOptions(storeOptions, false);
        filterStoreOptions();
        Array.from(storeFilter.options).forEach(option => { if (previouslySelectedStores.has(option.value)) option.selected = true; });
        if (storeFilter.selectedOptions.length === 0) storeFilter.selectedIndex = -1;
    };

    const setStoreFilterOptions = (optionsToShow, disable = true) => {
        const currentSearchTerm = storeSearch.value;
        storeFilter.innerHTML = '';
        optionsToShow.forEach(opt => {
             const option = document.createElement('option'); option.value = opt.value; option.textContent = opt.text; option.title = opt.text; storeFilter.appendChild(option);
         });
        storeFilter.disabled = disable; storeSearch.disabled = disable;
        storeSelectAll.disabled = disable || optionsToShow.length === 0; storeDeselectAll.disabled = disable || optionsToShow.length === 0;
        storeSearch.value = currentSearchTerm;
    };

    const filterStoreOptions = () => {
        const searchTerm = storeSearch.value.toLowerCase();
        const filteredOptions = storeOptions.filter(opt => opt.text.toLowerCase().includes(searchTerm));
        const selectedValues = new Set(Array.from(storeFilter.selectedOptions).map(opt => opt.value));
        storeFilter.innerHTML = '';
        filteredOptions.forEach(opt => {
            const option = document.createElement('option'); option.value = opt.value; option.textContent = opt.text; option.title = opt.text;
             if (selectedValues.has(opt.value)) option.selected = true;
            storeFilter.appendChild(option);
        });
        storeSelectAll.disabled = storeFilter.disabled || filteredOptions.length === 0;
        storeDeselectAll.disabled = storeFilter.disabled || filteredOptions.length === 0;
    };

    const applyFilters = () => {
        showLoading(true, true); resultsArea.style.display = 'none';
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
                const selectedFlags = {}; Object.entries(flagFiltersCheckboxes).forEach(([key, input]) => { if (input && input.checked) selectedFlags[key] = true; });
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
                        if (!(flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1')) return false;
                    }
                    return true;
                });
                updateSummary(filteredData);
                updateTopBottomTables(filteredData); // Update Top/Bottom tables
                updateCharts(filteredData);
                updateAttachRateTable(filteredData); // Update Attach Rate table
                if (filteredData.length === 1) { showStoreDetails(filteredData[0]); highlightTableRow(safeGet(filteredData[0], 'Store', null)); }
                else { hideStoreDetails(); }
                statusDiv.textContent = `Displaying ${filteredData.length} of ${rawData.length} rows based on filters.`;
                resultsArea.style.display = 'block'; exportCsvButton.disabled = filteredData.length === 0;
            } catch (error) {
                console.error("Error applying filters:", error); statusDiv.textContent = "Error applying filters. Check console for details.";
                filteredData = []; resultsArea.style.display = 'none'; exportCsvButton.disabled = true;
                updateSummary([]); updateTopBottomTables([]); updateCharts([]); updateAttachRateTable([]); hideStoreDetails();
            } finally { showLoading(false, true); }
        }, 10);
    };

    const resetFilters = () => {
        [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(sel => { if (sel) { sel.value = 'ALL'; sel.disabled = true;} });
        if (territoryFilter) { territoryFilter.selectedIndex = -1; territoryFilter.disabled = true; }
        if (storeFilter) { storeFilter.innerHTML = '<option value="ALL">-- Load File First --</option>'; storeFilter.selectedIndex = -1; storeFilter.disabled = true; }
        if (storeSearch) { storeSearch.value = ''; storeSearch.disabled = true; }
        storeOptions = []; allPossibleStores = [];
        Object.values(flagFiltersCheckboxes).forEach(input => { if(input) {input.checked = false; input.disabled = true;} });
        if (applyFiltersButton) applyFiltersButton.disabled = true;
        if (territorySelectAll) territorySelectAll.disabled = true; if (territoryDeselectAll) territoryDeselectAll.disabled = true;
        if (storeSelectAll) storeSelectAll.disabled = true; if (storeDeselectAll) storeDeselectAll.disabled = true;
        if (exportCsvButton) exportCsvButton.disabled = true;
        const handler = updateStoreFilterOptionsBasedOnHierarchy;
        [regionFilter, districtFilter, territoryFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(filter => { if (filter) filter.removeEventListener('change', handler); });
        Object.values(flagFiltersCheckboxes).forEach(input => { if (input) input.removeEventListener('change', handler); });
    };

    const resetUI = () => {
        resetFilters(); if (filterArea) filterArea.style.display = 'none'; if (resultsArea) resultsArea.style.display = 'none';
        if (mainChartInstance) { mainChartInstance.destroy(); mainChartInstance = null; }
        if (attachRateTableBody) attachRateTableBody.innerHTML = ''; if (attachRateTableFooter) attachRateTableFooter.innerHTML = '';
        if (attachTableStatus) attachTableStatus.textContent = '';
        if (topBottomSection) topBottomSection.style.display = 'none'; if (top5TableBody) top5TableBody.innerHTML = ''; if (bottom5TableBody) bottom5TableBody.innerHTML = '';
        hideStoreDetails(); updateSummary([]); if(statusDiv) statusDiv.textContent = 'No file selected.';
    };

    const updateSummary = (data) => {
        const totalCount = data.length;
        const fieldsToClear = [revenueWithDFValue, qtdRevenueTargetValue, qtdGapValue, quarterlyRevenueTargetValue, percentQuarterlyStoreTargetValue, revARValue, unitsWithDFValue, unitTargetValue, unitAchievementValue, visitCountValue, trainingCountValue, retailModeConnectivityValue, repSkillAchValue, vPmrAchValue, postTrainingScoreValue, eliteValue, percentQuarterlyTerritoryTargetValue, territoryRevPercentValue, districtRevPercentValue, regionRevPercentValue];
        fieldsToClear.forEach(el => { if (el) el.textContent = 'N/A'; });
        [percentQuarterlyTerritoryTargetP, territoryRevPercentP, districtRevPercentP, regionRevPercentP].forEach(p => { if(p) p.style.display = 'none';});
        if (totalCount === 0) { return; }
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
            let val;
            val = safeGet(row, 'Retail Mode Connectivity', null); if (isValidAndNonZeroForAverage(val)) { sumConnectivity += parsePercent(val); countConnectivity++; }
            val = safeGet(row, 'Rep Skill Ach', null); if (isValidAndNonZeroForAverage(val)) { sumRepSkill += parsePercent(val); countRepSkill++; }
            val = safeGet(row, '(V)PMR Ach', null); if (isValidAndNonZeroForAverage(val)) { sumPmr += parsePercent(val); countPmr++; }
            val = safeGet(row, 'Post Training Score', null); if (isValidForAverage(val)) { sumPostTraining += parseNumber(val); countPostTraining++; }
            val = safeGet(row, 'Elite', null); if (isValidForAverage(val)) { sumElite += parsePercent(val); countElite++; }
        });
        const avgConnectivity = countConnectivity === 0 ? NaN : sumConnectivity / countConnectivity;
        const avgRepSkill = countRepSkill === 0 ? NaN : sumRepSkill / countRepSkill;
        const avgPmr = countPmr === 0 ? NaN : sumPmr / countPmr;
        const avgPostTraining = countPostTraining === 0 ? NaN : sumPostTraining / countPostTraining;
        const avgElite = countElite === 0 ? NaN : sumElite / countElite;
        const calculatedRevAR = sumQtdTarget === 0 ? 0 : sumRevenue / sumQtdTarget;
        const overallUnitAchievement = sumUnitTarget === 0 ? 0 : sumUnits / sumUnitTarget;
        const overallPercentStoreTarget = sumQuarterlyTarget === 0 ? 0 : sumRevenue / sumQuarterlyTarget;
        // Update DOM Elements... [rest of the summary DOM updates unchanged]
        revenueWithDFValue && (revenueWithDFValue.textContent = formatCurrency(sumRevenue)); revenueWithDFValue && (revenueWithDFValue.title = `Sum of 'Revenue w/DF' for ${totalCount} filtered stores`);
        qtdRevenueTargetValue && (qtdRevenueTargetValue.textContent = formatCurrency(sumQtdTarget)); qtdRevenueTargetValue && (qtdRevenueTargetValue.title = `Sum of 'QTD Revenue Target' for ${totalCount} filtered stores`);
        qtdGapValue && (qtdGapValue.textContent = formatCurrency(sumRevenue - sumQtdTarget)); qtdGapValue && (qtdGapValue.title = `Calculated Gap (Total Revenue - QTD Target) for ${totalCount} filtered stores`);
        quarterlyRevenueTargetValue && (quarterlyRevenueTargetValue.textContent = formatCurrency(sumQuarterlyTarget)); quarterlyRevenueTargetValue && (quarterlyRevenueTargetValue.title = `Sum of 'Quarterly Revenue Target' for ${totalCount} filtered stores`);
        unitsWithDFValue && (unitsWithDFValue.textContent = formatNumber(sumUnits)); unitsWithDFValue && (unitsWithDFValue.title = `Sum of 'Unit w/ DF' for ${totalCount} filtered stores`);
        unitTargetValue && (unitTargetValue.textContent = formatNumber(sumUnitTarget)); unitTargetValue && (unitTargetValue.title = `Sum of 'Unit Target' for ${totalCount} filtered stores`);
        visitCountValue && (visitCountValue.textContent = formatNumber(sumVisits)); visitCountValue && (visitCountValue.title = `Sum of 'Visit count' for ${totalCount} filtered stores`);
        trainingCountValue && (trainingCountValue.textContent = formatNumber(sumTrainings)); trainingCountValue && (trainingCountValue.title = `Sum of 'Trainings' for ${totalCount} filtered stores`);
        revARValue && (revARValue.textContent = formatPercent(calculatedRevAR)); revARValue && (revARValue.title = `Calculated Rev AR% (Total Revenue / Total QTD Target)`);
        unitAchievementValue && (unitAchievementValue.textContent = formatPercent(overallUnitAchievement)); unitAchievementValue && (unitAchievementValue.title = `Overall Unit Achievement % (Total Units / Total Unit Target)`);
        percentQuarterlyStoreTargetValue && (percentQuarterlyStoreTargetValue.textContent = formatPercent(overallPercentStoreTarget)); percentQuarterlyStoreTargetValue && (percentQuarterlyStoreTargetValue.title = `Overall % Quarterly Target (Total Revenue / Total Quarterly Target)`);
        retailModeConnectivityValue && (retailModeConnectivityValue.textContent = formatPercent(avgConnectivity)); retailModeConnectivityValue && (retailModeConnectivityValue.title = `Average 'Retail Mode Connectivity' across ${countConnectivity} stores with non-zero data`);
        repSkillAchValue && (repSkillAchValue.textContent = formatPercent(avgRepSkill)); repSkillAchValue && (repSkillAchValue.title = `Average 'Rep Skill Ach' across ${countRepSkill} stores with non-zero data`);
        vPmrAchValue && (vPmrAchValue.textContent = formatPercent(avgPmr)); vPmrAchValue && (vPmrAchValue.title = `Average '(V)PMR Ach' across ${countPmr} stores with non-zero data`);
        postTrainingScoreValue && (postTrainingScoreValue.textContent = formatNumber(avgPostTraining.toFixed(1))); postTrainingScoreValue && (postTrainingScoreValue.title = `Average 'Post Training Score' across ${countPostTraining} stores with data`);
        eliteValue && (eliteValue.textContent = formatPercent(avgElite)); eliteValue && (eliteValue.title = `Average 'Elite' score % across ${countElite} stores with data`);
        updateContextualSummary(data);
    };

    const updateContextualSummary = (data) => {
        [percentQuarterlyTerritoryTargetP, territoryRevPercentP, districtRevPercentP, regionRevPercentP].forEach(p => {if (p) p.style.display = 'none'});
        if (data.length === 0) return;
        const singleRegion = regionFilter.value !== 'ALL'; const singleDistrict = districtFilter.value !== 'ALL'; const singleTerritory = territoryFilter.selectedOptions.length === 1;
        const calculateAverageExcludeBlanks = (column) => { let sum = 0; let count = 0; data.forEach(row => { const val = safeGet(row, column, null); if (isValidForAverage(val)) { sum += parsePercent(val); count++; } }); return count === 0 ? NaN : sum / count; };
        const avgPercentTerritoryTarget = calculateAverageExcludeBlanks('%Quarterly Territory Rev Target'); const avgTerritoryRevPercent = calculateAverageExcludeBlanks('Territory Rev%'); const avgDistrictRevPercent = calculateAverageExcludeBlanks('District Rev%'); const avgRegionRevPercent = calculateAverageExcludeBlanks('Region Rev%');
        if (percentQuarterlyTerritoryTargetValue) percentQuarterlyTerritoryTargetValue.textContent = formatPercent(avgPercentTerritoryTarget); if (percentQuarterlyTerritoryTargetP) percentQuarterlyTerritoryTargetP.style.display = 'block';
        if (singleTerritory || singleDistrict || singleRegion) { if (territoryRevPercentValue) territoryRevPercentValue.textContent = formatPercent(avgTerritoryRevPercent); if (territoryRevPercentP) territoryRevPercentP.style.display = 'block'; }
        if (singleDistrict || singleRegion) { if (districtRevPercentValue) districtRevPercentValue.textContent = formatPercent(avgDistrictRevPercent); if (districtRevPercentP) districtRevPercentP.style.display = 'block'; }
        if (singleRegion) { if (regionRevPercentValue) regionRevPercentValue.textContent = formatPercent(avgRegionRevPercent); if (regionRevPercentP) regionRevPercentP.style.display = 'block'; }
    };

    const updateTopBottomTables = (data) => {
        if (!topBottomSection || !top5TableBody || !bottom5TableBody) return;
        top5TableBody.innerHTML = ''; bottom5TableBody.innerHTML = '';
        const territoriesInData = new Set(data.map(row => safeGet(row, 'Q2 Territory', null)).filter(Boolean));
        const isSingleTerritory = territoriesInData.size === 1;
        if (!isSingleTerritory || data.length === 0) { topBottomSection.style.display = 'none'; return; }
        topBottomSection.style.display = 'flex';
        // Top 5
        const top5Data = [...data].sort((a, b) => parseNumber(safeGet(b, 'Revenue w/DF', -Infinity)) - parseNumber(safeGet(a, 'Revenue w/DF', -Infinity))).slice(0, TOP_N_TABLES);
        top5Data.forEach(row => {
            const tr = top5TableBody.insertRow(); const storeName = safeGet(row, 'Store', null); tr.dataset.storeName = storeName;
            tr.onclick = () => { showStoreDetails(row); highlightTableRow(storeName); };
            const revenue = parseNumber(safeGet(row, 'Revenue w/DF', NaN)); const revAR = calculateRevARPercent(row); const unitAch = calculateUnitAchievementPercent(row); const visits = parseNumber(safeGet(row, 'Visit count', NaN));
            tr.insertCell().textContent = storeName; tr.insertCell().textContent = formatCurrency(revenue); tr.insertCell().textContent = formatPercent(revAR); tr.insertCell().textContent = formatPercent(unitAch); tr.insertCell().textContent = formatNumber(visits);
        });
        // Bottom 5
        const bottom5Data = [...data].sort((a, b) => calculateQtdGap(a) - calculateQtdGap(b)).slice(0, TOP_N_TABLES); // Corrected sort: ascending gap
        bottom5Data.forEach(row => {
            const tr = bottom5TableBody.insertRow(); const storeName = safeGet(row, 'Store', null); tr.dataset.storeName = storeName;
            tr.onclick = () => { showStoreDetails(row); highlightTableRow(storeName); };
            const qtdGap = calculateQtdGap(row); const revAR = calculateRevARPercent(row); const unitAch = calculateUnitAchievementPercent(row); const visits = parseNumber(safeGet(row, 'Visit count', NaN));
            tr.insertCell().textContent = storeName; tr.insertCell().textContent = formatCurrency(qtdGap === Infinity ? NaN : qtdGap); tr.insertCell().textContent = formatPercent(revAR); tr.insertCell().textContent = formatPercent(unitAch); tr.insertCell().textContent = formatNumber(visits);
        });
    };

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
             type: 'bar', data: { labels: labels, datasets: [ { label: 'Total Revenue (incl. DF)', data: revenueDataSet, backgroundColor: backgroundColors, borderColor: borderColors, borderWidth: 1 }, { label: 'QTD Revenue Target', data: targetDataSet, type: 'line', borderColor: 'rgba(255, 206, 86, 1)', backgroundColor: 'rgba(255, 206, 86, 0.2)', fill: false, tension: 0.1, borderWidth: 2, pointRadius: 3, pointHoverRadius: 5 } ] },
             options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, ticks: { color: '#e0e0e0', callback: (value) => formatCurrency(value) }, grid: { color: 'rgba(224, 224, 224, 0.2)' } }, x: { ticks: { color: '#e0e0e0' }, grid: { display: false } } }, plugins: { legend: { labels: { color: '#e0e0e0' } }, tooltip: { callbacks: { label: function(context) { let label = context.dataset.label || ''; if (label) { label += ': '; } if (context.parsed.y !== null) { if (context.dataset.type === 'line' || context.dataset.label.toLowerCase().includes('target')) { label += formatCurrency(context.parsed.y); } else { label += formatCurrency(context.parsed.y); if (chartData && chartData[context.dataIndex]){ const storeData = chartData[context.dataIndex]; const percentTarget = parsePercent(safeGet(storeData, '% Quarterly Revenue Target', 0)); label += ` (${formatPercent(percentTarget)} of Qtr Target)`; } } } return label; } } } },
             onClick: (event, elements) => { if (elements.length > 0) { const index = elements[0].index; const storeName = labels[index]; const storeData = filteredData.find(row => safeGet(row, 'Store', null) === storeName); if (storeData) { showStoreDetails(storeData); highlightTableRow(storeName); } } } }
        });
    };

    // ** UPDATED ** Apply 0% filtering during row creation
    const updateAttachRateTable = (data) => {
        if (!attachRateTableBody || !attachRateTableFooter) return;
        attachRateTableBody.innerHTML = '';
        attachRateTableFooter.innerHTML = '';

        const originalDataLength = data.length;
        if (originalDataLength === 0) {
            if(attachTableStatus) attachTableStatus.textContent = 'No data to display based on filters.';
            return;
        }

        // 1. Sort the original data first
        const sortedData = [...data].sort((a, b) => {
             let valA = safeGet(a, currentSort.column, null); let valB = safeGet(b, currentSort.column, null);
             if (valA === null && valB === null) return 0; if (valA === null) return currentSort.ascending ? -1 : 1; if (valB === null) return currentSort.ascending ? 1 : -1;
             const isPercentCol = currentSort.column.includes('Attach Rate'); // Simplified check for attach rate table
             const numA = isPercentCol ? parsePercent(valA) : parseNumber(valA); const numB = isPercentCol ? parsePercent(valB) : parseNumber(valB);
             if (!isNaN(numA) && !isNaN(numB)) { return currentSort.ascending ? numA - numB : numB - numA; }
             else { valA = String(valA).toLowerCase(); valB = String(valB).toLowerCase(); return currentSort.ascending ? valA.localeCompare(valB) : valB.localeCompare(valA); }
         });

        // 2. Calculate averages based on the original data
        const averages = {};
        ATTACH_RATE_KEYS.forEach(key => {
             let sum = 0; let count = 0;
             data.forEach(row => {
                 const val = safeGet(row, key, null);
                 if (isValidForAverage(val)) { sum += parsePercent(val); count++; }
             });
             averages[key] = count === 0 ? NaN : sum / count;
         });

        let rowsAdded = 0; // Counter for displayed rows

        // 3. Iterate through sorted data and filter *before* adding row to table
        sortedData.forEach(row => {
            const storeName = safeGet(row, 'Store', null);
            if (storeName) {
                // Check if *any* attach rate is > 0 for this row
                const shouldShowRow = ATTACH_RATE_KEYS.some(key => {
                    const val = parsePercent(safeGet(row, key, null)); // Use null default
                    return !isNaN(val) && val > 0;
                });

                if (shouldShowRow) {
                    rowsAdded++; // Increment counter only if row is shown
                    const tr = document.createElement('tr');
                    tr.dataset.storeName = storeName;
                    tr.onclick = () => { showStoreDetails(row); highlightTableRow(storeName); };
                    // Define columns (manually or use ATTACH_RATE_KEYS)
                     const columns = [
                         { key: 'Store', format: (val) => val },
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
                        const isPercentCol = col.key.includes('Attach Rate');
                        const numericValue = isPercentCol ? parsePercent(rawValue) : parseNumber(rawValue);
                        let formattedValue = (rawValue === null || (col.key !== 'Store' && isNaN(numericValue))) ? 'N/A' : (isPercentCol ? col.format(numericValue) : rawValue);
                        td.textContent = formattedValue;
                        td.title = `${col.key}: ${formattedValue}`;
                        if (col.highlight && averages[col.key] !== undefined && !isNaN(numericValue)) {
                            td.classList.add(numericValue >= averages[col.key] ? 'highlight-green' : 'highlight-red');
                        }
                        tr.appendChild(td);
                    });
                    attachRateTableBody.appendChild(tr);
                }
            }
        });

        // 4. Add Average Row to Footer (calculated from original 'data')
        if (originalDataLength > 0) {
            const footerRow = attachRateTableFooter.insertRow();
            const avgLabelCell = footerRow.insertCell();
            avgLabelCell.textContent = 'Filtered Avg*';
            avgLabelCell.title = 'Average calculated only using stores with valid data for each column from the initial filter selection';
            avgLabelCell.style.textAlign = "right"; avgLabelCell.style.fontWeight = "bold";
            ATTACH_RATE_KEYS.forEach(key => {
                 const td = footerRow.insertCell();
                 const avgValue = averages[key]; td.textContent = formatPercent(avgValue);
                 let validCount = 0; data.forEach(row => { if (isValidForAverage(safeGet(row, key, null))) validCount++; });
                 td.title = `Average ${key}: ${formatPercent(avgValue)} (from ${validCount} stores)`; td.style.textAlign = "right";
             });
         }

        // 5. Update status text based on the count of rows *actually added*
        if(attachTableStatus) attachTableStatus.textContent = `Showing ${rowsAdded} stores with non-zero attach rates. Click row for details. Click headers to sort.`;
        updateSortArrows();
    };


    const handleSort = (event) => {
        const headerCell = event.target.closest('th'); if (!headerCell || !headerCell.classList.contains('sortable')) return;
        const sortKey = headerCell.dataset.sort; if (!sortKey) return;
        if (currentSort.column === sortKey) { currentSort.ascending = !currentSort.ascending; }
        else { currentSort.column = sortKey; currentSort.ascending = true; }
        updateAttachRateTable(filteredData); // Re-render table with new sort
    };

    const updateSortArrows = () => {
        if (!attachRateTable) return;
        attachRateTable.querySelectorAll('th.sortable .sort-arrow').forEach(arrow => { arrow.classList.remove('asc', 'desc'); arrow.textContent = ''; });
        try { const currentHeader = attachRateTable.querySelector(`th[data-sort="${CSS.escape(currentSort.column)}"] .sort-arrow`); if (currentHeader) { currentHeader.classList.add(currentSort.ascending ? 'asc' : 'desc'); currentHeader.textContent = currentSort.ascending ? ' ▲' : ' ▼'; }
        } catch (e) { console.error("Error updating sort arrows:", e); }
    };

    const showStoreDetails = (storeData) => {
        if (!storeDetailsContent || !storeDetailsSection || !closeStoreDetailsButton) return;
        const addressParts = [safeGet(storeData, 'ADDRESS1', null), safeGet(storeData, 'CITY', null), safeGet(storeData, 'STATE', null), safeGet(storeData, 'ZIPCODE', null)].filter(Boolean);
        const addressString = addressParts.length > 0 ? addressParts.join(', ') : null;
        const latitude = parseFloat(safeGet(storeData, 'LATITUDE_ORG', NaN)); const longitude = parseFloat(safeGet(storeData, 'LONGITUDE_ORG', NaN));
        let mapsLinkHtml = (!isNaN(latitude) && !isNaN(longitude)) ? `<p><a href="https://www.google.com/maps?q=$${latitude},${longitude}" target="_blank" title="Open in Google Maps">View on Google Maps</a></p>` : `<p style="color: #aaa; font-style: italic;">(Map coordinates not available)</p>`;
        let flagSummaryHtml = FLAG_HEADERS.map(flag => { const flagValue = safeGet(storeData, flag); const isTrue = (flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1'); return `<span title="${flag}" data-flag="${isTrue}">${flag.replace(/_/g, ' ')} ${isTrue ? '✔' : '✘'}</span>`; }).join(' | ');
        let detailsHtml = `<p><strong>Store:</strong> ${safeGet(storeData, 'Store')}</p>`;
        if (addressString) { detailsHtml += `<p><strong>Address:</strong> ${addressString}</p>`; }
        detailsHtml += mapsLinkHtml + `<hr>`;
        detailsHtml += `<p><strong>IDs:</strong> Store: ${safeGet(storeData, 'STORE ID')} | Org: ${safeGet(storeData, 'ORG_STORE_ID')} | CV: ${safeGet(storeData, 'CV_STORE_ID')} | CinglePoint: ${safeGet(storeData, 'CINGLEPOINT_ID')}</p>`;
        detailsHtml += `<p><strong>Type:</strong> ${safeGet(storeData, 'STORE_TYPE_NAME')} | Nat Tier: ${safeGet(storeData, 'National_Tier')} | Merch Lvl: ${safeGet(storeData, 'Merchandising_Level')} | Comb Tier: ${safeGet(storeData, 'Combined_Tier')}</p><hr>`;
        detailsHtml += `<p><strong>Hierarchy:</strong> ${safeGet(storeData, 'REGION')} > ${safeGet(storeData, 'DISTRICT')} > ${safeGet(storeData, 'Q2 Territory')}</p>`;
        detailsHtml += `<p><strong>FSM:</strong> ${safeGet(storeData, 'FSM NAME')}</p>`;
        detailsHtml += `<p><strong>Channel:</strong> ${safeGet(storeData, 'CHANNEL')} / ${safeGet(storeData, 'SUB_CHANNEL')}</p>`;
        detailsHtml += `<p><strong>Dealer:</strong> ${safeGet(storeData, 'DEALER_NAME')}</p><hr>`;
        detailsHtml += `<p><strong>Visits:</strong> ${formatNumber(parseNumber(safeGet(storeData, 'Visit count', 0)))} | <strong>Trainings:</strong> ${formatNumber(parseNumber(safeGet(storeData, 'Trainings', 0)))}</p>`;
        detailsHtml += `<p><strong>Connectivity:</strong> ${formatPercent(parsePercent(safeGet(storeData, 'Retail Mode Connectivity', 0)))}</p><hr>`;
        detailsHtml += `<p><strong>Flags:</strong> ${flagSummaryHtml}</p>`;
        storeDetailsContent.innerHTML = detailsHtml; storeDetailsSection.style.display = 'block'; closeStoreDetailsButton.style.display = 'inline-block';
        storeDetailsSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    };

    const hideStoreDetails = () => {
        if (!storeDetailsContent || !storeDetailsSection || !closeStoreDetailsButton) return;
        storeDetailsContent.innerHTML = 'Select a store from the table or chart for details, or apply filters resulting in a single store.';
        storeDetailsSection.style.display = 'none'; closeStoreDetailsButton.style.display = 'none';
        highlightTableRow(null);
    };

    const highlightTableRow = (storeName) => {
        if (selectedStoreRow) { selectedStoreRow.classList.remove('selected-row'); selectedStoreRow = null; }
        if (storeName) {
            const tables = [attachRateTableBody, top5TableBody, bottom5TableBody];
            for (const tableBody of tables) {
                if (tableBody) {
                    try { const rowToHighlight = tableBody.querySelector(`tr[data-store-name="${CSS.escape(storeName)}"]`); if (rowToHighlight) { rowToHighlight.classList.add('selected-row'); selectedStoreRow = rowToHighlight; break; }
                    } catch (e) { console.error("Error selecting table row:", e); }
                }
            }
        }
    };

    const exportData = () => {
        if (filteredData.length === 0) { alert("No filtered data to export."); return; }
        try {
            if (!attachRateTable) throw new Error("Attach rate table not found.");
            // Use Attach Rate Table headers for export still
             const headers = Array.from(attachRateTable.querySelectorAll('thead th')).map(th => th.dataset.sort || th.textContent.replace(/ [▲▼]$/, '').trim());
             // Export based on the main filteredData, not the visually filtered attach rate table
             const dataToExport = filteredData.map(row => {
                 return headers.map(header => {
                     let dataKey = header; // Assume header matches key initially
                     // Simplified mapping based on the attach rate table headers
                     if (header === 'Tablet') dataKey = 'Tablet Attach Rate'; else if (header === 'PC') dataKey = 'PC Attach Rate'; else if (header === 'NC') dataKey = 'NC Attach Rate'; else if (header === 'TWS') dataKey = 'TWS Attach Rate'; else if (header === 'WW') dataKey = 'WW Attach Rate'; else if (header === 'ME') dataKey = 'ME Attach Rate'; else if (header === 'NCME') dataKey = 'NCME Attach Rate';
                     let value = safeGet(row, dataKey, '');
                     const isPercentLike = dataKey.includes('%') || dataKey.includes('Rate') || dataKey.includes('Ach') || dataKey.includes('Connectivity') || dataKey.includes('Elite');
                     if (isPercentLike) { const numVal = parsePercent(value); value = isNaN(numVal) ? '' : numVal; }
                     else { if (typeof value === 'string' && (value.includes(',') || value.includes('"') || value.includes('\n'))) { value = `"${value.replace(/"/g, '""')}"`; } const numVal = parseNumber(value); if (!isPercentLike && !isNaN(numVal) && typeof value !== 'boolean') { value = numVal; } }
                     return value;
                 });
             });
             let csvContent = "data:text/csv;charset=utf-8," + headers.join(",") + "\n" + dataToExport.map(e => e.join(",")).join("\n");
             const encodedUri = encodeURI(csvContent); const link = document.createElement("a"); link.setAttribute("href", encodedUri); link.setAttribute("download", "fsm_dashboard_export.csv"); document.body.appendChild(link); link.click(); document.body.removeChild(link);
        } catch (error) { console.error("Error exporting CSV:", error); alert("Error generating CSV export. See console for details."); }
    };

    const generateEmailBody = () => {
        if (filteredData.length === 0) { return "No data available based on current filters."; }
        let body = "FSM Dashboard Summary:\n---------------------------------\n";
        body += `Filters Applied: ${getFilterSummary()}\nStores Found: ${filteredData.length}\n---------------------------------\n\n`;
        body += "Performance Summary:\n"; // ... [rest of summary unchanged] ...
        body += `- Total Revenue (incl. DF): ${revenueWithDFValue?.textContent || 'N/A'}\n`; body += `- QTD Revenue Target: ${qtdRevenueTargetValue?.textContent || 'N/A'}\n`; body += `- QTD Gap: ${qtdGapValue?.textContent || 'N/A'}\n`; body += `- % Store Quarterly Target: ${percentQuarterlyStoreTargetValue?.textContent || 'N/A'}\n`; body += `- Rev AR%: ${revARValue?.textContent || 'N/A'}\n`; body += `- Total Units (incl. DF): ${unitsWithDFValue?.textContent || 'N/A'}\n`; body += `- Unit Achievement %: ${unitAchievementValue?.textContent || 'N/A'}\n`; body += `- Total Visits: ${visitCountValue?.textContent || 'N/A'}\n`; body += `- Avg. Connectivity: ${retailModeConnectivityValue?.textContent || 'N/A'}\n\n`; body += "Mysteryshop & Training (Avg*):\n"; body += `- Rep Skill Ach: ${repSkillAchValue?.textContent || 'N/A'}\n`; body += `- (V)PMR Ach: ${vPmrAchValue?.textContent || 'N/A'}\n`; body += `- Post Training Score: ${postTrainingScoreValue?.textContent || 'N/A'}\n`; body += `- Elite Score %: ${eliteValue?.textContent || 'N/A'}\n\n`; body += "*Averages calculated only using stores with valid, non-zero data for each metric where applicable (Connectivity, Rep Skill, PMR).\n\n";
        const territoriesInData = new Set(filteredData.map(row => safeGet(row, 'Q2 Territory', null)).filter(Boolean));
        if (territoriesInData.size === 1) {
             const territoryName = territoriesInData.values().next().value; body += `--- Territory: ${territoryName} ---\n`;
             const top5Data = [...filteredData].sort((a, b) => parseNumber(safeGet(b, 'Revenue w/DF', -Infinity)) - parseNumber(safeGet(a, 'Revenue w/DF', -Infinity))).slice(0, TOP_N_TABLES);
             if (top5Data.length > 0) { body += "Top 5 (Revenue):\n"; top5Data.forEach((row, i) => { body += `${i+1}. ${safeGet(row, 'Store')} - Rev: ${formatCurrency(parseNumber(safeGet(row, 'Revenue w/DF')))}\n`; }); body += "\n"; }
             const bottom5Data = [...filteredData].sort((a, b) => calculateQtdGap(a) - calculateQtdGap(b)).slice(0, TOP_N_TABLES);
             if (bottom5Data.length > 0) { body += "Bottom 5 (Opportunities by QTD Gap):\n"; bottom5Data.forEach((row, i) => { body += `${i+1}. ${safeGet(row, 'Store')} - Gap: ${formatCurrency(calculateQtdGap(row))}\n`; }); body += "\n"; }
        }
        body += "---------------------------------\nGenerated by FSM Dashboard\n";
        return body;
    };

    const getFilterSummary = () => {
        let summary = [];
        if (regionFilter?.value !== 'ALL') summary.push(`Region: ${regionFilter.value}`); if (districtFilter?.value !== 'ALL') summary.push(`District: ${districtFilter.value}`); const territories = territoryFilter ? Array.from(territoryFilter.selectedOptions).map(o => o.value) : []; if (territories.length > 0) summary.push(`Territories: ${territories.length === 1 ? territories[0] : territories.length + ' selected'}`); if (fsmFilter?.value !== 'ALL') summary.push(`FSM: ${fsmFilter.value}`); if (channelFilter?.value !== 'ALL') summary.push(`Channel: ${channelFilter.value}`); if (subchannelFilter?.value !== 'ALL') summary.push(`Subchannel: ${subchannelFilter.value}`); if (dealerFilter?.value !== 'ALL') summary.push(`Dealer: ${dealerFilter.value}`); const stores = storeFilter ? Array.from(storeFilter.selectedOptions).map(o => o.value) : []; if (stores.length > 0) summary.push(`Stores: ${stores.length === 1 ? stores[0] : stores.length + ' selected'}`); const flags = Object.entries(flagFiltersCheckboxes).filter(([_, input]) => input && input.checked).map(([key])=> key.replace(/_/g, ' ')); if (flags.length > 0) summary.push(`Attributes: ${flags.join(', ')}`);
        return summary.length > 0 ? summary.join('; ') : 'None';
    };

    const handleShareEmail = () => {
        if (!emailRecipientInput || !shareStatus) return; const recipient = emailRecipientInput.value;
        if (!recipient || !/\S+@\S+\.\S+/.test(recipient)) { shareStatus.textContent = "Please enter a valid recipient email address."; return; }
        try {
            const subject = `FSM Dashboard Summary - ${new Date().toLocaleDateString()}`; const body = generateEmailBody(); const mailtoLink = `mailto:${recipient}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
             if (mailtoLink.length > 2000) { shareStatus.textContent = "Generated email body is too long for mailto link."; console.warn("Mailto link length exceeds 2000 characters:", mailtoLink.length); return; }
            window.location.href = mailtoLink; shareStatus.textContent = "Email client should open. Please review and send.";
        } catch (error) { console.error("Error generating mailto link:", error); shareStatus.textContent = "Error generating email content."; }
    };

    const selectAllOptions = (selectElement) => { if (!selectElement) return; Array.from(selectElement.options).forEach(option => option.selected = true); if (selectElement === territoryFilter) updateStoreFilterOptionsBasedOnHierarchy(); };
    const deselectAllOptions = (selectElement) => { if (!selectElement) return; selectElement.selectedIndex = -1; if (selectElement === territoryFilter) updateStoreFilterOptionsBasedOnHierarchy(); };

    // --- Event Listeners ---
    excelFileInput?.addEventListener('change', handleFile); applyFiltersButton?.addEventListener('click', applyFilters); storeSearch?.addEventListener('input', filterStoreOptions); exportCsvButton?.addEventListener('click', exportData); shareEmailButton?.addEventListener('click', handleShareEmail); closeStoreDetailsButton?.addEventListener('click', hideStoreDetails);
    territorySelectAll?.addEventListener('click', () => selectAllOptions(territoryFilter)); territoryDeselectAll?.addEventListener('click', () => deselectAllOptions(territoryFilter)); storeSelectAll?.addEventListener('click', () => selectAllOptions(storeFilter)); storeDeselectAll?.addEventListener('click', () => deselectAllOptions(storeFilter));
    attachRateTable?.querySelector('thead')?.addEventListener('click', handleSort); // Sorting only for Attach Rate table

    // --- Initial Setup ---
    resetUI();

}); // End DOMContentLoaded
