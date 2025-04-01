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
            val = safeGet(row, 'Retail Mode Connectivity', null); if (isValidAndNonZeroForAverage(
