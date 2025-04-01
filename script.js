/* Generated: 2025-03-31 11:56:44 PM EDT - Add "Top 5 / Bottom 5" tables section, conditionally displayed for single-territory filters. */
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
        const value = obj ? obj[path] : undefined;
        return (value !== undefined && value !== null) ? value : defaultValue;
    };
    // Helper to check if a value is valid for averaging (not null, not empty, parses to number)
    const isValidForAverage = (value) => {
         if (value === null || value === undefined || value === '') return false;
         // Check if it parses to a number (handles both numbers and numeric strings)
         return !isNaN(parseNumber(String(value).replace('%',''))); // Check parseNumber after removing % just in case
    };
    // Helper to check if a value is valid for averaging AND non-zero
    const isValidAndNonZeroForAverage = (value) => {
         if (!isValidForAverage(value)) return false; // Use previous check first
         const num = parseNumber(String(value).replace('%','')); // Parse again to ensure we have the number
         return num !== 0; // Return true only if it's not zero
    };
    // Function to calculate QTD Gap for sorting/display
    const calculateQtdGap = (row) => {
         const revenue = parseNumber(safeGet(row, 'Revenue w/DF', 0));
         const target = parseNumber(safeGet(row, 'QTD Revenue Target', 0));
         // If either is NaN, the gap is conceptually undefined for sorting, treat as worst case (negative infinity)
         if (isNaN(revenue) || isNaN(target)) {
             return -Infinity;
         }
         return revenue - target;
    };
    // Function to calculate Rev AR% for display in Top/Bottom tables
    const calculateRevARPercent = (row) => {
        const revenue = parseNumber(safeGet(row, 'Revenue w/DF', 0));
        const target = parseNumber(safeGet(row, 'QTD Revenue Target', 0));
        if (isNaN(revenue) || isNaN(target) || target === 0) {
            return NaN; // Undefined if revenue/target is missing or target is zero
        }
        return revenue / target;
    };
    // Function to calculate Unit Achievement % for display in Top/Bottom tables
    const calculateUnitAchievementPercent = (row) => {
        const units = parseNumber(safeGet(row, 'Unit w/ DF', 0));
        const target = parseNumber(safeGet(row, 'Unit Target', 0));
        if (isNaN(units) || isNaN(target) || target === 0) {
            return NaN; // Undefined if units/target is missing or target is zero
        }
        return units / target;
    };

    const getUniqueValues = (data, column) => {
        // Use safeGet with '' as default to handle potential missing values gracefully
        const values = new Set(data.map(item => safeGet(item, column, '')).filter(val => val !== ''));
        return ['ALL', ...Array.from(values).sort()];
    };
    const setOptions = (selectElement, options, disable = false) => { // Added disable flag
        selectElement.innerHTML = ''; // Clear existing options
        options.forEach(optionValue => {
            const option = document.createElement('option');
            option.value = optionValue;
            option.textContent = optionValue === 'ALL' ? `-- ${optionValue} --` : optionValue;
            option.title = optionValue; // Add tooltip
            selectElement.appendChild(option);
        });
         selectElement.disabled = disable; // Control disable state
    };
    const setMultiSelectOptions = (selectElement, options, disable = false) => { // Added disable flag
         selectElement.innerHTML = ''; // Clear existing options
         options.forEach(optionValue => {
             if (optionValue === 'ALL') return; // Skip 'ALL' for multi-select content
             const option = document.createElement('option');
             option.value = optionValue;
             option.textContent = optionValue;
             option.title = optionValue; // Add tooltip
             selectElement.appendChild(option);
         });
         selectElement.disabled = disable;
         // Keep existing selection if possible, otherwise deselect all
         // selectElement.selectedIndex = -1; // Start deselected
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
        if (!file) {
            statusDiv.textContent = 'No file selected.';
            return;
        }

        statusDiv.textContent = 'Reading file...';
        showLoading(true);
        filterArea.style.display = 'none';
        resultsArea.style.display = 'none';
        resetFilters(); // Reset filters visually

        try {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data);
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            // Use { defval: null } to preserve null/empty values instead of skipping rows/cols
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: null });

            // Validation
            if (jsonData.length > 0) {
                const headers = Object.keys(jsonData[0]);
                const missingHeaders = REQUIRED_HEADERS.filter(h => !headers.includes(h));
                if (missingHeaders.length > 0) {
                    // Log missing headers for debugging but maybe allow the app to continue
                    console.warn(`Warning: Missing expected columns: ${missingHeaders.join(', ')}. Some features might not work correctly.`);
                    // throw new Error(`Missing required columns: ${missingHeaders.join(', ')}`); // Make it non-fatal?
                }
            } else {
                throw new Error("Excel sheet appears to be empty.");
            }

            rawData = jsonData;
            // Extract unique store names, filter out null/empty strings specifically
            allPossibleStores = [...new Set(rawData.map(r => safeGet(r, 'Store', null)).filter(s => s))]
                                 .sort()
                                 .map(s => ({ value: s, text: s }));
            statusDiv.textContent = `Loaded ${rawData.length} rows. Adjust filters and click 'Apply Filters'.`;
            populateFilters(rawData); // Populate filters with all options
            filterArea.style.display = 'block'; // Show filters

        } catch (error) {
            console.error('Error processing file:', error);
            statusDiv.textContent = `Error: ${error.message}`;
            rawData = [];
            allPossibleStores = [];
            filteredData = [];
            resetUI();
        } finally {
            showLoading(false);
            excelFileInput.value = ''; // Reset file input
        }
    };

    const populateFilters = (data) => {
        // Populate top-level filters first
        setOptions(regionFilter, getUniqueValues(data, 'REGION'));
        setOptions(districtFilter, getUniqueValues(data, 'DISTRICT'));
        setMultiSelectOptions(territoryFilter, getUniqueValues(data, 'Q2 Territory').slice(1));
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

        applyFiltersButton.disabled = false;

        // Add event listeners to dependency filters *after* they are populated
        addDependencyFilterListeners();
    };

    const addDependencyFilterListeners = () => {
        // Define the handler function once
        const handler = updateStoreFilterOptionsBasedOnHierarchy;

        // Remove existing listeners first to avoid duplicates if file is reloaded
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

     // Updates store filter options based on other selections
     const updateStoreFilterOptionsBasedOnHierarchy = () => {
        if (rawData.length === 0) return; // No data loaded

        // 1. Get current selections from *dependency* filters
        const selectedRegion = regionFilter.value;
        const selectedDistrict = districtFilter.value;
        const selectedTerritories = Array.from(territoryFilter.selectedOptions).map(opt => opt.value);
        const selectedFsm = fsmFilter.value;
        const selectedChannel = channelFilter.value;
        const selectedSubchannel = subchannelFilter.value;
        const selectedDealer = dealerFilter.value;
        const selectedFlags = {};
        Object.entries(flagFiltersCheckboxes).forEach(([key, input]) => {
             if (input && input.checked) {
                 selectedFlags[key] = true;
             }
         });

        // 2. Filter rawData based *only* on these dependency filters
        const potentiallyValidStoresData = rawData.filter(row => {
            if (selectedRegion !== 'ALL' && safeGet(row, 'REGION', null) !== selectedRegion) return false;
            if (selectedDistrict !== 'ALL' && safeGet(row, 'DISTRICT', null) !== selectedDistrict) return false;
            if (selectedTerritories.length > 0 && !selectedTerritories.includes(safeGet(row, 'Q2 Territory', null))) return false;
            if (selectedFsm !== 'ALL' && safeGet(row, 'FSM NAME', null) !== selectedFsm) return false;
            if (selectedChannel !== 'ALL' && safeGet(row, 'CHANNEL', null) !== selectedChannel) return false;
            if (selectedSubchannel !== 'ALL' && safeGet(row, 'SUB_CHANNEL', null) !== selectedSubchannel) return false;
            if (selectedDealer !== 'ALL' && safeGet(row, 'DEALER_NAME', null) !== selectedDealer) return false;

            for (const flag in selectedFlags) {
                const flagValue = safeGet(row, flag, 'NO'); // Assume 'NO' or null means false
                if (!(flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1')) {
                   return false;
                }
            }
            return true;
        });

        // 3. Extract unique store names from the filtered data
        const validStoreNames = new Set(potentiallyValidStoresData.map(row => safeGet(row, 'Store', null)).filter(Boolean)); // Filter out null/empty
        storeOptions = Array.from(validStoreNames).sort().map(s => ({ value: s, text: s }));

        // 4. Get currently selected stores BEFORE repopulating
        const previouslySelectedStores = new Set(Array.from(storeFilter.selectedOptions).map(opt => opt.value));

        // 5. Update the store filter <select> element visually
        setStoreFilterOptions(storeOptions, false); // Update the dropdown

        // 6. Re-apply the search filter if there was text in the search box
        filterStoreOptions(); // This updates the visual list based on search term

        // 7. Attempt to re-select previously selected stores *if* they are still in the valid options
        // Iterate over the options *currently displayed* in the select element (after search filter)
        Array.from(storeFilter.options).forEach(option => {
            if (previouslySelectedStores.has(option.value)) {
                option.selected = true;
            }
            // Note: We don't explicitly deselect here, as filterStoreOptions already rebuilt the list
        });
        // If no options are selected after trying to restore, deselect all
        if (storeFilter.selectedOptions.length === 0) {
             storeFilter.selectedIndex = -1;
        }
    };

    // Modified to accept options and disable state
    const setStoreFilterOptions = (optionsToShow, disable = true) => {
        const currentSearchTerm = storeSearch.value; // Preserve search term
        storeFilter.innerHTML = ''; // Clear existing visual options
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
         storeSearch.value = currentSearchTerm; // Restore search term
    };

    // Modified to filter the *current* storeOptions array and update visual list
    const filterStoreOptions = () => {
        const searchTerm = storeSearch.value.toLowerCase();
        // Filter the dynamically updated *global* storeOptions array
        const filteredOptions = storeOptions.filter(opt => opt.text.toLowerCase().includes(searchTerm));

        // Get currently selected values before clearing visual list
        const selectedValues = new Set(Array.from(storeFilter.selectedOptions).map(opt => opt.value));

        // Update the visual <select> element with only the search-filtered options
        storeFilter.innerHTML = ''; // Clear existing visual options
        filteredOptions.forEach(opt => {
            const option = document.createElement('option');
            option.value = opt.value;
            option.textContent = opt.text;
            option.title = opt.text;
             // Re-select if it was selected before the search filter was applied/changed
             if (selectedValues.has(opt.value)) {
                 option.selected = true;
             }
            storeFilter.appendChild(option);
        });

        // Re-enable/disable select all based on *visible* options
        storeSelectAll.disabled = storeFilter.disabled || filteredOptions.length === 0;
        storeDeselectAll.disabled = storeFilter.disabled || filteredOptions.length === 0;
    };


    const applyFilters = () => {
        showLoading(true, true);
        resultsArea.style.display = 'none';

        setTimeout(() => {
            try {
                // 1. Get final filter values
                const selectedRegion = regionFilter.value;
                const selectedDistrict = districtFilter.value;
                const selectedTerritories = Array.from(territoryFilter.selectedOptions).map(opt => opt.value);
                const selectedFsm = fsmFilter.value;
                const selectedChannel = channelFilter.value;
                const selectedSubchannel = subchannelFilter.value;
                const selectedDealer = dealerFilter.value;
                const selectedStores = Array.from(storeFilter.selectedOptions).map(opt => opt.value); // Final store selection
                const selectedFlags = {};
                 Object.entries(flagFiltersCheckboxes).forEach(([key, input]) => {
                     if (input && input.checked) {
                         selectedFlags[key] = true;
                     }
                 });

                // 2. Filter rawData based on *all* filters simultaneously
                filteredData = rawData.filter(row => {
                    if (selectedRegion !== 'ALL' && safeGet(row, 'REGION', null) !== selectedRegion) return false;
                    if (selectedDistrict !== 'ALL' && safeGet(row, 'DISTRICT', null) !== selectedDistrict) return false;
                    if (selectedTerritories.length > 0 && !selectedTerritories.includes(safeGet(row, 'Q2 Territory', null))) return false;
                    if (selectedFsm !== 'ALL' && safeGet(row, 'FSM NAME', null) !== selectedFsm) return false;
                    if (selectedChannel !== 'ALL' && safeGet(row, 'CHANNEL', null) !== selectedChannel) return false;
                    if (selectedSubchannel !== 'ALL' && safeGet(row, 'SUB_CHANNEL', null) !== selectedSubchannel) return false;
                    if (selectedDealer !== 'ALL' && safeGet(row, 'DEALER_NAME', null) !== selectedDealer) return false;

                    // Apply store filter - this list might have been pre-filtered by hierarchy changes
                    if (selectedStores.length > 0 && !selectedStores.includes(safeGet(row, 'Store', null))) return false;

                    for (const flag in selectedFlags) {
                        const flagValue = safeGet(row, flag, 'NO');
                        if (!(flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1')) {
                           return false;
                        }
                    }
                    return true;
                });

                // 3. Update UI elements
                updateSummary(filteredData);
                updateTopBottomTables(filteredData); // ** NEW ** Update Top/Bottom tables
                updateCharts(filteredData);
                updateAttachRateTable(filteredData);

                // 4. Auto Show/Hide Store Details based on Filter Count (Keep this logic)
                if (filteredData.length === 1) {
                    showStoreDetails(filteredData[0]);
                    highlightTableRow(safeGet(filteredData[0], 'Store', null)); // Ensure single row is highlighted
                } else {
                    hideStoreDetails(); // Hides details if >1 or 0 stores
                }

                // 5. Finalize UI state
                statusDiv.textContent = `Displaying ${filteredData.length} of ${rawData.length} rows based on filters.`; // Update status text
                resultsArea.style.display = 'block';
                exportCsvButton.disabled = filteredData.length === 0;

            } catch (error) {
                console.error("Error applying filters:", error);
                statusDiv.textContent = "Error applying filters. Check console for details."; // Update status text
                filteredData = [];
                resultsArea.style.display = 'none';
                 exportCsvButton.disabled = true;
                 updateSummary([]);
                 updateTopBottomTables([]); // ** NEW ** Clear Top/Bottom tables on error
                 updateCharts([]);
                 updateAttachRateTable([]);
                 hideStoreDetails();

            } finally {
                 showLoading(false, true);
            }
        }, 10); // Small delay for rendering loading indicator
    };

    const resetFilters = () => {
        // Reset dropdowns and multi-selects
         [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(sel => { if (sel) { sel.value = 'ALL'; sel.disabled = true;} });
         if (territoryFilter) { territoryFilter.selectedIndex = -1; territoryFilter.disabled = true; }
         if (storeFilter) {
             storeFilter.innerHTML = '<option value="ALL">-- Load File First --</option>'; // Clear options
             storeFilter.selectedIndex = -1;
             storeFilter.disabled = true;
         }
         if (storeSearch) { storeSearch.value = ''; storeSearch.disabled = true; }
         storeOptions = []; // Clear dynamic store options
         allPossibleStores = []; // Clear all possible stores too


         // Reset checkboxes
         Object.values(flagFiltersCheckboxes).forEach(input => { if(input) {input.checked = false; input.disabled = true;} });

         // Disable buttons
         if (applyFiltersButton) applyFiltersButton.disabled = true;
         if (territorySelectAll) territorySelectAll.disabled = true;
         if (territoryDeselectAll) territoryDeselectAll.disabled = true;
         if (storeSelectAll) storeSelectAll.disabled = true;
         if (storeDeselectAll) storeDeselectAll.disabled = true;
         if (exportCsvButton) exportCsvButton.disabled = true;

         // Remove dependency listeners if they exist
         const handler = updateStoreFilterOptionsBasedOnHierarchy;
         [regionFilter, districtFilter, territoryFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(filter => {
             if (filter) filter.removeEventListener('change', handler);
         });
          Object.values(flagFiltersCheckboxes).forEach(input => {
             if (input) input.removeEventListener('change', handler);
          });
    };

     const resetUI = () => {
         resetFilters(); // Resets and disables filters
         if (filterArea) filterArea.style.display = 'none';
         if (resultsArea) resultsArea.style.display = 'none';
         // Clear potential leftover data displays
         if (mainChartInstance) { mainChartInstance.destroy(); mainChartInstance = null; }
         if (attachRateTableBody) attachRateTableBody.innerHTML = '';
         if (attachRateTableFooter) attachRateTableFooter.innerHTML = '';
         if (attachTableStatus) attachTableStatus.textContent = '';
         if (topBottomSection) topBottomSection.style.display = 'none'; // ** NEW ** Hide Top/Bottom section
         if (top5TableBody) top5TableBody.innerHTML = ''; // ** NEW ** Clear Top 5 table
         if (bottom5TableBody) bottom5TableBody.innerHTML = ''; // ** NEW ** Clear Bottom 5 table
         hideStoreDetails();
         // Reset summary fields
         updateSummary([]);
         if(statusDiv) statusDiv.textContent = 'No file selected.'; // Reset status message
     };

    // --- UPDATED updateSummary with new calculation logic ---
    const updateSummary = (data) => {
        const totalCount = data.length; // Total rows matching filters

        // Clear summary fields first
        const fieldsToClear = [revenueWithDFValue, qtdRevenueTargetValue, qtdGapValue, quarterlyRevenueTargetValue,
                        percentQuarterlyStoreTargetValue, revARValue, unitsWithDFValue, unitTargetValue,
                        unitAchievementValue, visitCountValue, trainingCountValue, retailModeConnectivityValue,
                        repSkillAchValue, vPmrAchValue, postTrainingScoreValue, eliteValue,
                        percentQuarterlyTerritoryTargetValue, territoryRevPercentValue, districtRevPercentValue, regionRevPercentValue];
        fieldsToClear.forEach(el => { if (el) el.textContent = 'N/A'; });
        [percentQuarterlyTerritoryTargetP, territoryRevPercentP, districtRevPercentP, regionRevPercentP].forEach(p => { if(p) p.style.display = 'none';});

        if (totalCount === 0) {
            return; // Nothing more to do if no data
        }

        // --- Calculate SUMS (Used for several calculations) ---
        const sumRevenue = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Revenue w/DF', 0)), 0);
        const sumQtdTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'QTD Revenue Target', 0)), 0);
        const sumQuarterlyTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Quarterly Revenue Target', 0)), 0);
        const sumUnits = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Unit w/ DF', 0)), 0);
        const sumUnitTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Unit Target', 0)), 0);
        const sumVisits = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Visit count', 0)), 0);
        const sumTrainings = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Trainings', 0)), 0);

        // --- Calculate AVERAGES (excluding blanks/invalid values AND ZEROS where specified) ---
        // Use the new isValidAndNonZeroForAverage helper for metrics where 0 should be ignored
        let sumConnectivity = 0, countConnectivity = 0;
        let sumRepSkill = 0, countRepSkill = 0;
        let sumPmr = 0, countPmr = 0;
        // Use original isValidForAverage for metrics where 0 is a valid value to include
        let sumPostTraining = 0, countPostTraining = 0;
        let sumElite = 0, countElite = 0;


        data.forEach(row => {
            let val;
            // Retail Mode Connectivity (Ignore Blanks & Zeros)
            val = safeGet(row, 'Retail Mode Connectivity', null);
            if (isValidAndNonZeroForAverage(val)) { sumConnectivity += parsePercent(val); countConnectivity++; }
            // Rep Skill Ach (Ignore Blanks & Zeros)
            val = safeGet(row, 'Rep Skill Ach', null);
            if (isValidAndNonZeroForAverage(val)) { sumRepSkill += parsePercent(val); countRepSkill++; }
            // (V)PMR Ach (Ignore Blanks & Zeros)
            val = safeGet(row, '(V)PMR Ach', null);
            if (isValidAndNonZeroForAverage(val)) { sumPmr += parsePercent(val); countPmr++; }
            // Post Training Score (Ignore Blanks, Include Zeros)
            val = safeGet(row, 'Post Training Score', null);
            if (isValidForAverage(val)) { sumPostTraining += parseNumber(val); countPostTraining++; }
            // Elite (Ignore Blanks, Include Zeros - assuming 0% is valid)
            val = safeGet(row, 'Elite', null);
            if (isValidForAverage(val)) { sumElite += parsePercent(val); countElite++; }
        });

        const avgConnectivity = countConnectivity === 0 ? NaN : sumConnectivity / countConnectivity;
        const avgRepSkill = countRepSkill === 0 ? NaN : sumRepSkill / countRepSkill;
        const avgPmr = countPmr === 0 ? NaN : sumPmr / countPmr;
        const avgPostTraining = countPostTraining === 0 ? NaN : sumPostTraining / countPostTraining;
        const avgElite = countElite === 0 ? NaN : sumElite / countElite;

        // --- Calculate Overall Percentages (based on SUMS) ---
        // ** NEW CALCULATION **: Rev AR% = Sum Revenue / Sum QTD Target
        const calculatedRevAR = sumQtdTarget === 0 ? 0 : sumRevenue / sumQtdTarget;
        // ** VERIFIED **: Unit Achievement % = Sum Units / Sum Unit Target
        const overallUnitAchievement = sumUnitTarget === 0 ? 0 : sumUnits / sumUnitTarget;
        // % Store Quarterly Target = Sum Revenue / Sum Quarterly Target
        const overallPercentStoreTarget = sumQuarterlyTarget === 0 ? 0 : sumRevenue / sumQuarterlyTarget;


        // --- Update DOM Elements (safely using optional chaining) ---
        // Sums
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

        // Overall % based on Sums
        revARValue && (revARValue.textContent = formatPercent(calculatedRevAR)); // NEW CALCULATION DISPLAYED
        revARValue && (revARValue.title = `Calculated Rev AR% (Total Revenue / Total QTD Target)`); // NEW TITLE
        unitAchievementValue && (unitAchievementValue.textContent = formatPercent(overallUnitAchievement)); // Correct Calculation
        unitAchievementValue && (unitAchievementValue.title = `Overall Unit Achievement % (Total Units / Total Unit Target)`); // Correct Title
        percentQuarterlyStoreTargetValue && (percentQuarterlyStoreTargetValue.textContent = formatPercent(overallPercentStoreTarget));
        percentQuarterlyStoreTargetValue && (percentQuarterlyStoreTargetValue.title = `Overall % Quarterly Target (Total Revenue / Total Quarterly Target)`);


        // Averages (showing count of valid entries in title, updated for new zero-exclusion logic)
        retailModeConnectivityValue && (retailModeConnectivityValue.textContent = formatPercent(avgConnectivity));
        retailModeConnectivityValue && (retailModeConnectivityValue.title = `Average 'Retail Mode Connectivity' across ${countConnectivity} stores with non-zero data`); // Updated Title
        repSkillAchValue && (repSkillAchValue.textContent = formatPercent(avgRepSkill));
        repSkillAchValue && (repSkillAchValue.title = `Average 'Rep Skill Ach' across ${countRepSkill} stores with non-zero data`); // Updated Title
        vPmrAchValue && (vPmrAchValue.textContent = formatPercent(avgPmr));
        vPmrAchValue && (vPmrAchValue.title = `Average '(V)PMR Ach' across ${countPmr} stores with non-zero data`); // Updated Title
        postTrainingScoreValue && (postTrainingScoreValue.textContent = formatNumber(avgPostTraining.toFixed(1))); // Keep formatting
        postTrainingScoreValue && (postTrainingScoreValue.title = `Average 'Post Training Score' across ${countPostTraining} stores with data`); // Title assumes 0 is valid
        eliteValue && (eliteValue.textContent = formatPercent(avgElite));
        eliteValue && (eliteValue.title = `Average 'Elite' score % across ${countElite} stores with data`); // Title assumes 0 is valid

        // Contextual Hierarchy Percentages (Recalculate averages excluding blanks here too)
        updateContextualSummary(data);
    };

    // --- UPDATED updateContextualSummary to exclude blanks from averages ---
    const updateContextualSummary = (data) => {
        // Hide all contextual fields initially safely
        [percentQuarterlyTerritoryTargetP, territoryRevPercentP, districtRevPercentP, regionRevPercentP].forEach(p => {if (p) p.style.display = 'none'});

        if (data.length === 0) return;

        const singleRegion = regionFilter.value !== 'ALL';
        const singleDistrict = districtFilter.value !== 'ALL';
        const singleTerritory = territoryFilter.selectedOptions.length === 1;

        // Helper function to calculate average excluding blanks for a specific column
        const calculateAverageExcludeBlanks = (column) => {
            let sum = 0;
            let count = 0;
            data.forEach(row => {
                const val = safeGet(row, column, null);
                if (isValidForAverage(val)) { // Keep original check here unless zeros need excluding too
                    sum += parsePercent(val); // Assuming these are percentages
                    count++;
                }
            });
            return count === 0 ? NaN : sum / count;
        };

        // Calculate averages excluding blanks
        const avgPercentTerritoryTarget = calculateAverageExcludeBlanks('%Quarterly Territory Rev Target');
        const avgTerritoryRevPercent = calculateAverageExcludeBlanks('Territory Rev%');
        const avgDistrictRevPercent = calculateAverageExcludeBlanks('District Rev%');
        const avgRegionRevPercent = calculateAverageExcludeBlanks('Region Rev%');

        // Update DOM
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

     // ** NEW ** Function to update Top 5 / Bottom 5 tables
     const updateTopBottomTables = (data) => {
        if (!topBottomSection || !top5TableBody || !bottom5TableBody) return; // Ensure elements exist

        // Clear previous content
        top5TableBody.innerHTML = '';
        bottom5TableBody.innerHTML = '';

        // Determine if data belongs to a single territory
        const territoriesInData = new Set(data.map(row => safeGet(row, 'Q2 Territory', null)).filter(Boolean));
        const isSingleTerritory = territoriesInData.size === 1;

        if (!isSingleTerritory || data.length === 0) {
            topBottomSection.style.display = 'none'; // Hide the section
            return;
        }

        topBottomSection.style.display = 'flex'; // Show the section (use flex for side-by-side)

        // --- Top 5 Table (Revenue) ---
        const top5Data = [...data]
            .sort((a, b) => {
                // Sort descending by Revenue w/DF. Handle NaN by placing them lower.
                const revA = parseNumber(safeGet(a, 'Revenue w/DF', -Infinity));
                const revB = parseNumber(safeGet(b, 'Revenue w/DF', -Infinity));
                return revB - revA;
            })
            .slice(0, TOP_N_TABLES);

        top5Data.forEach(row => {
            const tr = top5TableBody.insertRow();
            const storeName = safeGet(row, 'Store', null);
            tr.dataset.storeName = storeName; // For highlighting and details
            tr.onclick = () => {
                showStoreDetails(row);
                highlightTableRow(storeName);
            };

            const revenue = parseNumber(safeGet(row, 'Revenue w/DF', NaN));
            const revAR = calculateRevARPercent(row); // Use helper
            const unitAch = calculateUnitAchievementPercent(row); // Use helper
            const visits = parseNumber(safeGet(row, 'Visit count', NaN));

            tr.insertCell().textContent = storeName;
            tr.insertCell().textContent = formatCurrency(revenue);
            tr.insertCell().textContent = formatPercent(revAR);
            tr.insertCell().textContent = formatPercent(unitAch);
            tr.insertCell().textContent = formatNumber(visits);
        });

        // --- Bottom 5 Table (QTD Gap - Opportunities) ---
        const bottom5Data = [...data]
            .sort((a, b) => {
                // Sort descending by QTD Gap (Revenue - Target). Biggest Gap first.
                // Handle undefined gaps (due to NaN revenue/target) by placing them lower.
                const gapA = calculateQtdGap(a);
                const gapB = calculateQtdGap(b);
                return gapB - gapA; // Higher gap (worse performance) comes first
            })
            .slice(0, TOP_N_TABLES);

         bottom5Data.forEach(row => {
             const tr = bottom5TableBody.insertRow();
             const storeName = safeGet(row, 'Store', null);
             tr.dataset.storeName = storeName;
             tr.onclick = () => {
                 showStoreDetails(row);
                 highlightTableRow(storeName);
             };

             const qtdGap = calculateQtdGap(row); // Use helper
             const revAR = calculateRevARPercent(row); // Use helper
             const unitAch = calculateUnitAchievementPercent(row); // Use helper
             const visits = parseNumber(safeGet(row, 'Visit count', NaN));

             tr.insertCell().textContent = storeName;
             tr.insertCell().textContent = formatCurrency(qtdGap === -Infinity ? NaN : qtdGap); // Format gap, handle -Infinity case
             tr.insertCell().textContent = formatPercent(revAR);
             tr.insertCell().textContent = formatPercent(unitAch);
             tr.insertCell().textContent = formatNumber(visits);
         });
     };


    const updateCharts = (data) => {
        if (mainChartInstance) {
            mainChartInstance.destroy();
            mainChartInstance = null;
        }

        if (data.length === 0 || !mainChartCanvas) return;

        // --- Main Chart: Revenue Performance Bar Chart ---
        const sortedData = [...data].sort((a, b) => parseNumber(safeGet(b, 'Revenue w/DF', 0)) - parseNumber(safeGet(a, 'Revenue w/DF', 0)));
        const chartData = sortedData.slice(0, TOP_N_CHART);

        const labels = chartData.map(row => safeGet(row, 'Store', 'Unknown')); // Provide default label
        const revenueDataSet = chartData.map(row => parseNumber(safeGet(row, 'Revenue w/DF', 0)));
        const targetDataSet = chartData.map(row => parseNumber(safeGet(row, 'QTD Revenue Target', 0)));

        const backgroundColors = chartData.map((row, index) => {
             const revenue = revenueDataSet[index];
             const target = targetDataSet[index];
             return revenue >= target ? 'rgba(75, 192, 192, 0.6)' : 'rgba(255, 99, 132, 0.6)';
        });
        const borderColors = chartData.map((row, index) => {
             const revenue = revenueDataSet[index];
             const target = targetDataSet[index];
             return revenue >= target ? 'rgba(75, 192, 192, 1)' : 'rgba(255, 99, 132, 1)';
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
                        type: 'line', // Show target as a line overlay
                        borderColor: 'rgba(255, 206, 86, 1)', // Yellow line
                        backgroundColor: 'rgba(255, 206, 86, 0.2)',
                        fill: false,
                        tension: 0.1,
                        borderWidth: 2,
                        pointRadius: 3,
                        pointHoverRadius: 5
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: { color: '#e0e0e0', callback: (value) => formatCurrency(value) },
                        grid: { color: 'rgba(224, 224, 224, 0.2)' }
                    },
                    x: {
                        ticks: { color: '#e0e0e0' },
                        grid: { display: false }
                    }
                },
                plugins: {
                    legend: { labels: { color: '#e0e0e0' } },
                    tooltip: {
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
                                          if (chartData && chartData[context.dataIndex]){
                                               const storeData = chartData[context.dataIndex];
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
                        const storeName = labels[index];
                        // Find the original data row from filteredData, not just chartData which might be sliced/sorted differently
                        const storeData = filteredData.find(row => safeGet(row, 'Store', null) === storeName);
                        if (storeData) {
                            showStoreDetails(storeData);
                             highlightTableRow(storeName);
                        }
                    }
                }
            }
        });
    };

    // --- UPDATED updateAttachRateTable to exclude blanks from footer average ---
    // --- Removed % Quarterly Revenue Target Column ---
    const updateAttachRateTable = (data) => {
        if (!attachRateTableBody || !attachRateTableFooter) return; // Exit if elements don't exist
        attachRateTableBody.innerHTML = '';
        attachRateTableFooter.innerHTML = '';

        if (data.length === 0) {
            if(attachTableStatus) attachTableStatus.textContent = 'No data to display based on filters.';
            return;
        }

        // Sort data based on currentSort state for this table
        const sortedData = [...data].sort((a, b) => {
             let valA = safeGet(a, currentSort.column, null); // Use null default for comparison consistency
             let valB = safeGet(b, currentSort.column, null);

             // Handle nulls first: consistently place them at the beginning or end
             if (valA === null && valB === null) return 0;
             if (valA === null) return currentSort.ascending ? -1 : 1; // Nulls first when ascending
             if (valB === null) return currentSort.ascending ? 1 : -1; // Nulls first when ascending

             // Attempt numeric conversion for sorting if applicable
             const isPercentCol = currentSort.column.includes('Attach Rate'); // Only attach rates here
             const numA = isPercentCol ? parsePercent(valA) : parseNumber(valA);
             const numB = isPercentCol ? parsePercent(valB) : parseNumber(valB);

             if (!isNaN(numA) && !isNaN(numB)) {
                 // Numeric comparison
                 return currentSort.ascending ? numA - numB : numB - numA;
             } else {
                 // String comparison (case-insensitive) for Store column primarily
                 valA = String(valA).toLowerCase();
                 valB = String(valB).toLowerCase();
                 return currentSort.ascending ? valA.localeCompare(valB) : valB.localeCompare(valA);
             }
         });


        // Calculate averages for highlighting AND footer (excluding blanks)
        // Removed '% Quarterly Revenue Target' from metrics to average
        const averageMetrics = [
             'Tablet Attach Rate', 'PC Attach Rate', 'NC Attach Rate',
             'TWS Attach Rate', 'WW Attach Rate', 'ME Attach Rate', 'NCME Attach Rate'
        ];
        const averages = {};
        averageMetrics.forEach(key => {
             let sum = 0;
             let count = 0;
             data.forEach(row => {
                 const val = safeGet(row, key, null);
                 if (isValidForAverage(val)) { // Keep original check for attach rates (0% is valid)
                     sum += parsePercent(val);
                     count++;
                 }
             });
             averages[key] = count === 0 ? NaN : sum / count;
         });


        sortedData.forEach(row => {
            const tr = document.createElement('tr');
            const storeName = safeGet(row, 'Store', null); // Get store name safely
            if (storeName) { // Only add rows with a valid store name
                 tr.dataset.storeName = storeName;
                 tr.onclick = () => {
                     showStoreDetails(row);
                     highlightTableRow(storeName);
                 };

                 // Define columns and their formatting/highlighting logic
                 // Removed '% Quarterly Revenue Target' column definition
                 const columns = [
                     { key: 'Store', format: (val) => val }, // Keep store name as is
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
                     const rawValue = safeGet(row, col.key, null); // Get raw value safely
                     // Determine if parsing as percent is needed
                     // Adjusted isPercentCol check slightly as % Target is gone
                     const isPercentCol = col.key.includes('Attach Rate'); // Only attach rates
                     // Use parsePercent for percent-like columns, parseNumber otherwise, to get numeric value for comparison
                     const numericValue = isPercentCol ? parsePercent(rawValue) : parseNumber(rawValue);

                     let formattedValue;
                     // Display N/A if raw value is null or resulted in NaN (except Store column)
                     if (rawValue === null || (col.key !== 'Store' && isNaN(numericValue))) {
                         formattedValue = 'N/A';
                     } else {
                         // Format the numeric value if percent, otherwise use raw (which might be string like Store name)
                         formattedValue = isPercentCol ? col.format(numericValue) : rawValue;
                     }


                     td.textContent = formattedValue;
                     td.title = `${col.key}: ${formattedValue}`; // Add tooltip

                     // Apply highlighting based on average (only if value is numeric)
                     if (col.highlight && averages[col.key] !== undefined && !isNaN(numericValue)) {
                          if (numericValue >= averages[col.key]) {
                              td.classList.add('highlight-green');
                          } else {
                              td.classList.add('highlight-red');
                          }
                      }
                     tr.appendChild(td);
                 });

                 attachRateTableBody.appendChild(tr);
             } // End if(storeName)
        });

        // Add Average Row to Footer using calculated averages (which exclude blanks)
        if (data.length > 0) { // Check if there was any data to calculate averages from
            const footerRow = attachRateTableFooter.insertRow();
            const avgLabelCell = footerRow.insertCell();
            avgLabelCell.textContent = 'Filtered Avg*'; // Add asterisk note
            avgLabelCell.title = 'Average calculated only using stores with valid data for each column';
            avgLabelCell.style.textAlign = "right";
            avgLabelCell.style.fontWeight = "bold";

            // Use the same metrics keys used for calculation (which no longer includes % Qtr Rev Target)
            averageMetrics.forEach(key => {
                 const td = footerRow.insertCell();
                 const avgValue = averages[key]; // Already calculated excluding blanks
                 td.textContent = formatPercent(avgValue); // Format the calculated average
                 // Count valid entries for tooltip
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
         if (!sortKey) return; // No sort key defined

         if (currentSort.column === sortKey) {
             currentSort.ascending = !currentSort.ascending;
         } else {
             currentSort.column = sortKey;
             currentSort.ascending = true; // Default to ascending for new column
         }
         // IMPORTANT: Update the attach rate table using the currently *globally filtered* data
         updateAttachRateTable(filteredData);
    };


    const updateSortArrows = () => {
        if (!attachRateTable) return;
        attachRateTable.querySelectorAll('th.sortable .sort-arrow').forEach(arrow => {
            arrow.classList.remove('asc', 'desc');
            arrow.textContent = ''; // Clear arrow text
        });
        // Safely attempt to query the header using CSS.escape for potentially complex names
        try {
            const currentHeader = attachRateTable.querySelector(`th[data-sort="${CSS.escape(currentSort.column)}"] .sort-arrow`);
            if (currentHeader) {
                currentHeader.classList.add(currentSort.ascending ? 'asc' : 'desc');
                // Add arrow character for visual cue
                currentHeader.textContent = currentSort.ascending ? ' ▲' : ' ▼'; // Add space before arrow
            }
        } catch (e) {
            console.error("Error updating sort arrows for column:", currentSort.column, e);
        }
    };

    // Updated showStoreDetails to include more fields and Map Link
    const showStoreDetails = (storeData) => {
        if (!storeDetailsContent || !storeDetailsSection || !closeStoreDetailsButton) return;

        // --- Generate Address String ---
        const addressParts = [
            safeGet(storeData, 'ADDRESS1', null), // Use null default
            safeGet(storeData, 'CITY', null),
            safeGet(storeData, 'STATE', null),
            safeGet(storeData, 'ZIPCODE', null)
        ].filter(Boolean); // Filter out null/empty parts
        const addressString = addressParts.length > 0 ? addressParts.join(', ') : null;

        // --- Generate Google Maps Link ---
        const latitude = parseFloat(safeGet(storeData, 'LATITUDE_ORG', NaN)); // Use NaN default
        const longitude = parseFloat(safeGet(storeData, 'LONGITUDE_ORG', NaN));
        let mapsLinkHtml = '';
        if (!isNaN(latitude) && !isNaN(longitude)) {
            const mapsUrl = `https://www.google.com/maps?q=${latitude},${longitude}`; // Corrected Maps URL
            mapsLinkHtml = `<p><a href="${mapsUrl}" target="_blank" title="Open in Google Maps">View on Google Maps</a></p>`;
        } else {
            mapsLinkHtml = `<p style="color: #aaa; font-style: italic;">(Map coordinates not available)</p>`;
        }

        // --- Generate Flag Summary (using FLAG_HEADERS defined at top) ---
        let flagSummaryHtml = FLAG_HEADERS.map(flag => {
            const flagValue = safeGet(storeData, flag); // Get raw value
            const isTrue = (flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1');
            return `<span title="${flag}" data-flag="${isTrue}">${flag.replace(/_/g, ' ')} ${isTrue ? '✔' : '✘'}</span>`;
        }).join(' | '); // Join with separator


        // --- Construct Details HTML ---
        let detailsHtml = `<p><strong>Store:</strong> ${safeGet(storeData, 'Store')}</p>`;
        if (addressString) {
             detailsHtml += `<p><strong>Address:</strong> ${addressString}</p>`;
        }
        detailsHtml += mapsLinkHtml; // Add the generated map link paragraph
        detailsHtml += `<hr>`;
        // Combine IDs, provide 'N/A' if specific ID is missing
        detailsHtml += `<p><strong>IDs:</strong> Store: ${safeGet(storeData, 'STORE ID')} | Org: ${safeGet(storeData, 'ORG_STORE_ID')} | CV: ${safeGet(storeData, 'CV_STORE_ID')} | CinglePoint: ${safeGet(storeData, 'CINGLEPOINT_ID')}</p>`;
        // Combine Type/Tier info
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
        detailsHtml += `<p><strong>Flags:</strong> ${flagSummaryHtml}</p>`; // Add the generated flag summary

        // --- Update DOM ---
        storeDetailsContent.innerHTML = detailsHtml;
        storeDetailsSection.style.display = 'block';
        closeStoreDetailsButton.style.display = 'inline-block';
        storeDetailsSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); // Scroll to details when shown
    };

    const hideStoreDetails = () => {
        if (!storeDetailsContent || !storeDetailsSection || !closeStoreDetailsButton) return;
        storeDetailsContent.innerHTML = 'Select a store from the table or chart for details, or apply filters resulting in a single store.';
        storeDetailsSection.style.display = 'none';
        closeStoreDetailsButton.style.display = 'none';
        highlightTableRow(null); // Remove table row highlight
    };

     // ** UPDATED ** to handle highlighting across multiple tables
     const highlightTableRow = (storeName) => {
         // Remove highlight from any previously selected row across all tables
         if (selectedStoreRow) {
             selectedStoreRow.classList.remove('selected-row');
             selectedStoreRow = null; // Reset tracker
         }
         // If a storeName is provided, find and highlight the corresponding row in *any* table
         if (storeName) {
             const tables = [attachRateTableBody, top5TableBody, bottom5TableBody];
             for (const tableBody of tables) {
                 if (tableBody) {
                     try {
                         // Use querySelector for safety with special characters
                         const rowToHighlight = tableBody.querySelector(`tr[data-store-name="${CSS.escape(storeName)}"]`);
                         if (rowToHighlight) {
                             rowToHighlight.classList.add('selected-row');
                             selectedStoreRow = rowToHighlight; // Track the newly selected row
                             break; // Stop searching once found
                         }
                     } catch (e) {
                         console.error("Error selecting table row:", e);
                     }
                 }
             }
         }
    };


    const exportData = () => {
        if (filteredData.length === 0) {
            alert("No filtered data to export.");
            return;
        }
        try {
            if (!attachRateTable) throw new Error("Attach rate table not found.");
            // Get headers directly from the updated HTML table header structure
            const headers = Array.from(attachRateTable.querySelectorAll('thead th'))
                                  .map(th => th.dataset.sort || th.textContent.replace(/ [▲▼]$/, '').trim()); // Get clean headers

             const dataToExport = filteredData.map(row => {
                return headers.map(header => {
                    // Need to handle the case where the header name in HTML might differ from the data key
                    // Example: 'Tablet' header maps to 'Tablet Attach Rate' key
                    let dataKey = header; // Assume header matches key initially
                    if (header === 'Tablet') dataKey = 'Tablet Attach Rate';
                    else if (header === 'PC') dataKey = 'PC Attach Rate';
                    else if (header === 'NC') dataKey = 'NC Attach Rate';
                    else if (header === 'TWS') dataKey = 'TWS Attach Rate';
                    else if (header === 'WW') dataKey = 'WW Attach Rate';
                    else if (header === 'ME') dataKey = 'ME Attach Rate';
                    else if (header === 'NCME') dataKey = 'NCME Attach Rate';
                    // Add other mappings if header text differs from data key

                    let value = safeGet(row, dataKey, ''); // Get raw value using the potentially mapped key

                    // Check if the *data key* indicates a percentage or rate
                    const isPercentLike = dataKey.includes('%') || dataKey.includes('Rate') || dataKey.includes('Ach') || dataKey.includes('Connectivity') || dataKey.includes('Elite');

                    if (isPercentLike) {
                        // Parse as percent, export as decimal number, handle non-numeric gracefully
                         const numVal = parsePercent(value);
                         value = isNaN(numVal) ? '' : numVal; // Export empty if not a valid percentage
                     } else {
                         // Handle potential commas and quotes in string values
                         if (typeof value === 'string' && (value.includes(',') || value.includes('"') || value.includes('\n'))) {
                             value = `"${value.replace(/"/g, '""')}"`; // Quote fields with commas, quotes, or newlines; escape internal quotes
                         }
                         // Preserve numbers as numbers, leave other types as strings
                         const numVal = parseNumber(value);
                          // Check if it's NOT percent-like AND is a valid number AND not boolean
                         if (!isPercentLike && !isNaN(numVal) && typeof value !== 'boolean') {
                             value = numVal;
                         }
                     }
                    return value;
                });
             });

            let csvContent = "data:text/csv;charset=utf-8,"
                + headers.join(",") + "\n"
                + dataToExport.map(e => e.join(",")).join("\n");

            const encodedUri = encodeURI(csvContent);
            const link = document.createElement("a");
            link.setAttribute("href", encodedUri);
            link.setAttribute("download", "fsm_dashboard_export.csv");
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        } catch (error) {
            console.error("Error exporting CSV:", error);
            alert("Error generating CSV export. See console for details.");
        }
    };

    const generateEmailBody = () => {
        if (filteredData.length === 0) {
            return "No data available based on current filters.";
        }
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
        body += `- Rev AR%: ${revARValue?.textContent || 'N/A'}\n`; // Uses updated calculation
        body += `- Total Units (incl. DF): ${unitsWithDFValue?.textContent || 'N/A'}\n`;
        body += `- Unit Achievement %: ${unitAchievementValue?.textContent || 'N/A'}\n`;
        body += `- Total Visits: ${visitCountValue?.textContent || 'N/A'}\n`;
        body += `- Avg. Connectivity: ${retailModeConnectivityValue?.textContent || 'N/A'}\n\n`; // Uses updated calculation
        body += "Mysteryshop & Training (Avg*):\n"; // Added asterisk note
        body += `- Rep Skill Ach: ${repSkillAchValue?.textContent || 'N/A'}\n`; // Uses updated calculation
        body += `- (V)PMR Ach: ${vPmrAchValue?.textContent || 'N/A'}\n`; // Uses updated calculation
        body += `- Post Training Score: ${postTrainingScoreValue?.textContent || 'N/A'}\n`;
        body += `- Elite Score %: ${eliteValue?.textContent || 'N/A'}\n\n`;
        body += "*Averages calculated only using stores with valid, non-zero data for each metric where applicable (Connectivity, Rep Skill, PMR).\n\n"; // Updated note


        // Add Top/Bottom 5 if applicable
        const territoriesInData = new Set(filteredData.map(row => safeGet(row, 'Q2 Territory', null)).filter(Boolean));
        if (territoriesInData.size === 1) {
             const territoryName = territoriesInData.values().next().value; // Get the single territory name
             body += `--- Territory: ${territoryName} ---\n`;

             // Top 5
             const top5Data = [...filteredData]
                 .sort((a, b) => parseNumber(safeGet(b, 'Revenue w/DF', -Infinity)) - parseNumber(safeGet(a, 'Revenue w/DF', -Infinity)))
                 .slice(0, TOP_N_TABLES);
             if (top5Data.length > 0) {
                 body += "Top 5 (Revenue):\n";
                 top5Data.forEach((row, i) => {
                     body += `${i+1}. ${safeGet(row, 'Store')} - Rev: ${formatCurrency(parseNumber(safeGet(row, 'Revenue w/DF')))}\n`;
                 });
                 body += "\n";
             }

             // Bottom 5
             const bottom5Data = [...filteredData]
                  .sort((a, b) => calculateQtdGap(b) - calculateQtdGap(a)) // Sort descending by gap
                 .slice(0, TOP_N_TABLES);
              if (bottom5Data.length > 0) {
                 body += "Bottom 5 (Opportunities by QTD Gap):\n";
                 bottom5Data.forEach((row, i) => {
                     body += `${i+1}. ${safeGet(row, 'Store')} - Gap: ${formatCurrency(calculateQtdGap(row))}\n`;
                 });
                 body += "\n";
             }
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
        if (!recipient || !/\S+@\S+\.\S+/.test(recipient)) {
            shareStatus.textContent = "Please enter a valid recipient email address.";
            return;
        }
        try {
            const subject = `FSM Dashboard Summary - ${new Date().toLocaleDateString()}`;
            const body = generateEmailBody();
            const mailtoLink = `mailto:${recipient}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
             if (mailtoLink.length > 2000) {
                 shareStatus.textContent = "Generated email body is too long for mailto link. Try applying more filters.";
                 console.warn("Mailto link length exceeds 2000 characters:", mailtoLink.length);
                 return;
             }
            window.location.href = mailtoLink;
            shareStatus.textContent = "Email client should open. Please review and send.";
        } catch (error) {
            console.error("Error generating mailto link:", error);
            shareStatus.textContent = "Error generating email content.";
        }
    };

     const selectAllOptions = (selectElement) => {
         if (!selectElement) return;
         Array.from(selectElement.options).forEach(option => option.selected = true);
         // Trigger hierarchy update after selection change
         if (selectElement === territoryFilter) {
            updateStoreFilterOptionsBasedOnHierarchy();
         }
    };

     const deselectAllOptions = (selectElement) => {
         if (!selectElement) return;
         selectElement.selectedIndex = -1; // Deselects all
          // Trigger hierarchy update after selection change
         if (selectElement === territoryFilter) {
             updateStoreFilterOptionsBasedOnHierarchy();
         }
    };


    // --- Event Listeners ---
    excelFileInput?.addEventListener('change', handleFile);
    applyFiltersButton?.addEventListener('click', applyFilters);
    storeSearch?.addEventListener('input', filterStoreOptions); // Update visual list on search input
    exportCsvButton?.addEventListener('click', exportData);
    shareEmailButton?.addEventListener('click', handleShareEmail);
    closeStoreDetailsButton?.addEventListener('click', hideStoreDetails);

    // Multi-select buttons
    territorySelectAll?.addEventListener('click', () => selectAllOptions(territoryFilter));
    territoryDeselectAll?.addEventListener('click', () => deselectAllOptions(territoryFilter));
    storeSelectAll?.addEventListener('click', () => selectAllOptions(storeFilter)); // No hierarchy update needed here
    storeDeselectAll?.addEventListener('click', () => deselectAllOptions(storeFilter)); // No hierarchy update needed here

    // Table Sorting (Attach Rate Table Only)
    attachRateTable?.querySelector('thead')?.addEventListener('click', handleSort);

    // --- Initial Setup ---
    resetUI(); // Ensure clean state on load

}); // End DOMContentLoaded
