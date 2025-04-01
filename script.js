/* Generated: 2025-04-01 12:35:53 AM EDT - Filter Attach Rate table to hide rows where all attach rate values are 0%. */
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
        // This helps differentiate between a missing value and an actual zero value.
        if (defaultValue === null && obj && obj[path] === 0) {
             return 0;
        }
        const value = obj ? obj[path] : undefined;
        // Return defaultValue only if value is undefined OR null (unless defaultValue is null and value is 0)
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
         // If either is NaN, the gap is conceptually undefined for sorting, treat as worst case (positive infinity for asc sort)
         if (isNaN(revenue) || isNaN(target)) {
             return Infinity; // Treat undefined gaps as highest possible for *ascending* sort (lowest gap first)
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
                updateTopBottomTables(filteredData); // Update Top/Bottom tables
                updateCharts(filteredData);
                updateAttachRateTable(filteredData); // Update Attach Rate table

                // 4. Auto Show/Hide Store Details based on Filter Count
                if (filteredData.length === 1) {
                    showStoreDetails(filteredData[0]);
                    highlightTableRow(safeGet(filteredData[0], 'Store', null));
                } else {
                    hideStoreDetails();
                }

                // 5. Finalize UI state
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
                 updateTopBottomTables([]);
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
             storeFilter.innerHTML = '<option value="ALL">-- Load File First --</option>';
             storeFilter.selectedIndex = -1;
             storeFilter.disabled = true;
         }
         if (storeSearch) { storeSearch.value = ''; storeSearch.disabled = true; }
         storeOptions = [];
         allPossibleStores = [];


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
         resetFilters();
         if (filterArea) filterArea.style.display = 'none';
         if (resultsArea) resultsArea.style.display = 'none';
         // Clear potential leftover data displays
         if (mainChartInstance) { mainChartInstance.destroy(); mainChartInstance = null; }
         if (attachRateTableBody) attachRateTableBody.innerHTML = '';
         if (attachRateTableFooter) attachRateTableFooter.innerHTML = '';
         if (attachTableStatus) attachTableStatus.textContent = '';
         if (topBottomSection) topBottomSection.style.display = 'none';
         if (top5TableBody) top5TableBody.innerHTML = '';
         if (bottom5TableBody) bottom5TableBody.innerHTML = '';
         hideStoreDetails();
         // Reset summary fields
         updateSummary([]);
         if(statusDiv) statusDiv.textContent = 'No file selected.';
     };

    // --- UPDATED updateSummary with new calculation logic ---
    const updateSummary = (data) => {
        const totalCount = data.length;

        // Clear summary fields first
        const fieldsToClear = [revenueWithDFValue, qtdRevenueTargetValue, qtdGapValue, quarterlyRevenueTargetValue,
                        percentQuarterlyStoreTargetValue, revARValue, unitsWithDFValue, unitTargetValue,
                        unitAchievementValue, visitCountValue, trainingCountValue, retailModeConnectivityValue,
                        repSkillAchValue, vPmrAchValue, postTrainingScoreValue, eliteValue,
                        percentQuarterlyTerritoryTargetValue, territoryRevPercentValue, districtRevPercentValue, regionRevPercentValue];
        fieldsToClear.forEach(el => { if (el) el.textContent = 'N/A'; });
        [percentQuarterlyTerritoryTargetP, territoryRevPercentP, districtRevPercentP, regionRevPercentP].forEach(p => { if(p) p.style.display = 'none';});

        if (totalCount === 0) {
            return;
        }

        // --- Calculate SUMS ---
        const sumRevenue = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Revenue w/DF', 0)), 0);
        const sumQtdTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'QTD Revenue Target', 0)), 0);
        const sumQuarterlyTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Quarterly Revenue Target', 0)), 0);
        const sumUnits = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Unit w/ DF', 0)), 0);
        const sumUnitTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Unit Target', 0)), 0);
        const sumVisits = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Visit count', 0)), 0);
        const sumTrainings = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Trainings', 0)), 0);

        // --- Calculate AVERAGES (excluding blanks/invalid values AND ZEROS where specified) ---
        let sumConnectivity = 0, countConnectivity = 0;
        let sumRepSkill = 0, countRepSkill = 0;
        let sumPmr = 0, countPmr = 0;
        let sumPostTraining = 0, countPostTraining = 0;
        let sumElite = 0, countElite = 0;

        data.forEach(row => {
            let val;
            // Connectivity, Rep Skill, PMR (Ignore Blanks & Zeros)
            val = safeGet(row, 'Retail Mode Connectivity', null); if (isValidAndNonZeroForAverage(val)) { sumConnectivity += parsePercent(val); countConnectivity++; }
            val = safeGet(row, 'Rep Skill Ach', null); if (isValidAndNonZeroForAverage(val)) { sumRepSkill += parsePercent(val); countRepSkill++; }
            val = safeGet(row, '(V)PMR Ach', null); if (isValidAndNonZeroForAverage(val)) { sumPmr += parsePercent(val); countPmr++; }
            // Post Training, Elite (Ignore Blanks, Include Zeros)
            val = safeGet(row, 'Post Training Score', null); if (isValidForAverage(val)) { sumPostTraining += parseNumber(val); countPostTraining++; }
            val = safeGet(row, 'Elite', null); if (isValidForAverage(val)) { sumElite += parsePercent(val); countElite++; }
        });

        const avgConnectivity = countConnectivity === 0 ? NaN : sumConnectivity / countConnectivity;
        const avgRepSkill = countRepSkill === 0 ? NaN : sumRepSkill / countRepSkill;
        const avgPmr = countPmr === 0 ? NaN : sumPmr / countPmr;
        const avgPostTraining = countPostTraining === 0 ? NaN : sumPostTraining / countPostTraining;
        const avgElite = countElite === 0 ? NaN : sumElite / countElite;

        // --- Calculate Overall Percentages ---
        const calculatedRevAR = sumQtdTarget === 0 ? 0 : sumRevenue / sumQtdTarget;
        const overallUnitAchievement = sumUnitTarget === 0 ? 0 : sumUnits / sumUnitTarget;
        const overallPercentStoreTarget = sumQuarterlyTarget === 0 ? 0 : sumRevenue / sumQuarterlyTarget;

        // --- Update DOM ---
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
        // Overall %
        revARValue && (revARValue.textContent = formatPercent(calculatedRevAR));
        revARValue && (revARValue.title = `Calculated Rev AR% (Total Revenue / Total QTD Target)`);
        unitAchievementValue && (unitAchievementValue.textContent = formatPercent(overallUnitAchievement));
        unitAchievementValue && (unitAchievementValue.title = `Overall Unit Achievement % (Total Units / Total Unit Target)`);
        percentQuarterlyStoreTargetValue && (percentQuarterlyStoreTargetValue.textContent = formatPercent(overallPercentStoreTarget));
        percentQuarterlyStoreTargetValue && (percentQuarterlyStoreTargetValue.title = `Overall % Quarterly Target (Total Revenue / Total Quarterly Target)`);
        // Averages
        retailModeConnectivityValue && (retailModeConnectivityValue.textContent = formatPercent(avgConnectivity));
        retailModeConnectivityValue && (retailModeConnectivityValue.title = `Average 'Retail Mode Connectivity' across ${countConnectivity} stores with non-zero data`);
        repSkillAchValue && (repSkillAchValue.textContent = formatPercent(avgRepSkill));
        repSkillAchValue && (repSkillAchValue.title = `Average 'Rep Skill Ach' across ${countRepSkill} stores with non-zero data`);
        vPmrAchValue && (vPmrAchValue.textContent = formatPercent(avgPmr));
        vPmrAchValue && (vPmrAchValue.title = `Average '(V)PMR Ach' across ${countPmr} stores with non-zero data`);
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
        const calculateAverageExcludeBlanks = (column) => { /* ... [implementation unchanged] ... */ }; // Assuming this is correct
        const avgPercentTerritoryTarget = calculateAverageExcludeBlanks('%Quarterly Territory Rev Target');
        const avgTerritoryRevPercent = calculateAverageExcludeBlanks('Territory Rev%');
        const avgDistrictRevPercent = calculateAverageExcludeBlanks('District Rev%');
        const avgRegionRevPercent = calculateAverageExcludeBlanks('Region Rev%');
        if (percentQuarterlyTerritoryTargetValue) percentQuarterlyTerritoryTargetValue.textContent = formatPercent(avgPercentTerritoryTarget);
        if (percentQuarterlyTerritoryTargetP) percentQuarterlyTerritoryTargetP.style.display = 'block';
        if (singleTerritory || singleDistrict || singleRegion) { if (territoryRevPercentValue) territoryRevPercentValue.textContent = formatPercent(avgTerritoryRevPercent); if (territoryRevPercentP) territoryRevPercentP.style.display = 'block'; }
        if (singleDistrict || singleRegion) { if (districtRevPercentValue) districtRevPercentValue.textContent = formatPercent(avgDistrictRevPercent); if (districtRevPercentP) districtRevPercentP.style.display = 'block'; }
        if (singleRegion) { if (regionRevPercentValue) regionRevPercentValue.textContent = formatPercent(avgRegionRevPercent); if (regionRevPercentP) regionRevPercentP.style.display = 'block'; }
    };

    const updateTopBottomTables = (data) => {
        if (!topBottomSection || !top5TableBody || !bottom5TableBody) return;
        top5TableBody.innerHTML = '';
        bottom5TableBody.innerHTML = '';
        const territoriesInData = new Set(data.map(row => safeGet(row, 'Q2 Territory', null)).filter(Boolean));
        const isSingleTerritory = territoriesInData.size === 1;
        if (!isSingleTerritory || data.length === 0) { topBottomSection.style.display = 'none'; return; }
        topBottomSection.style.display = 'flex';
        const top5Data = [...data].sort((a, b) => parseNumber(safeGet(b, 'Revenue w/DF', -Infinity)) - parseNumber(safeGet(a, 'Revenue w/DF', -Infinity))).slice(0, TOP_N_TABLES);
        top5Data.forEach(row => { /* ... [implementation unchanged] ... */ });
        const bottom5Data = [...data].sort((a, b) => calculateQtdGap(a) - calculateQtdGap(b)).slice(0, TOP_N_TABLES); // Corrected sort
        bottom5Data.forEach(row => { /* ... [implementation unchanged] ... */ });
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
        mainChartInstance = new Chart(mainChartCanvas, { /* ... [options unchanged] ... */ });
    };

    // ** UPDATED ** Attach Rate Table filtering logic
    const updateAttachRateTable = (data) => {
        if (!attachRateTableBody || !attachRateTableFooter) return;
        attachRateTableBody.innerHTML = '';
        attachRateTableFooter.innerHTML = '';

        // Original data length check for status/footer calculation
        const originalDataLength = data.length;
        if (originalDataLength === 0) {
            if(attachTableStatus) attachTableStatus.textContent = 'No data to display based on filters.';
            return;
        }

        // Filter out rows where ALL attach rates are 0%
        const attachDataToShow = data.filter(row => {
            // Use ATTACH_RATE_KEYS defined globally
            return ATTACH_RATE_KEYS.some(key => {
                // Use safeGet with null default to parse correctly, treat NaN from parse as 0 for this check
                const val = parsePercent(safeGet(row, key, null));
                return !isNaN(val) && val > 0;
            });
        });

        // Sort the filtered data based on currentSort state
        const sortedAttachData = [...attachDataToShow].sort((a, b) => {
             let valA = safeGet(a, currentSort.column, null);
             let valB = safeGet(b, currentSort.column, null);
             if (valA === null && valB === null) return 0;
             if (valA === null) return currentSort.ascending ? -1 : 1;
             if (valB === null) return currentSort.ascending ? 1 : -1;
             const isPercentCol = currentSort.column.includes('Attach Rate');
             const numA = isPercentCol ? parsePercent(valA) : parseNumber(valA);
             const numB = isPercentCol ? parsePercent(valB) : parseNumber(valB);
             if (!isNaN(numA) && !isNaN(numB)) { return currentSort.ascending ? numA - numB : numB - numA; }
             else { valA = String(valA).toLowerCase(); valB = String(valB).toLowerCase(); return currentSort.ascending ? valA.localeCompare(valB) : valB.localeCompare(valA); }
         });

        // Calculate averages for highlighting AND footer (use original 'data' for averages)
        const averages = {};
        ATTACH_RATE_KEYS.forEach(key => { // Use ATTACH_RATE_KEYS
             let sum = 0; let count = 0;
             data.forEach(row => { // Iterate original data for averages
                 const val = safeGet(row, key, null);
                 if (isValidForAverage(val)) { sum += parsePercent(val); count++; }
             });
             averages[key] = count === 0 ? NaN : sum / count;
         });

        // Populate table body using the filtered & sorted data
        sortedAttachData.forEach(row => {
            const tr = document.createElement('tr');
            const storeName = safeGet(row, 'Store', null);
            if (storeName) {
                 tr.dataset.storeName = storeName;
                 tr.onclick = () => { showStoreDetails(row); highlightTableRow(storeName); };
                 // Define columns for the attach rate table
                 const columns = [
                     { key: 'Store', format: (val) => val },
                     // Use ATTACH_RATE_KEYS to generate columns dynamically if needed, or list manually
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
                     let formattedValue;
                     if (rawValue === null || (col.key !== 'Store' && isNaN(numericValue))) { formattedValue = 'N/A'; }
                     else { formattedValue = isPercentCol ? col.format(numericValue) : rawValue; }
                     td.textContent = formattedValue;
                     td.title = `${col.key}: ${formattedValue}`;
                     if (col.highlight && averages[col.key] !== undefined && !isNaN(numericValue)) {
                         td.classList.add(numericValue >= averages[col.key] ? 'highlight-green' : 'highlight-red');
                     }
                     tr.appendChild(td);
                 });
                 attachRateTableBody.appendChild(tr);
             }
        });

        // Add Average Row to Footer (calculated from original 'data')
        if (originalDataLength > 0) {
            const footerRow = attachRateTableFooter.insertRow();
            const avgLabelCell = footerRow.insertCell();
            avgLabelCell.textContent = 'Filtered Avg*';
            avgLabelCell.title = 'Average calculated only using stores with valid data for each column from the initial filter selection';
            avgLabelCell.style.textAlign = "right";
            avgLabelCell.style.fontWeight = "bold";
            // Add the Store column blank cell
            // footerRow.insertCell(); // Add blank cell for store name alignment if needed

            // Use ATTACH_RATE_KEYS for footer cells
            ATTACH_RATE_KEYS.forEach(key => {
                 const td = footerRow.insertCell();
                 const avgValue = averages[key];
                 td.textContent = formatPercent(avgValue);
                 let validCount = 0;
                 data.forEach(row => { if (isValidForAverage(safeGet(row, key, null))) validCount++; });
                 td.title = `Average ${key}: ${formatPercent(avgValue)} (from ${validCount} stores)`;
                 td.style.textAlign = "right";
             });
         }

        // Update status text based on the length of the *shown* data
        if(attachTableStatus) attachTableStatus.textContent = `Showing ${attachDataToShow.length} stores with non-zero attach rates. Click row for details. Click headers to sort.`;
        updateSortArrows(); // Update arrows based on currentSort state
    };

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
    territorySelectAll?.addEventListener('click', () => selectAllOptions(territoryFilter));
    territoryDeselectAll?.addEventListener('click', () => deselectAllOptions(territoryFilter));
    storeSelectAll?.addEventListener('click', () => selectAllOptions(storeFilter));
    storeDeselectAll?.addEventListener('click', () => deselectAllOptions(storeFilter));
    attachRateTable?.querySelector('thead')?.addEventListener('click', handleSort);

    // --- Initial Setup ---
    resetUI();

}); // End DOMContentLoaded
