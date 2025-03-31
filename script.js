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

    // --- DOM Elements ---
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
    let currentSort = { column: 'Store', ascending: true };
    let selectedStoreRow = null;

    // --- Helper Functions ---
    const formatCurrency = (value) => isNaN(value) ? 'N/A' : CURRENCY_FORMAT.format(value);
    const formatPercent = (value) => isNaN(value) ? 'N/A' : PERCENT_FORMAT.format(value);
    const formatNumber = (value) => isNaN(value) ? 'N/A' : NUMBER_FORMAT.format(value);
    const parseNumber = (value) => {
        if (typeof value === 'number') return value;
        if (typeof value === 'string') {
            value = value.replace(/[$,%]/g, '');
            const num = parseFloat(value);
            return isNaN(num) ? 0 : num;
        }
        return 0;
    };
    const parsePercent = (value) => {
         if (typeof value === 'number') return value; // Assume it's already a decimal representation
         if (typeof value === 'string') {
            const num = parseFloat(value.replace('%', ''));
            return isNaN(num) ? 0 : num / 100; // Convert percentage string to decimal
         }
         return 0;
    };
    const safeGet = (obj, path, defaultValue = 'N/A') => {
        const value = obj ? obj[path] : undefined;
        // Consider null/empty strings as potentially valid, adjust if needed
        // Return defaultValue if value is null or undefined, but allow empty strings
        return (value !== undefined && value !== null) ? value : defaultValue;
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

     // NEW function to update store filter options based on other selections
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
                updateCharts(filteredData);
                updateAttachRateTable(filteredData); // This already calls hideStoreDetails initially

                // 4. Auto Show/Hide Store Details based on Filter Count
                if (filteredData.length === 1) {
                    showStoreDetails(filteredData[0]);
                    highlightTableRow(safeGet(filteredData[0], 'Store', null));
                } else {
                    // Ensures details are hidden if filters change to >1 or 0 stores
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
                 updateCharts([]);
                 updateAttachRateTable([]);
                 hideStoreDetails(); // Ensure details are hidden on error too

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
         hideStoreDetails();
         // Reset summary fields
         updateSummary([]);
         if(statusDiv) statusDiv.textContent = 'No file selected.'; // Reset status message
     };

    const updateSummary = (data) => {
        const count = data.length;
        if (count === 0) {
            // Reset summary fields
            const fields = [revenueWithDFValue, qtdRevenueTargetValue, qtdGapValue, quarterlyRevenueTargetValue,
                            percentQuarterlyStoreTargetValue, revARValue, unitsWithDFValue, unitTargetValue,
                            unitAchievementValue, visitCountValue, trainingCountValue, retailModeConnectivityValue,
                            repSkillAchValue, vPmrAchValue, postTrainingScoreValue, eliteValue,
                            percentQuarterlyTerritoryTargetValue, territoryRevPercentValue, districtRevPercentValue, regionRevPercentValue];
            // Use optional chaining in case elements aren't found
            fields.forEach(el => { if (el) el.textContent = 'N/A'; });
            [percentQuarterlyTerritoryTargetP, territoryRevPercentP, districtRevPercentP, regionRevPercentP].forEach(p => { if(p) p.style.display = 'none';});
            return;
        }

        // Calculate sums and averages (ensure safeGet returns 0 for calculations if value is missing/null)
        const sumRevenue = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Revenue w/DF', 0)), 0);
        const sumQtdTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'QTD Revenue Target', 0)), 0);
        const sumQuarterlyTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Quarterly Revenue Target', 0)), 0);
        const sumUnits = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Unit w/ DF', 0)), 0);
        const sumUnitTarget = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Unit Target', 0)), 0);
        const sumVisits = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Visit count', 0)), 0);
        const sumTrainings = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Trainings', 0)), 0);
        const sumConnectivity = data.reduce((sum, row) => sum + parsePercent(safeGet(row, 'Retail Mode Connectivity', 0)), 0);
        const sumRepSkill = data.reduce((sum, row) => sum + parsePercent(safeGet(row, 'Rep Skill Ach', 0)), 0);
        const sumPmr = data.reduce((sum, row) => sum + parsePercent(safeGet(row, '(V)PMR Ach', 0)), 0);
        const sumPostTraining = data.reduce((sum, row) => sum + parseNumber(safeGet(row, 'Post Training Score', 0)), 0); // Assuming score is number, not %
        const sumElite = data.reduce((sum, row) => sum + parsePercent(safeGet(row, 'Elite', 0)), 0);
        const sumRevAR = data.reduce((sum, row) => sum + parsePercent(safeGet(row, 'Rev AR%', 0)), 0);

        // Calculate overall percentages
        const overallPercentStoreTarget = sumQuarterlyTarget === 0 ? 0 : sumRevenue / sumQuarterlyTarget;
        const overallUnitAchievement = sumUnitTarget === 0 ? 0 : sumUnits / sumUnitTarget;

        // Update DOM elements safely using optional chaining
        revenueWithDFValue && (revenueWithDFValue.textContent = formatCurrency(sumRevenue));
        revenueWithDFValue && (revenueWithDFValue.title = `Sum of 'Revenue w/DF' for ${count} stores`);
        qtdRevenueTargetValue && (qtdRevenueTargetValue.textContent = formatCurrency(sumQtdTarget));
        qtdRevenueTargetValue && (qtdRevenueTargetValue.title = `Sum of 'QTD Revenue Target' for ${count} stores`);
        qtdGapValue && (qtdGapValue.textContent = formatCurrency(sumRevenue - sumQtdTarget));
        qtdGapValue && (qtdGapValue.title = `Calculated Gap (Total Revenue - QTD Target) for ${count} stores`);
        quarterlyRevenueTargetValue && (quarterlyRevenueTargetValue.textContent = formatCurrency(sumQuarterlyTarget));
        quarterlyRevenueTargetValue && (quarterlyRevenueTargetValue.title = `Sum of 'Quarterly Revenue Target' for ${count} stores`);

        percentQuarterlyStoreTargetValue && (percentQuarterlyStoreTargetValue.textContent = formatPercent(overallPercentStoreTarget));
        percentQuarterlyStoreTargetValue && (percentQuarterlyStoreTargetValue.title = `Overall % Quarterly Target (Total Revenue / Total Quarterly Target)`);

        revARValue && (revARValue.textContent = formatPercent(sumRevAR / count)); // Average Rev AR%
        revARValue && (revARValue.title = `Average 'Rev AR%' across ${count} stores`);

        unitsWithDFValue && (unitsWithDFValue.textContent = formatNumber(sumUnits));
        unitsWithDFValue && (unitsWithDFValue.title = `Sum of 'Unit w/ DF' for ${count} stores`);
        unitTargetValue && (unitTargetValue.textContent = formatNumber(sumUnitTarget));
        unitTargetValue && (unitTargetValue.title = `Sum of 'Unit Target' for ${count} stores`);
        unitAchievementValue && (unitAchievementValue.textContent = formatPercent(overallUnitAchievement));
        unitAchievementValue && (unitAchievementValue.title = `Overall Unit Achievement % (Total Units / Total Unit Target)`);

        visitCountValue && (visitCountValue.textContent = formatNumber(sumVisits));
        visitCountValue && (visitCountValue.title = `Sum of 'Visit count' for ${count} stores`);
        trainingCountValue && (trainingCountValue.textContent = formatNumber(sumTrainings));
        trainingCountValue && (trainingCountValue.title = `Sum of 'Trainings' for ${count} stores`);
        retailModeConnectivityValue && (retailModeConnectivityValue.textContent = formatPercent(sumConnectivity / count));
        retailModeConnectivityValue && (retailModeConnectivityValue.title = `Average 'Retail Mode Connectivity' across ${count} stores`);

        repSkillAchValue && (repSkillAchValue.textContent = formatPercent(sumRepSkill / count));
        repSkillAchValue && (repSkillAchValue.title = `Average 'Rep Skill Ach' across ${count} stores`);
        vPmrAchValue && (vPmrAchValue.textContent = formatPercent(sumPmr / count));
        vPmrAchValue && (vPmrAchValue.title = `Average '(V)PMR Ach' across ${count} stores`);
        postTrainingScoreValue && (postTrainingScoreValue.textContent = formatNumber((sumPostTraining / count).toFixed(1))); // Avg score, 1 decimal
        postTrainingScoreValue && (postTrainingScoreValue.title = `Average 'Post Training Score' across ${count} stores`);
        eliteValue && (eliteValue.textContent = formatPercent(sumElite / count));
        eliteValue && (eliteValue.title = `Average 'Elite' score % across ${count} stores`);

        // Contextual Hierarchy Percentages
        updateContextualSummary(data);
    };

    const updateContextualSummary = (data) => {
        // Hide all contextual fields initially safely
        [percentQuarterlyTerritoryTargetP, territoryRevPercentP, districtRevPercentP, regionRevPercentP].forEach(p => {if (p) p.style.display = 'none'});

        if (data.length === 0) return;

        // Calculate average contribution percentages IF the filter scope makes sense
        const singleRegion = regionFilter.value !== 'ALL';
        const singleDistrict = districtFilter.value !== 'ALL';
        const singleTerritory = territoryFilter.selectedOptions.length === 1;

        // We can always calculate the average % of Territory Target contribution
        const avgPercentTerritoryTarget = data.reduce((sum, row) => sum + parsePercent(safeGet(row, '%Quarterly Territory Rev Target', 0)), 0) / data.length;
        if (percentQuarterlyTerritoryTargetValue) percentQuarterlyTerritoryTargetValue.textContent = formatPercent(avgPercentTerritoryTarget);
        if (percentQuarterlyTerritoryTargetP) percentQuarterlyTerritoryTargetP.style.display = 'block'; // Show this one generally

        if (singleTerritory || singleDistrict || singleRegion) {
             const avgTerritoryRevPercent = data.reduce((sum, row) => sum + parsePercent(safeGet(row, 'Territory Rev%', 0)), 0) / data.length;
             if (territoryRevPercentValue) territoryRevPercentValue.textContent = formatPercent(avgTerritoryRevPercent);
             if (territoryRevPercentP) territoryRevPercentP.style.display = 'block';
        }
        if (singleDistrict || singleRegion) {
             const avgDistrictRevPercent = data.reduce((sum, row) => sum + parsePercent(safeGet(row, 'District Rev%', 0)), 0) / data.length;
             if (districtRevPercentValue) districtRevPercentValue.textContent = formatPercent(avgDistrictRevPercent);
             if (districtRevPercentP) districtRevPercentP.style.display = 'block';
        }
         if (singleRegion) {
             const avgRegionRevPercent = data.reduce((sum, row) => sum + parsePercent(safeGet(row, 'Region Rev%', 0)), 0) / data.length;
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

    const updateAttachRateTable = (data) => {
        if (!attachRateTableBody || !attachRateTableFooter) return; // Exit if elements don't exist
        attachRateTableBody.innerHTML = '';
        attachRateTableFooter.innerHTML = '';
        // Don't automatically hide details here, applyFilters handles that based on count

        if (data.length === 0) {
            if(attachTableStatus) attachTableStatus.textContent = 'No data to display based on filters.';
            return;
        }

        // Sort data based on currentSort state
        const sortedData = [...data].sort((a, b) => {
             let valA = safeGet(a, currentSort.column, null); // Use null default for comparison consistency
             let valB = safeGet(b, currentSort.column, null);

             // Handle nulls first: consistently place them at the beginning or end
             if (valA === null && valB === null) return 0;
             if (valA === null) return currentSort.ascending ? -1 : 1; // Nulls first when ascending
             if (valB === null) return currentSort.ascending ? 1 : -1; // Nulls first when ascending

             // Attempt numeric conversion for sorting if applicable (percentages or numbers)
             const isPercentCol = currentSort.column.includes('Attach Rate') || currentSort.column.includes('%') || currentSort.column.includes('Target'); // Adjust column check as needed
             const numA = isPercentCol ? parsePercent(valA) : parseNumber(valA);
             const numB = isPercentCol ? parsePercent(valB) : parseNumber(valB);

             if (!isNaN(numA) && !isNaN(numB)) {
                 // Numeric comparison
                 return currentSort.ascending ? numA - numB : numB - numA;
             } else {
                 // String comparison (case-insensitive)
                 valA = String(valA).toLowerCase();
                 valB = String(valB).toLowerCase();
                 return currentSort.ascending ? valA.localeCompare(valB) : valB.localeCompare(valA);
             }
         });


        // Calculate averages for highlighting (handle potential division by zero)
        const count = data.length;
        const averages = {
             'Tablet Attach Rate': count === 0 ? 0 : data.reduce((sum, r) => sum + parsePercent(safeGet(r, 'Tablet Attach Rate', 0)), 0) / count,
             'PC Attach Rate': count === 0 ? 0 : data.reduce((sum, r) => sum + parsePercent(safeGet(r, 'PC Attach Rate', 0)), 0) / count,
             'NC Attach Rate': count === 0 ? 0 : data.reduce((sum, r) => sum + parsePercent(safeGet(r, 'NC Attach Rate', 0)), 0) / count,
             'TWS Attach Rate': count === 0 ? 0 : data.reduce((sum, r) => sum + parsePercent(safeGet(r, 'TWS Attach Rate', 0)), 0) / count,
             'WW Attach Rate': count === 0 ? 0 : data.reduce((sum, r) => sum + parsePercent(safeGet(r, 'WW Attach Rate', 0)), 0) / count,
             'ME Attach Rate': count === 0 ? 0 : data.reduce((sum, r) => sum + parsePercent(safeGet(r, 'ME Attach Rate', 0)), 0) / count,
             'NCME Attach Rate': count === 0 ? 0 : data.reduce((sum, r) => sum + parsePercent(safeGet(r, 'NCME Attach Rate', 0)), 0) / count,
             '% Quarterly Revenue Target': count === 0 ? 0 : data.reduce((sum, r) => sum + parsePercent(safeGet(r, '% Quarterly Revenue Target', 0)), 0) / count,
        };

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
                 const columns = [
                     { key: 'Store', format: (val) => val }, // Keep store name as is
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
                     const rawValue = safeGet(row, col.key, null); // Get raw value safely
                     // Determine if parsing as percent is needed
                     const isPercentCol = col.key.includes('Attach Rate') || col.key.includes('%');
                     const numericValue = isPercentCol ? parsePercent(rawValue) : parseNumber(rawValue); // Use number for non-% too

                     let formattedValue;
                     if (rawValue === null || (col.key !== 'Store' && isNaN(numericValue))) {
                         formattedValue = 'N/A'; // Display N/A if original is null or parsing fails (except for Store name)
                     } else {
                         // If it's the Store column, use the rawValue directly. Otherwise format.
                         formattedValue = (col.key === 'Store') ? rawValue : col.format(isPercentCol ? numericValue : rawValue);
                     }

                     td.textContent = formattedValue;
                     td.title = `${col.key}: ${formattedValue}`; // Add tooltip

                     // Apply highlighting based on average (only if value is numeric)
                     if (col.highlight && averages[col.key] !== undefined && !isNaN(numericValue) && rawValue !== null) {
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

        // Add Average Row to Footer
        if (count > 0) {
            const footerRow = attachRateTableFooter.insertRow();
            footerRow.insertCell().textContent = 'Filtered Avg';
            footerRow.cells[0].style.textAlign = "right";
            footerRow.cells[0].style.fontWeight = "bold";

            const footerColumns = [
                '% Quarterly Revenue Target', 'Tablet Attach Rate', 'PC Attach Rate', 'NC Attach Rate',
                 'TWS Attach Rate', 'WW Attach Rate', 'ME Attach Rate', 'NCME Attach Rate'
            ];
             footerColumns.forEach(key => {
                 const td = footerRow.insertCell();
                 const avgValue = averages[key];
                 td.textContent = formatPercent(avgValue);
                 td.title = `Average ${key}: ${formatPercent(avgValue)}`;
                 td.style.textAlign = "right";
             });
         }

        if(attachTableStatus) attachTableStatus.textContent = `Showing ${count} stores. Click row for details. Click headers to sort.`;
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
         // IMPORTANT: Update the table using the currently *globally filtered* data
         updateAttachRateTable(filteredData);
    };


    const updateSortArrows = () => {
        if (!attachRateTable) return;
        attachRateTable.querySelectorAll('th.sortable .sort-arrow').forEach(arrow => {
            arrow.classList.remove('asc', 'desc');
            arrow.textContent = ''; // Clear arrow text
        });
        const currentHeader = attachRateTable.querySelector(`th[data-sort="${currentSort.column}"] .sort-arrow`);
        if (currentHeader) {
            currentHeader.classList.add(currentSort.ascending ? 'asc' : 'desc');
            // Add arrow character for visual cue
            currentHeader.textContent = currentSort.ascending ? ' ▲' : ' ▼'; // Add space before arrow
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
            const mapsUrl = `https://www.google.com/maps?q=${latitude},${longitude}`;
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
        // Optional: scroll into view only if triggered by click, not auto-filter?
        // storeDetailsSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    };

    const hideStoreDetails = () => {
        if (!storeDetailsContent || !storeDetailsSection || !closeStoreDetailsButton) return;
        storeDetailsContent.innerHTML = 'Select a store from the table or chart for details, or apply filters resulting in a single store.';
        storeDetailsSection.style.display = 'none';
        closeStoreDetailsButton.style.display = 'none';
        highlightTableRow(null); // Remove table row highlight
    };

     const highlightTableRow = (storeName) => {
         if (selectedStoreRow) {
             selectedStoreRow.classList.remove('selected-row');
         }
        if (storeName && attachRateTableBody) {
             try {
                 // Use querySelector which handles special characters in attribute values better
                 selectedStoreRow = attachRateTableBody.querySelector(`tr[data-store-name="${CSS.escape(storeName)}"]`);
                 if (selectedStoreRow) {
                     selectedStoreRow.classList.add('selected-row');
                 } else {
                     selectedStoreRow = null; // Store not found in current table view
                 }
             } catch (e) {
                 console.error("Error selecting table row:", e);
                 selectedStoreRow = null;
             }
         } else {
             selectedStoreRow = null;
         }
    };


    const exportData = () => {
        if (filteredData.length === 0) {
            alert("No filtered data to export.");
            return;
        }
        try {
            if (!attachRateTable) throw new Error("Attach rate table not found.");
            const headers = Array.from(attachRateTable.querySelectorAll('thead th')).map(th => th.dataset.sort || th.textContent.replace(/ [▲▼]$/, '').trim()); // Get clean headers
             const dataToExport = filteredData.map(row => {
                return headers.map(header => {
                    let value = safeGet(row, header, ''); // Get raw value, default to empty string
                     // Check if the column header indicates a percentage or rate
                    const isPercentLike = header.includes('%') || header.includes('Rate') || header.includes('Ach') || header.includes('Connectivity') || header.includes('Elite');

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
        body += `- Rev AR%: ${revARValue?.textContent || 'N/A'}\n`;
        body += `- Total Units (incl. DF): ${unitsWithDFValue?.textContent || 'N/A'}\n`;
        body += `- Unit Achievement %: ${unitAchievementValue?.textContent || 'N/A'}\n`;
        body += `- Total Visits: ${visitCountValue?.textContent || 'N/A'}\n`;
        body += `- Avg. Connectivity: ${retailModeConnectivityValue?.textContent || 'N/A'}\n\n`;
        body += "Mysteryshop & Training (Avg):\n";
        body += `- Rep Skill Ach: ${repSkillAchValue?.textContent || 'N/A'}\n`;
        body += `- (V)PMR Ach: ${vPmrAchValue?.textContent || 'N/A'}\n`;
        body += `- Post Training Score: ${postTrainingScoreValue?.textContent || 'N/A'}\n`;
        body += `- Elite Score %: ${eliteValue?.textContent || 'N/A'}\n\n`;

        // Use current sort order from the table for top stores
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
    };

     const deselectAllOptions = (selectElement) => {
         if (!selectElement) return;
         selectElement.selectedIndex = -1; // Deselects all
    };


    // --- Event Listeners ---
    excelFileInput?.addEventListener('change', handleFile);
    applyFiltersButton?.addEventListener('click', applyFilters);
    storeSearch?.addEventListener('input', filterStoreOptions); // Update visual list on search input
    exportCsvButton?.addEventListener('click', exportData);
    shareEmailButton?.addEventListener('click', handleShareEmail);
    closeStoreDetailsButton?.addEventListener('click', hideStoreDetails);

    // Multi-select buttons (Ensure they call hierarchy update if needed)
    territorySelectAll?.addEventListener('click', () => { selectAllOptions(territoryFilter); updateStoreFilterOptionsBasedOnHierarchy(); });
    territoryDeselectAll?.addEventListener('click', () => { deselectAllOptions(territoryFilter); updateStoreFilterOptionsBasedOnHierarchy(); });
    storeSelectAll?.addEventListener('click', () => selectAllOptions(storeFilter)); // No hierarchy update needed here
    storeDeselectAll?.addEventListener('click', () => deselectAllOptions(storeFilter)); // No hierarchy update needed here

    // Table Sorting (Check if thead exists)
    attachRateTable?.querySelector('thead')?.addEventListener('click', handleSort);

    // --- Initial Setup ---
    resetUI(); // Ensure clean state on load

}); // End DOMContentLoaded
