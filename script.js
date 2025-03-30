// --- Global Variables ---
let originalData = [];
let lineChartInstance = null;

// --- HTML Element References ---
const fileInput = document.getElementById('excelFile');
const statusElement = document.getElementById('status');
const summaryElement = document.getElementById('summaryData');
const storeDetailsElement = document.getElementById('storeDetailsSection');
const lineChartCanvas = document.getElementById('lineChartCanvas').getContext('2d');
const attachRateTableBody = document.getElementById('attachRateTableBody');
const attachRateTableContainer = document.getElementById('attachRateTableContainer');
const attachRateTableHeading = attachRateTableContainer.querySelector('h2');
const attachTableStatusElement = document.getElementById('attachTableStatus');
const territoryFilter = document.getElementById('territoryFilter'); // Now first logically
const subchannelFilter = document.getElementById('subchannelFilter');
const storeFilter = document.getElementById('storeFilter');
const repPmrSection = document.getElementById('repPmrSection');
const repSkillAchValue = document.getElementById('repSkillAchValue');
const vPmrAchValue = document.getElementById('vPmrAchValue');
const trainingStatsSection = document.getElementById('trainingStatsSection');
const eliteValue = document.getElementById('eliteValue');
const eliteP = document.getElementById('eliteP');
const postTrainingScoreValue = document.getElementById('postTrainingScoreValue');
const postTrainingP = document.getElementById('postTrainingP');
const emailRecipientInput = document.getElementById('emailRecipient');
const shareEmailButton = document.getElementById('shareEmailButton');
const shareStatusElement = document.getElementById('shareStatus');


// --- Event Listeners ---
fileInput.addEventListener('change', handleFileSelect, false);
territoryFilter.addEventListener('change', filterAndDisplayData); // Add listener
subchannelFilter.addEventListener('change', handleSubchannelChange);
storeFilter.addEventListener('change', filterAndDisplayData);
shareEmailButton.addEventListener('click', handleShareEmail);


// --- Main File Handling Logic ---
function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) { statusElement.textContent = 'No file selected.'; clearDashboard(); return; }
    statusElement.textContent = `Processing file: ${file.name}...`;
    console.warn("Processing all rows. This might be slow for very large files."); // Performance warning
    clearDashboard();
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = e.target.result;
            const workbook = XLSX.read(data, { type: 'binary' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            // Use all rows
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
            const rowLimitMessage = ` Processed all ${jsonData.length} data rows (Excel rows 2-${jsonData.length + 1}).`;

            if (!jsonData || jsonData.length === 0) {
                 statusElement.textContent = 'Error: No data found in the first sheet or file is empty.';
                return;
            }
            originalData = jsonData;
            populateFiltersOnInit(originalData);
            filterAndDisplayData();
            statusElement.textContent = `Successfully processed ${file.name}.${rowLimitMessage} Displaying data. Use filters to refine.`;
        } catch (error) {
            console.error("Error processing Excel file:", error);
            statusElement.textContent = `Error processing file: ${error.message}. Check console for details.`;
            clearDashboard();
        }
    };
    reader.onerror = function(ex) { console.error("FileReader error:", ex); statusElement.textContent = 'Error reading file.'; clearDashboard(); };
    reader.readAsBinaryString(file);
}

// --- Filter Population Logic ---
function populateFiltersOnInit(data) {
    const subchannels = new Set();
    const territories = new Set();
    const stores = new Set();
    let columnsPresent = true;
    const requiredFilterCols = ['SUB_CHANNEL', 'Q2 Territory', 'Store'];

    if (data.length === 0 || !data[0]) {
        columnsPresent = false;
    } else {
        requiredFilterCols.forEach(col => {
            if (data[0][col] === undefined) {
                columnsPresent = false;
                console.warn(`Filter column missing: ${col}`);
                if(!statusElement.textContent.includes('Warning:')) statusElement.textContent += ' Warning: ';
                statusElement.textContent += ` Filter column '${col}' missing.`;
            }
        });
    }

    if (columnsPresent) {
        data.forEach(row => {
            if (row['SUB_CHANNEL'] && String(row['SUB_CHANNEL']).trim()) subchannels.add(String(row['SUB_CHANNEL']).trim());
            if (row['Q2 Territory'] && String(row['Q2 Territory']).trim()) territories.add(String(row['Q2 Territory']).trim());
            if (row['Store'] && String(row['Store']).trim()) stores.add(String(row['Store']).trim());
        });
    } else {
         if(!statusElement.textContent.includes('Warning:')) statusElement.textContent += ' Warning: Could not fully populate filters due to missing columns.';
    }

    updateDropdownOptions(territoryFilter, [...territories].sort(), 'Territory');
    updateDropdownOptions(subchannelFilter, [...subchannels].sort(), 'Subchannel');
    updateDropdownOptions(storeFilter, [...stores].sort(), 'Store');

    resetMultiSelect(territoryFilter, 'ALL');
    resetMultiSelect(storeFilter, 'ALL');
    subchannelFilter.disabled = !columnsPresent || subchannels.size === 0;
    territoryFilter.disabled = !columnsPresent || territories.size === 0;
    storeFilter.disabled = !columnsPresent || stores.size === 0;
}
function updateDropdownOptions(selectElement, optionsArray, filterName) { const isMultiple = selectElement.multiple; selectElement.innerHTML = ''; const allOptionValue = 'ALL'; const allOptionText = `-- All ${filterName}s --`; selectElement.add(new Option(allOptionText, allOptionValue)); optionsArray.forEach(optionValue => { if (optionValue !== allOptionValue) { selectElement.add(new Option(optionValue, optionValue)); } }); if (!isMultiple) { selectElement.value = allOptionValue; } }
function updateStoreFilterOptions() {
    const selectedSubchannel = subchannelFilter.value;
    const relevantStores = new Set();
    let dataToScan = originalData;
    if (selectedSubchannel !== 'ALL') { dataToScan = dataToScan.filter(row => String(row['SUB_CHANNEL'] || '').trim() === selectedSubchannel); }
    // TODO: Consider filtering stores by selected territory as well?
    if (dataToScan.length > 0 && dataToScan[0]['Store'] !== undefined) { dataToScan.forEach(row => { if (row['Store'] && String(row['Store']).trim()) relevantStores.add(String(row['Store']).trim()); }); }
    else if (originalData.length > 0 && (!originalData[0] || originalData[0]['Store'] === undefined)) { if (!statusElement.textContent.includes("Warning: 'Store' column")) statusElement.textContent += ` Warning: 'Store' column may be missing, Store filter cannot be updated.`; }
    updateDropdownOptions(storeFilter, [...relevantStores].sort(), 'Store'); resetMultiSelect(storeFilter, 'ALL');
}

// --- Helper Function to Reset Multi-Select ---
function resetMultiSelect(selectElement, valueToSelect) { const valuesToSelect = Array.isArray(valueToSelect) ? valueToSelect : [valueToSelect]; for (let i = 0; i < selectElement.options.length; i++) { selectElement.options[i].selected = valuesToSelect.includes(selectElement.options[i].value); } }

// --- Event Handlers ---
function handleSubchannelChange() { updateStoreFilterOptions(); filterAndDisplayData(); }

// --- Data Filtering and Display Trigger ---
function filterAndDisplayData() {
    // Read selections
    const selectedSubchannel = subchannelFilter.value;
    // Territory Filter Logic
    const selectedTerritoryOptions = Array.from(territoryFilter.selectedOptions); let selectedTerritories = selectedTerritoryOptions.map(option => option.value); const allTerritoriesOption = territoryFilter.querySelector('option[value="ALL"]'); const containsAllTerr = selectedTerritories.includes('ALL'); const containsIndividualTerr = selectedTerritories.some(val => val !== 'ALL');
    if (containsAllTerr && containsIndividualTerr) { if (allTerritoriesOption) allTerritoriesOption.selected = false; selectedTerritories = selectedTerritories.filter(val => val !== 'ALL'); }
    else if (!containsAllTerr && selectedTerritories.length === 0 && allTerritoriesOption) { allTerritoriesOption.selected = true; selectedTerritories = ['ALL']; }
    else if (containsAllTerr && !containsIndividualTerr && selectedTerritories.length === 1 && allTerritoriesOption) { if (!allTerritoriesOption.selected) allTerritoriesOption.selected = true; for(let i=0; i<territoryFilter.options.length; i++){ if(territoryFilter.options[i].value !== 'ALL') territoryFilter.options[i].selected = false; }}
    const isAllTerritoriesSelected = selectedTerritories.includes('ALL');
    // Store Filter Logic
    const selectedStoreOptions = Array.from(storeFilter.selectedOptions); let selectedStores = selectedStoreOptions.map(option => option.value); const allStoresOption = storeFilter.querySelector('option[value="ALL"]'); const containsAllStores = selectedStores.includes('ALL'); const containsIndividualStores = selectedStores.some(val => val !== 'ALL');
    if (containsAllStores && containsIndividualStores) { if (allStoresOption) allStoresOption.selected = false; selectedStores = selectedStores.filter(val => val !== 'ALL'); }
    else if (!containsAllStores && selectedStores.length === 0 && allStoresOption) { allStoresOption.selected = true; selectedStores = ['ALL']; }
    else if (containsAllStores && !containsIndividualStores && selectedStores.length === 1 && allStoresOption) { if (!allStoresOption.selected) allStoresOption.selected = true; for(let i=0; i<storeFilter.options.length; i++){ if(storeFilter.options[i].value !== 'ALL') storeFilter.options[i].selected = false; }}
    const isAllStoresSelected = selectedStores.includes('ALL');
    const isSingleStoreSelected = !isAllStoresSelected && selectedStores.length === 1;
    // Apply Filters
    let fullyFilteredData = originalData;
    if (selectedSubchannel !== 'ALL') { fullyFilteredData = fullyFilteredData.filter(row => String(row['SUB_CHANNEL'] || '').trim() === selectedSubchannel); }
    if (!isAllTerritoriesSelected) { fullyFilteredData = fullyFilteredData.filter(row => selectedTerritories.includes(String(row['Q2 Territory'] || '').trim())); }
    if (!isAllStoresSelected) { fullyFilteredData = fullyFilteredData.filter(row => selectedStores.includes(String(row['Store'] || '').trim())); }
    // Prepare subchannel data for specific case
    let subchannelFilteredData = originalData;
    if (selectedSubchannel !== 'ALL') { subchannelFilteredData = subchannelFilteredData.filter(row => String(row['SUB_CHANNEL'] || '').trim() === selectedSubchannel); }
    // Update status
    const baseMessage = `Displaying ${fullyFilteredData.length} matching records for summary from all ${originalData.length} data rows.`; let currentStatus = statusElement.textContent; let warningPart = ""; if(currentStatus.includes("Warning:")) { warningPart = currentStatus.substring(currentStatus.indexOf("Warning:")); }
    if (fullyFilteredData.length === 0 && originalData.length > 0) { statusElement.textContent = "No data matches the current Subchannel/Territory/Store filter selection." + warningPart; }
    else if (originalData.length > 0) { statusElement.textContent = baseMessage + warningPart; }
    // Process and display
    processAndDisplayData(fullyFilteredData, subchannelFilteredData, isSingleStoreSelected, isAllStoresSelected);
}


// --- Helper Functions ---
function parsePercentage(value) { if (typeof value === 'number') { if (Math.abs(value) < 5) return value; else return value / 100; } if (typeof value === 'string') { const cleanedValue = value.replace('%', '').trim(); if (cleanedValue === '') return null; const number = parseFloat(cleanedValue); if (isNaN(number)) { return null; } return number / 100; } return null; }
function formatPercentage(decimalValue, digits = 1) { if (decimalValue === null || isNaN(decimalValue)) return "N/A"; return (decimalValue * 100).toFixed(digits) + '%'; }
function getAchievementHighlightColor(achievementDecimal, territoryBenchmarkDecimal) { if (achievementDecimal === null || territoryBenchmarkDecimal === null || territoryBenchmarkDecimal === 0) return 'inherit'; if (achievementDecimal >= territoryBenchmarkDecimal) return '#90EE90'; if (achievementDecimal >= territoryBenchmarkDecimal * 0.9) return '#FFFFE0'; if (achievementDecimal >= territoryBenchmarkDecimal * 0.8) return '#FFDDA0'; return '#FFCCCB'; }
function getGradientColor(value, min, max) { if (value === null || min === null || max === null || min === max) return 'inherit'; const position = Math.max(0, Math.min(1, (value - min) / (max - min))); let r, g, b; if (position < 0.5) { r = 255; g = Math.round(255 * (position * 2)); b = 0; } else { r = Math.round(255 * (1 - (position - 0.5) * 2)); g = 255; b = 0; } return `rgb(${r}, ${g}, ${b})`; }
function calculateColumnStats(data, columnName) { const values = data.map(row => parsePercentage(row[columnName])).filter(val => val !== null && !isNaN(val)); if (values.length === 0) { return { average: null, stdDev: null, count: 0 }; } const sum = values.reduce((acc, val) => acc + val, 0); const average = sum / values.length; let stdDev = null; if (values.length > 1) { const squareDiffs = values.map(val => Math.pow(val - average, 2)); const avgSquareDiff = squareDiffs.reduce((acc, val) => acc + val, 0) / values.length; stdDev = Math.sqrt(avgSquareDiff); } else { stdDev = 0; } return { average: average, stdDev: stdDev, count: values.length }; }


// --- Function: Process Data and Update Dashboard ---
function processAndDisplayData(fullyFilteredData, subchannelFilteredData, isSingleStoreSelected, isAllStoresSelected) {

    let totalRevenueUnfiltered = 0; let totalTargetUnfiltered = 0;
    originalData.forEach(row => { totalRevenueUnfiltered += (parseFloat(row['Revenue w/DF']) || 0); totalTargetUnfiltered += (parseFloat(row['QTD Revenue Target']) || 0); });
    const territoryRevARUnfilteredDecimal = totalTargetUnfiltered > 0 ? (totalRevenueUnfiltered / totalTargetUnfiltered) : null;

    let totalQtdTargetFiltered = 0, totalRevenueWithDFFiltered = 0, totalUnitsWithDFFiltered = 0, totalUnitTargetFiltered = 0;
    let sumRepSkill = 0, countRepSkill = 0, sumVPmr = 0, countVPmr = 0, hasRepPmrData = false;
    let sumElite = 0, countElite = 0, sumPostTrain = 0, countPostTrain = 0, hasNonZeroElite = false;
    let lineChartLabels = [], lineChartRevARData = [], lineChartUnitAchData = [];
    let attachRateDataForTable = [];
    let attachRateStats = {};
    let territoryValueForDisplay = null; let districtValueForDisplay = null; let postTrainValueForDisplay = null; let storeDetailsData = null;
    let showCombinedTerritoryRev = false; let showCombinedDistrictRev = false;

    // Define required columns
    const baseRequiredCols = ['Store', 'STORE ID', 'SUB_CHANNEL', 'Q2 Territory', 'Revenue w/DF', 'QTD Revenue Target', 'Territory Rev%', 'Unit w/ DF', 'Unit Target', 'Rev AR%', 'Unit Achievement', 'Rep Skill Ach', '(V)PMR Ach', 'Elite', 'Post Training Score'];
    const districtCols = ['DISTRICT', 'District Rev%'];
    const storeDetailCols = ['ORG_STORE_ID', 'STORE_NAME', 'ADDRESS1', 'CITY', 'STATE', 'ZIPCODE', 'PHONE_NO', 'Q1 TERRITORY', 'FSM NAME', 'DEALER_NAME', 'CHANNEL_TYPE', 'SUPER STORE', 'GOLDEN RHINO', 'LATITUDE_ORG', 'LONGITUDE_ORG', 'RETAIL_MAP_OPP_TIER_NAME', 'CINGLEPOINT_ID', 'CHANNEL', 'Hispanic_Market', 'GCE', 'AI_Zone'];
    const attachRateColumns = ['Tablet Attach Rate', 'PC Attach Rate', 'NC Attach Rate', 'TWS Attach Rate', 'WW Attach Rate', 'ME Attach Rate', 'NCME Attach Rate'];
    const allRequiredCols = [...new Set([...baseRequiredCols, ...districtCols, ...storeDetailCols, ...attachRateColumns])];

    // Check for Missing Columns
    let missingCols = [];
    const dataToCheckColumns = fullyFilteredData.length > 0 ? fullyFilteredData : (subchannelFilteredData.length > 0 ? subchannelFilteredData : originalData);
    if (dataToCheckColumns.length > 0) { allRequiredCols.forEach(col => { if (dataToCheckColumns[0][col] === undefined) missingCols.push(col); }); }
    else if (originalData.length > 0) { allRequiredCols.forEach(col => { if (originalData[0][col] === undefined) missingCols.push(col); }); }
    if (missingCols.length > 0) { console.warn("Required columns missing:", missingCols); if (!statusElement.textContent.includes('Warning:')) statusElement.textContent += ' Warning: '; statusElement.textContent += ` Missing columns (${missingCols.slice(0,3).join(', ')}${missingCols.length > 3 ? '...' : ''}). Some sections/calculations may fail.`; }

    // Process FULLY Filtered Data
    if (fullyFilteredData.length > 0) {
        const coreSummaryColsPresent = !missingCols.some(c => ['Revenue w/DF', 'QTD Revenue Target', 'Territory Rev%', 'District Rev%', 'Unit w/ DF', 'Unit Target'].includes(c));
        const conditionalStatsColsPresent = !missingCols.some(c => ['Rep Skill Ach', '(V)PMR Ach', 'Elite', 'Post Training Score'].includes(c));

        if (coreSummaryColsPresent) {
            fullyFilteredData.forEach(row => {
                totalQtdTargetFiltered += (parseFloat(row['QTD Revenue Target']) || 0); totalRevenueWithDFFiltered += (parseFloat(row['Revenue w/DF']) || 0); totalUnitsWithDFFiltered += (parseFloat(row['Unit w/ DF']) || 0); totalUnitTargetFiltered += (parseFloat(row['Unit Target']) || 0);
                if(conditionalStatsColsPresent) {
                    const repSkill = parsePercentage(row['Rep Skill Ach']); if (repSkill !== null) { sumRepSkill += repSkill; countRepSkill++; hasRepPmrData = true; }
                    const vPmr = parsePercentage(row['(V)PMR Ach']); if (vPmr !== null) { sumVPmr += vPmr; countVPmr++; hasRepPmrData = true; }
                    const elite = parsePercentage(row['Elite']); if (elite !== null) { sumElite += elite; countElite++; if (elite > 0.0001) hasNonZeroElite = true; }
                    const postTrainScore = parseFloat(row['Post Training Score']); if (!isNaN(postTrainScore)) { sumPostTrain += postTrainScore; countPostTrain++; }
                }
            });

            if (isSingleStoreSelected && fullyFilteredData.length === 1) {
                const singleRow = fullyFilteredData[0]; territoryValueForDisplay = parsePercentage(singleRow['Territory Rev%']); districtValueForDisplay = parsePercentage(singleRow['District Rev%']); showCombinedTerritoryRev = false; showCombinedDistrictRev = false;
                const coreDetailColsPresent = !missingCols.some(c => ['Store', 'Q2 Territory', 'ORG_STORE_ID', 'STORE_NAME', 'ADDRESS1', 'CITY', 'STATE', 'ZIPCODE'].includes(c));
                 if (coreDetailColsPresent) { storeDetailsData = {}; allRequiredCols.forEach(col => storeDetailsData[col] = singleRow[col]); }
                 else { storeDetailsData = null; console.warn("Cannot display Store Details due to missing columns."); }
            } else if (!isAllStoresSelected && fullyFilteredData.length > 1) {
                const firstTerritory = String(fullyFilteredData[0]['Q2 Territory'] || '').trim(); const allSameTerritory = fullyFilteredData.every(row => String(row['Q2 Territory'] || '').trim() === firstTerritory);
                const firstDistrict = String(fullyFilteredData[0]['DISTRICT'] || '').trim(); const allSameDistrict = fullyFilteredData.every(row => String(row['DISTRICT'] || '').trim() === firstDistrict);
                if (allSameTerritory && firstTerritory !== '') { let sumT = 0; let countT = 0; fullyFilteredData.forEach(row => { const val = parsePercentage(row['Territory Rev%']); if (val !== null) { sumT += val; countT++; } }); territoryValueForDisplay = countT > 0 ? sumT : null; showCombinedTerritoryRev = true; }
                else { territoryValueForDisplay = null; showCombinedTerritoryRev = false; }
                if (allSameDistrict && firstDistrict !== '') { let sumD = 0; let countD = 0; fullyFilteredData.forEach(row => { const val = parsePercentage(row['District Rev%']); if (val !== null) { sumD += val; countD++; } }); districtValueForDisplay = countD > 0 ? sumD : null; showCombinedDistrictRev = true; }
                else { districtValueForDisplay = null; showCombinedDistrictRev = false; }
                storeDetailsData = null;
            } else { territoryValueForDisplay = null; districtValueForDisplay = null; showCombinedTerritoryRev = false; showCombinedDistrictRev = false; storeDetailsData = null; }

             if (conditionalStatsColsPresent) {
                if (countPostTrain === 1 && fullyFilteredData.length === 1) { postTrainValueForDisplay = parseFloat(fullyFilteredData[0]['Post Training Score']); if(isNaN(postTrainValueForDisplay)) postTrainValueForDisplay = null; }
                else if (countPostTrain > 0) { postTrainValueForDisplay = sumPostTrain / countPostTrain; }
             }
        } else { console.warn("Core summary columns missing, cannot calculate summary stats."); territoryValueForDisplay = null; districtValueForDisplay = null; postTrainValueForDisplay = null; storeDetailsData = null; showCombinedTerritoryRev = false; showCombinedDistrictRev = false; }
    } else { territoryValueForDisplay = null; districtValueForDisplay = null; postTrainValueForDisplay = null; storeDetailsData = null; showCombinedTerritoryRev = false; showCombinedDistrictRev = false; countPostTrain = 0; }

    const overallRevenueAchievementFilteredDecimal = totalQtdTargetFiltered > 0 ? (totalRevenueWithDFFiltered / totalQtdTargetFiltered) : null;
    const overallUnitAchievementFilteredPercent = totalUnitTargetFiltered > 0 ? (totalUnitsWithDFFiltered / totalUnitTargetFiltered) * 100 : null;
    const avgRepSkillAch = countRepSkill > 0 ? sumRepSkill / countRepSkill : null;
    const avgVPmrAch = countVPmr > 0 ? sumVPmr / countVPmr : null;
    const avgElite = countElite > 0 ? sumElite / countElite : null;

    // Display Sections
    displayStoreDetails(storeDetailsData);
    displaySummary( totalQtdTargetFiltered, totalRevenueWithDFFiltered, overallRevenueAchievementFilteredDecimal, territoryRevARUnfilteredDecimal, isSingleStoreSelected, territoryValueForDisplay, showCombinedTerritoryRev, districtValueForDisplay, showCombinedDistrictRev, totalUnitsWithDFFiltered, totalUnitTargetFiltered, overallUnitAchievementFilteredPercent );
    displayRepPmrData(avgRepSkillAch, avgVPmrAch, hasRepPmrData);
    displayTrainingStats(avgElite, postTrainValueForDisplay, countPostTrain, hasNonZeroElite);

    // Process Data for Chart and Table
    const dataForDetails = isAllStoresSelected ? subchannelFilteredData : fullyFilteredData;
    const chartColsPresent = !missingCols.some(c => ['Rev AR%', 'Unit Achievement', 'Store', 'STORE ID'].includes(c));
    const tableColsPresent = !missingCols.some(c => attachRateColumns.includes(c) || ['Store', 'STORE ID'].includes(c));

    if (dataForDetails.length > 0) {
        if (tableColsPresent) { attachRateColumns.forEach(colName => { attachRateStats[colName] = calculateColumnStats(dataForDetails, colName); }); }
        else { attachRateColumns.forEach(colName => { attachRateStats[colName] = { average: null, stdDev: null, count: 0 }; }); }
        dataForDetails.forEach((row, index) => {
            const storeLabel = (row['Store'] && String(row['Store']).trim()) || (row['STORE ID'] && String(row['STORE ID']).trim()) || `Unknown ${index + 1}`;
            if (chartColsPresent) { lineChartLabels.push(storeLabel); lineChartRevARData.push(parsePercentage(row['Rev AR%'])); lineChartUnitAchData.push(parsePercentage(row['Unit Achievement'])); }
            if (tableColsPresent) { const tableRow = { store: storeLabel }; attachRateColumns.forEach(col => tableRow[col] = row[col]); attachRateDataForTable.push(tableRow); }
        });
    }

    // Display Chart and Table
    const chartData = (lineChartLabels.length > 0) ? { labels: lineChartLabels, datasets: [ { label: 'Rev AR %', data: lineChartRevARData, borderColor: 'rgb(75, 192, 192)', backgroundColor: 'rgba(75, 192, 192, 0.7)', yAxisID: 'yPercentage' }, { label: 'Unit Achievement %', data: lineChartUnitAchData, borderColor: 'rgb(255, 99, 132)', backgroundColor: 'rgba(255, 99, 132, 0.7)', yAxisID: 'yPercentage' } ] } : { labels: [], datasets: [] };
    displayLineChart(chartData, isAllStoresSelected);
    if(attachRateDataForTable.length > 0){ displayAttachRateTable(attachRateDataForTable, attachRateStats, isAllStoresSelected); }
    else { displayAttachRateTable([], {}, isAllStoresSelected); }
}


// --- Display Functions ---
function displayStoreDetails(details) {
    if (!storeDetailsElement) return;
    if (details) {
        let html = `<h3>Store Details</h3>`;
        const addDetail = (label, value) => { if (value !== null && value !== undefined && String(value).trim() !== '') { html += `<p><strong>${label}:</strong> ${String(value).trim()}</p>`; } };
        const addConditionalDetail = (label, value, conditionVal = 'yes') => { if (value !== null && value !== undefined && String(value).trim().toLowerCase() === conditionVal.toLowerCase()) { html += `<p><strong>${label}:</strong> Yes</p>`; } };
        addDetail('Store', details['Store']);
        addDetail('ORG Store ID', details['ORG_STORE_ID']);
        addDetail('Store Name', details['STORE_NAME']);
        addDetail('Address', details['ADDRESS1']);
        addDetail('City', details['CITY']);
        addDetail('State', details['STATE']);
        addDetail('Zipcode', details['ZIPCODE']);
        addDetail('Phone #', details['PHONE_NO']);
        const q1Terr = String(details['Q1 TERRITORY'] || '').trim(); const q2Terr = String(details['Q2 Territory'] || '').trim();
        if (q1Terr && q2Terr && q1Terr !== q2Terr) { addDetail('Previous Territory', q1Terr); }
        const fsmName = String(details['FSM NAME'] || '').trim(); const q2Display = fsmName ? `${q2Terr} (${fsmName})` : q2Terr; addDetail('Q2 Territory', q2Display);
        if (String(details['CHANNEL_TYPE'] || '').toLowerCase() === 'dealer') { addDetail('Dealer Name', details['DEALER_NAME']); }
        if (String(details['CHANNEL'] || '').toLowerCase() === 'at&t') { addDetail('CinglePoint ID', details['CINGLEPOINT_ID']); }
        addConditionalDetail('Super Store', details['SUPER STORE']); addConditionalDetail('Golden Rhino', details['GOLDEN RHINO']); addConditionalDetail('Hispanic Market', details['Hispanic_Market']); addConditionalDetail('GCE', details['GCE']); addConditionalDetail('AI Zone', details['AI_Zone']);
        addDetail('Latitude', details['LATITUDE_ORG']); addDetail('Longitude', details['LONGITUDE_ORG']); addDetail('Rank', details['RETAIL_MAP_OPP_TIER_NAME']);
        storeDetailsElement.innerHTML = html; storeDetailsElement.style.display = 'block'; summaryElement.style.marginTop = '0';
    } else { storeDetailsElement.innerHTML = ''; storeDetailsElement.style.display = 'none'; summaryElement.style.marginTop = '1rem'; }
}
function displaySummary( totalQtdTargetFiltered, totalRevenueWithDFFiltered, overallRevenueAchievementFilteredDecimal, territoryRevARUnfilteredDecimal, isSingleStoreSelected, territoryValueDecimal, showCombinedTerritoryRev, districtValueDecimal, showCombinedDistrictRev, totalUnitsWithDFFiltered, totalUnitTargetFiltered, overallUnitAchievementFilteredPercent ) { const displayRevenueAchPercent = overallRevenueAchievementFilteredDecimal !== null ? formatPercentage(overallRevenueAchievementFilteredDecimal, 2) : 'N/A'; const achievementBgColor = getAchievementHighlightColor(overallRevenueAchievementFilteredDecimal, territoryRevARUnfilteredDecimal); let territoryHtml = ''; if (isSingleStoreSelected) { territoryHtml = `<p><strong>Territory Rev %:</strong> ${formatPercentage(territoryValueDecimal, 2)}</p>`; } else if (showCombinedTerritoryRev) { territoryHtml = `<p><strong>Combined Territory Rev %:</strong> ${formatPercentage(territoryValueDecimal, 2)}</p>`; } let districtHtml = ''; if (isSingleStoreSelected) { districtHtml = `<p><strong>District Rev %:</strong> ${formatPercentage(districtValueDecimal, 2)}</p>`; } else if (showCombinedDistrictRev) { districtHtml = `<p><strong>Combined District Rev %:</strong> ${formatPercentage(districtValueDecimal, 2)}</p>`; } const displayUnitAchPercent = overallUnitAchievementFilteredPercent !== null ? overallUnitAchievementFilteredPercent.toFixed(2) + '%' : 'N/A'; summaryElement.innerHTML = ` <h3>Filtered QTD Summary</h3> <p><strong>Total QTD Revenue Target:</strong> ${formatCurrency(totalQtdTargetFiltered)}</p> <p><strong>Total Revenue w/DF (QTD):</strong> ${formatCurrency(totalRevenueWithDFFiltered)}</p> <p><strong>Overall QTD Revenue Achievement:</strong> <span style="background-color: ${achievementBgColor}; padding: 2px 5px; border-radius: 3px;">${displayRevenueAchPercent}</span> ${territoryRevARUnfilteredDecimal !== null ? ` (Territory Benchmark: ${formatPercentage(territoryRevARUnfilteredDecimal, 2)})` : ''}</p> ${territoryHtml} ${districtHtml} <hr style="border-top: 1px solid #444; margin: 10px 0;"> <p><strong>Total Units w/ DF:</strong> ${totalUnitsWithDFFiltered.toLocaleString()}</p> <p><strong>Total Unit Target:</strong> ${totalUnitTargetFiltered.toLocaleString()}</p> <p><strong>Overall Unit Achievement:</strong> ${displayUnitAchPercent}</p> `; const hasSummaryData = overallRevenueAchievementFilteredDecimal !== null || territoryValueDecimal !== null || districtValueDecimal !== null || overallUnitAchievementFilteredPercent !== null || totalUnitsWithDFFiltered > 0 || totalUnitTargetFiltered > 0; summaryElement.style.textAlign = hasSummaryData ? 'left' : 'center'; if (!hasSummaryData && originalData.length > 0 && fullyFilteredData.length > 0) { summaryElement.innerHTML += '<p style="font-style: italic; text-align: center;">N/A or missing data for summary calculations based on current filters.</p>'; } else if (fullyFilteredData.length === 0 && originalData.length > 0) { summaryElement.innerHTML = '<p style="font-style: italic; text-align: center;">No data matches filters for summary.</p>'; } else if (originalData.length === 0) { summaryElement.innerHTML = 'Summary data will appear here...'; } }
function displayRepPmrData(avgRepSkillAch, avgVPmrAch, hasData) { if (hasData) { repSkillAchValue.textContent = formatPercentage(avgRepSkillAch, 1); vPmrAchValue.textContent = formatPercentage(avgVPmrAch, 1); repPmrSection.style.display = 'block'; } else { repPmrSection.style.display = 'none'; } }
function displayTrainingStats(avgElite, postTrainScore, postTrainCount, hasNonZeroElite) { let showSection = false; if (avgElite !== null && hasNonZeroElite) { eliteValue.textContent = formatPercentage(avgElite, 1); eliteP.style.display = 'block'; showSection = true; } else { eliteP.style.display = 'none'; } if (postTrainScore !== null) { const avgLabel = postTrainCount > 1 ? " (Avg)" : ""; postTrainingScoreValue.textContent = postTrainScore.toFixed(1) + avgLabel; postTrainingP.style.display = 'block'; showSection = true; } else { postTrainingScoreValue.textContent = 'N/A'; postTrainingP.style.display = 'block'; } if (showSection || postTrainScore !== null) { trainingStatsSection.style.display = 'block'; } else { trainingStatsSection.style.display = 'none'; } }
function displayLineChart(chartData, isAllStoresSelected) { if (lineChartInstance) { lineChartInstance.destroy(); } const isEmpty = !chartData || !chartData.labels || chartData.labels.length === 0; const gridColor = 'rgba(255, 255, 255, 0.1)'; const textColor = '#e0e0e0'; const titleColor = '#f0f0f0'; const chartTitleText = isEmpty ? 'Bar Chart - Select Filters / No Data' : `Rev AR % vs Unit Achievement % per Store (${isAllStoresSelected ? 'Subchannel Filtered' : 'Selected Stores'})`; lineChartInstance = new Chart(lineChartCanvas, { type: 'bar', data: chartData, options: { responsive: true, maintainAspectRatio: false, plugins: { title: { display: true, text: chartTitleText, color: titleColor }, tooltip: { callbacks: { label: function(context) { let label = context.dataset.label || ''; if (label) { label += ': '; } if (context.formattedValue !== null) { label += formatPercentage(context.parsed.y); } else { label += 'N/A'; } return label; } } }, legend: { labels: { color: textColor } } }, interaction: { intersect: false, mode: 'index', }, scales: { x: { title: { display: true, text: 'Store', color: textColor }, ticks: { color: textColor }, grid: { color: gridColor } }, yPercentage: { type: 'linear', position: 'left', beginAtZero: true, title: { display: true, text: 'Percentage (%)', color: textColor }, ticks: { callback: function(value, index, values) { return formatPercentage(value, 0); }, color: textColor }, grid: { color: gridColor } } } } }); }
function displayAttachRateTable(tableData, columnStats, isAllStoresSelected) { attachRateTableBody.innerHTML = ''; attachTableStatusElement.textContent = ''; if (attachRateTableHeading) { attachRateTableHeading.textContent = `Attach Rates (${isAllStoresSelected ? 'Subchannel Filtered' : 'Selected Stores'})`; } if (!tableData || tableData.length === 0) { attachTableStatusElement.textContent = `No data for ${isAllStoresSelected ? 'selected subchannel' : 'selected stores'} or required columns missing.`; return; } if (!columnStats) { console.error("Attach Rate Stats missing!"); attachTableStatusElement.textContent = 'Error calculating stats for highlighting.'; columnStats = {}; } tableData.forEach(rowData => { const row = attachRateTableBody.insertRow(); const addCell = (text) => { const cell = row.insertCell(); cell.textContent = text; return cell; }; const addPercentageCellWithHighlight = (rawValue, colName) => { const cell = row.insertCell(); const decimalValue = parsePercentage(rawValue); cell.textContent = formatPercentage(decimalValue); cell.style.textAlign = 'right'; const stats = columnStats[colName] || { average: null, stdDev: null }; if (decimalValue !== null && stats.average !== null && stats.stdDev !== null && stats.stdDev > 0.0001) { const lowerBound = stats.average - stats.stdDev; const upperBound = stats.average + stats.stdDev; if (decimalValue > upperBound) { cell.classList.add('highlight-green'); } else if (decimalValue < lowerBound) { cell.classList.add('highlight-red'); } } return cell; }; addCell(rowData.store); addPercentageCellWithHighlight(rowData['Tablet Attach Rate'], 'Tablet Attach Rate'); addPercentageCellWithHighlight(rowData['PC Attach Rate'], 'PC Attach Rate'); addPercentageCellWithHighlight(rowData['NC Attach Rate'], 'NC Attach Rate'); addPercentageCellWithHighlight(rowData['TWS Attach Rate'], 'TWS Attach Rate'); addPercentageCellWithHighlight(rowData['WW Attach Rate'], 'WW Attach Rate'); addPercentageCellWithHighlight(rowData['ME Attach Rate'], 'ME Attach Rate'); addPercentageCellWithHighlight(rowData['NCME Attach Rate'], 'NCME Attach Rate'); }); attachTableStatusElement.textContent = `Displaying ${tableData.length} stores. Highlight based on stats from ${isAllStoresSelected ? 'subchannel' : 'selected stores'} (> avg + 1 std dev (green), < avg - 1 std dev (red)).`; }

// --- Email Share Handler ---
function handleShareEmail() { const recipient = emailRecipientInput.value.trim(); shareStatusElement.textContent = ''; if (!recipient) { shareStatusElement.textContent = 'Please enter a recipient email address.'; emailRecipientInput.focus(); return; } if (!recipient.includes('@') || !recipient.includes('.')) { shareStatusElement.textContent = 'Please enter a valid email address format.'; emailRecipientInput.focus(); return; } let emailBody = "Dashboard Summary:\n"; emailBody += "--------------------\n"; const summaryParas = summaryElement.querySelectorAll('p'); if (summaryParas.length > 0 && !summaryElement.textContent.includes('No matching data')) { summaryParas.forEach(p => { let text = p.textContent.replace(/\s+/g, ' ').trim(); text = text.replace(' (Filtered)', ''); if (text.includes('Territory Benchmark:')) { emailBody += text + "\n"; } else if (text.startsWith('Combined Territory Rev %') || text.startsWith('Territory Rev %') || text.startsWith('Combined District Rev %') || text.startsWith('District Rev %')) { emailBody += text + "\n"; } else if (text.startsWith('Total Units w/ DF')) { emailBody += "\n" + text + "\n"; } else if (!text.startsWith('Average Territory Rev %')) { emailBody += text + "\n"; } }); } else { emailBody += "(No summary data available for current filters)\n"; } emailBody += "\n"; let storeDetailsText = ''; if (storeDetailsElement && storeDetailsElement.style.display !== 'none') { storeDetailsText += "\nStore Details:\n"; storeDetailsText += "--------------------\n"; const detailParas = storeDetailsElement.querySelectorAll('p'); detailParas.forEach(p => { storeDetailsText += p.textContent.trim() + "\n"; }); emailBody += storeDetailsText + "\n"; } if (repPmrSection.style.display !== 'none') { emailBody += "Mysteryshop Results:\n"; emailBody += "--------------------\n"; emailBody += repSkillP.textContent.trim() + "\n"; emailBody += vPmrP.textContent.trim() + "\n\n"; } if (trainingStatsSection.style.display !== 'none') { emailBody += "Training Stats:\n"; emailBody += "--------------------\n"; if (eliteP.style.display !== 'none') { emailBody += eliteP.textContent.trim() + "\n"; } if (postTrainingP.style.display !== 'none') { emailBody += postTrainingP.textContent.trim() + "\n"; } emailBody += "\n"; } emailBody += "--\n"; emailBody += "Christopher Wasney\n"; emailBody += "Samsung Field Sales Manager\n"; emailBody += "(989) 751-9277\n"; emailBody += "c9.wasney@sea.samsung.com\n"; const subject = "Dashboard Results Summary"; const encodedSubject = encodeURIComponent(subject); const encodedBody = encodeURIComponent(emailBody); const mailtoLink = `mailto:${recipient}?subject=${encodedSubject}&body=${encodedBody}`; if (mailtoLink.length > 1800) { console.warn("Generated mailto link is very long:", mailtoLink.length, "characters"); shareStatusElement.textContent = 'Warning: Results data might be too long for email; content may be truncated.'; } else { shareStatusElement.textContent = 'Opening your email client...'; } const mailtoWindow = window.open(mailtoLink, '_blank'); if (!mailtoWindow) { console.log("window.open failed, trying location.href for mailto"); window.location.href = mailtoLink; } setTimeout(() => { shareStatusElement.textContent = ''; }, 5000); }

// --- Clear Functions ---
function clearChart() { if (lineChartInstance) { lineChartInstance.destroy(); lineChartInstance = null; } }
function clearDashboard() { clearChart(); if (storeDetailsElement) { storeDetailsElement.innerHTML = ''; storeDetailsElement.style.display = 'none'; } summaryElement.innerHTML = 'Summary data will appear here...'; summaryElement.style.textAlign = 'center'; summaryElement.style.marginTop = '1rem'; statusElement.textContent = 'No file selected.'; territoryFilter.innerHTML = '<option value="ALL">-- Load File First --</option>'; resetMultiSelect(territoryFilter, 'ALL'); territoryFilter.disabled = true; subchannelFilter.innerHTML = '<option value="ALL">-- Load File First --</option>'; subchannelFilter.disabled = true; subchannelFilter.value = 'ALL'; storeFilter.innerHTML = '<option value="ALL">-- Load File First --</option>'; resetMultiSelect(storeFilter, 'ALL'); storeFilter.disabled = true; originalData = []; attachRateTableBody.innerHTML = ''; attachTableStatusElement.textContent = ''; if(attachRateTableHeading) attachRateTableHeading.textContent = 'Attach Rates'; if(repPmrSection) repPmrSection.style.display = 'none'; if(trainingStatsSection) trainingStatsSection.style.display = 'none'; if(eliteP) eliteP.style.display = 'none'; if(eliteValue) eliteValue.textContent = 'N/A'; if(postTrainingScoreValue) postTrainingScoreValue.textContent = 'N/A'; if(repSkillAchValue) repSkillAchValue.textContent = 'N/A'; if(vPmrAchValue) vPmrAchValue.textContent = 'N/A'; if(emailRecipientInput) emailRecipientInput.value = ''; if(shareStatusElement) shareStatusElement.textContent = ''; if(fileInput) fileInput.value = ''; }

// --- Helper Function to Format Currency ---
function formatCurrency(number) { if (isNaN(number) || number === null) return "$0.00"; const num = parseFloat(number); if (isNaN(num)) return "$0.00"; return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(num); }

// --- PWA Service Worker Registration ---
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    // *** UPDATED PATH back to root ***
    // Scope option is removed as the default '/' is now correct
    navigator.serviceWorker.register('/sw.js')
      .then(registration => {
        console.log('Service Worker registered successfully with scope: ', registration.scope);
      })
      .catch(error => {
        console.log('Service Worker registration failed: ', error);
      });
  });
} else {
    console.log('Service Worker is not supported by this browser.');
}
