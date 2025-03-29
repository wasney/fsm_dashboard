// --- Global Variables ---
let originalData = [];
let lineChartInstance = null;

// --- HTML Element References ---
const fileInput = document.getElementById('excelFile');
const statusElement = document.getElementById('status');
const summaryElement = document.getElementById('summaryData');
const lineChartCanvas = document.getElementById('lineChartCanvas').getContext('2d');
const attachRateTableBody = document.getElementById('attachRateTableBody');
const attachTableStatusElement = document.getElementById('attachTableStatus');
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
subchannelFilter.addEventListener('change', handleSubchannelChange);
storeFilter.addEventListener('change', filterAndDisplayData);
shareEmailButton.addEventListener('click', handleShareEmail);


// --- Main File Handling Logic --- (No changes)
function handleFileSelect(event) { const file = event.target.files[0]; if (!file) { statusElement.textContent = 'No file selected.'; clearDashboard(); return; } statusElement.textContent = `Processing file: ${file.name}...`; clearDashboard(); const reader = new FileReader(); reader.onload = function(e) { try { const data = e.target.result; const workbook = XLSX.read(data, { type: 'binary' }); const firstSheetName = workbook.SheetNames[0]; const worksheet = workbook.Sheets[firstSheetName]; const allJsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" }); const jsonData = allJsonData.slice(0, 61); let rowLimitMessage = ` Kept first ${jsonData.length} data rows (Excel rows 2-${jsonData.length + 1}, max 61 data rows).`; if (allJsonData.length < 61) { rowLimitMessage = ` Processed all ${jsonData.length} data rows (Excel rows 2-${jsonData.length + 1}).`; } if (!jsonData || jsonData.length === 0) { if (allJsonData.length > 0) { statusElement.textContent = `Error: No data found in the first 61 data rows (Excel rows 2-62) of the sheet.`; } else { statusElement.textContent = 'Error: No data found in the first sheet or file is empty.'; } return; } originalData = jsonData; populateFiltersOnInit(originalData); filterAndDisplayData(); statusElement.textContent = `Successfully processed ${file.name}.${rowLimitMessage} Displaying data. Use filters to refine.`; } catch (error) { console.error("Error processing Excel file:", error); statusElement.textContent = `Error processing file: ${error.message}. Check console for details.`; clearDashboard(); } }; reader.onerror = function(ex) { console.error("FileReader error:", ex); statusElement.textContent = 'Error reading file.'; clearDashboard(); }; reader.readAsBinaryString(file); }

// --- Filter Population Logic --- (No changes)
function populateFiltersOnInit(data) { const subchannels = new Set(); const stores = new Set(); let columnsPresent = true; if (data.length === 0 || !data[0] || data[0]['SUBCHANNEL'] === undefined || data[0]['Store'] === undefined) { if(data.length > 0) { statusElement.textContent += ` Warning: Could not fully populate filters. 'SUBCHANNEL' or 'Store' column might be missing in first 61 data rows.`; } columnsPresent = false; } if (columnsPresent) { data.forEach(row => { if (row['SUBCHANNEL'] && String(row['SUBCHANNEL']).trim()) subchannels.add(String(row['SUBCHANNEL']).trim()); if (row['Store'] && String(row['Store']).trim()) stores.add(String(row['Store']).trim()); }); } updateDropdownOptions(subchannelFilter, [...subchannels].sort(), 'Subchannel'); updateDropdownOptions(storeFilter, [...stores].sort(), 'Store'); resetMultiSelect(storeFilter, 'ALL'); subchannelFilter.disabled = !columnsPresent || subchannels.size === 0; storeFilter.disabled = !columnsPresent || stores.size === 0; }
function updateDropdownOptions(selectElement, optionsArray, filterName) { const isMultiple = selectElement.multiple; selectElement.innerHTML = ''; const allOptionValue = 'ALL'; const allOptionText = `-- All ${filterName}s --`; selectElement.add(new Option(allOptionText, allOptionValue)); optionsArray.forEach(optionValue => { if (optionValue !== allOptionValue) { selectElement.add(new Option(optionValue, optionValue)); } }); if (!isMultiple) { selectElement.value = allOptionValue; } }
function updateStoreFilterOptions() { const selectedSubchannel = subchannelFilter.value; const relevantStores = new Set(); let dataToScan = originalData; if (selectedSubchannel !== 'ALL') { dataToScan = originalData.filter(row => String(row['SUBCHANNEL'] || '').trim() === selectedSubchannel); } if (dataToScan.length > 0 && dataToScan[0]['Store'] !== undefined) { dataToScan.forEach(row => { if (row['Store'] && String(row['Store']).trim()) relevantStores.add(String(row['Store']).trim()); }); } else if (originalData.length > 0 && (!originalData[0] || originalData[0]['Store'] === undefined)) { if (!statusElement.textContent.includes("Warning: 'Store' column")) statusElement.textContent += ` Warning: 'Store' column may be missing, Store filter cannot be updated.`; } updateDropdownOptions(storeFilter, [...relevantStores].sort(), 'Store'); resetMultiSelect(storeFilter, 'ALL'); }

// --- Helper Function to Reset Multi-Select --- (No changes)
function resetMultiSelect(selectElement, valueToSelect) { const valuesToSelect = Array.isArray(valueToSelect) ? valueToSelect : [valueToSelect]; for (let i = 0; i < selectElement.options.length; i++) { selectElement.options[i].selected = valuesToSelect.includes(selectElement.options[i].value); } }

// --- Event Handlers --- (No changes)
function handleSubchannelChange() { updateStoreFilterOptions(); filterAndDisplayData(); }

// --- Data Filtering and Display Trigger --- (No changes)
function filterAndDisplayData() { const selectedSubchannel = subchannelFilter.value; const selectedStoreOptions = Array.from(storeFilter.selectedOptions); let selectedStores = selectedStoreOptions.map(option => option.value); const allStoresOption = storeFilter.querySelector('option[value="ALL"]'); const containsAll = selectedStores.includes('ALL'); const containsIndividual = selectedStores.some(val => val !== 'ALL'); if (containsAll && containsIndividual) { if (allStoresOption) allStoresOption.selected = false; selectedStores = selectedStores.filter(val => val !== 'ALL'); } else if (!containsAll && selectedStores.length === 0 && allStoresOption) { allStoresOption.selected = true; selectedStores = ['ALL']; } else if (containsAll && !containsIndividual && selectedStores.length === 1 && allStoresOption) { if (!allStoresOption.selected) allStoresOption.selected = true; for(let i = 0; i < storeFilter.options.length; i++) { if (storeFilter.options[i].value !== 'ALL') storeFilter.options[i].selected = false; } } const isAllStoresSelected = selectedStores.includes('ALL'); const isSingleStoreSelected = !isAllStoresSelected && selectedStores.length === 1; let fullyFilteredData = originalData; if (selectedSubchannel !== 'ALL') { fullyFilteredData = fullyFilteredData.filter(row => String(row['SUBCHANNEL'] || '').trim() === selectedSubchannel); } if (!isAllStoresSelected) { fullyFilteredData = fullyFilteredData.filter(row => selectedStores.includes(String(row['Store'] || '').trim())); } let subchannelFilteredData = originalData; if (selectedSubchannel !== 'ALL') { subchannelFilteredData = subchannelFilteredData.filter(row => String(row['SUBCHANNEL'] || '').trim() === selectedSubchannel); } const baseMessage = `Displaying ${fullyFilteredData.length} matching records for summary from first ${originalData.length} data rows.`; let currentStatus = statusElement.textContent; let warningPart = ""; if(currentStatus.includes("Warning:")) { warningPart = currentStatus.substring(currentStatus.indexOf("Warning:")); } if (fullyFilteredData.length === 0 && originalData.length > 0) { statusElement.textContent = "No data matches the current Subchannel/Store filter selection." + warningPart; } else if (originalData.length > 0) { statusElement.textContent = baseMessage + warningPart; } processAndDisplayData(fullyFilteredData, subchannelFilteredData, isSingleStoreSelected); }


// --- Helper Functions --- (No changes)
function parsePercentage(value) { if (typeof value === 'number') { if (Math.abs(value) < 5) return value; else return value / 100; } if (typeof value === 'string') { const cleanedValue = value.replace('%', '').trim(); if (cleanedValue === '') return null; const number = parseFloat(cleanedValue); if (isNaN(number)) { return null; } return number / 100; } return null; }
function formatPercentage(decimalValue, digits = 1) { if (decimalValue === null || isNaN(decimalValue)) return "N/A"; return (decimalValue * 100).toFixed(digits) + '%'; }
function getAchievementHighlightColor(achievementDecimal, territoryBenchmarkDecimal) { if (achievementDecimal === null || territoryBenchmarkDecimal === null || territoryBenchmarkDecimal === 0) return 'inherit'; if (achievementDecimal >= territoryBenchmarkDecimal) return '#90EE90'; if (achievementDecimal >= territoryBenchmarkDecimal * 0.9) return '#FFFFE0'; if (achievementDecimal >= territoryBenchmarkDecimal * 0.8) return '#FFDDA0'; return '#FFCCCB'; }
function getGradientColor(value, min, max) { if (value === null || min === null || max === null || min === max) return 'inherit'; const position = Math.max(0, Math.min(1, (value - min) / (max - min))); let r, g, b; if (position < 0.5) { r = 255; g = Math.round(255 * (position * 2)); b = 0; } else { r = Math.round(255 * (1 - (position - 0.5) * 2)); g = 255; b = 0; } return `rgb(${r}, ${g}, ${b})`; }
function calculateColumnStats(data, columnName) { const values = data.map(row => parsePercentage(row[columnName])).filter(val => val !== null && !isNaN(val)); if (values.length === 0) { return { average: null, stdDev: null, count: 0 }; } const sum = values.reduce((acc, val) => acc + val, 0); const average = sum / values.length; let stdDev = null; if (values.length > 1) { const squareDiffs = values.map(val => Math.pow(val - average, 2)); const avgSquareDiff = squareDiffs.reduce((acc, val) => acc + val, 0) / values.length; stdDev = Math.sqrt(avgSquareDiff); } else { stdDev = 0; } return { average: average, stdDev: stdDev, count: values.length }; }


// --- Function: Process Data and Update Dashboard --- (No changes)
function processAndDisplayData(fullyFilteredData, subchannelFilteredData, isSingleStoreSelected) { let totalRevenueUnfiltered = 0; let totalTargetUnfiltered = 0; originalData.forEach(row => { totalRevenueUnfiltered += (parseFloat(row['Revenue w/DF']) || 0); totalTargetUnfiltered += (parseFloat(row['QTD Revenue Target']) || 0); }); const territoryRevARUnfilteredDecimal = totalTargetUnfiltered > 0 ? (totalRevenueUnfiltered / totalTargetUnfiltered) : null; let totalQtdTargetFiltered = 0, totalRevenueWithDFFiltered = 0, totalUnitsWithDFFiltered = 0, totalUnitTargetFiltered = 0; let sumRepSkill = 0, countRepSkill = 0, sumVPmr = 0, countVPmr = 0, hasRepPmrData = false; let sumElite = 0, countElite = 0, sumPostTrain = 0, countPostTrain = 0, hasNonZeroElite = false; let lineChartLabels = [], lineChartRevARData = [], lineChartUnitAchData = []; let attachRateDataForTable = []; let attachRateStats = {}; let territoryValueForDisplay = null; const attachRateColumns = ['Tablet Attach Rate', 'PC Attach Rate', 'NC Attach Rate', 'TWS Attach Rate', 'WW Attach Rate', 'ME Attach Rate', 'NCME Attach Rate']; const summaryRequiredCols = ['Revenue w/DF', 'QTD Revenue Target', 'Territory Rev%', 'Unit w/ DF', 'Unit Target', 'Rep Skill Ach', '(V)PMR Ach', 'Elite', 'Post Training Score']; const lineChartRequiredCols = ['Rev AR%', 'Unit Achievement', 'Store', 'STORE ID']; const tableRequiredCols = [...attachRateColumns, 'Store', 'STORE ID']; const dataToCheckColumns = fullyFilteredData.length > 0 ? fullyFilteredData : (subchannelFilteredData.length > 0 ? subchannelFilteredData : originalData); let missingSummaryCols = [], missingLineCols = [], missingTableCols = []; if (dataToCheckColumns.length > 0) { summaryRequiredCols.forEach(col => { if (dataToCheckColumns[0][col] === undefined) missingSummaryCols.push(col); }); lineChartRequiredCols.forEach(col => { if (dataToCheckColumns[0][col] === undefined) missingLineCols.push(col); }); tableRequiredCols.forEach(col => { if (dataToCheckColumns[0][col] === undefined) missingTableCols.push(col); }); } if (missingSummaryCols.length > 0) console.warn("Missing columns for Summary/Stats:", missingSummaryCols); if (missingLineCols.length > 0) console.warn("Missing columns for Line Chart:", missingLineCols); if (missingTableCols.length > 0) console.warn("Missing columns for Attach Rate Table:", missingTableCols); if (fullyFilteredData.length > 0 && missingSummaryCols.length === 0) { fullyFilteredData.forEach(row => { totalQtdTargetFiltered += (parseFloat(row['QTD Revenue Target']) || 0); totalRevenueWithDFFiltered += (parseFloat(row['Revenue w/DF']) || 0); totalUnitsWithDFFiltered += (parseFloat(row['Unit w/ DF']) || 0); totalUnitTargetFiltered += (parseFloat(row['Unit Target']) || 0); const repSkill = parsePercentage(row['Rep Skill Ach']); if (repSkill !== null) { sumRepSkill += repSkill; countRepSkill++; hasRepPmrData = true; } const vPmr = parsePercentage(row['(V)PMR Ach']); if (vPmr !== null) { sumVPmr += vPmr; countVPmr++; hasRepPmrData = true; } const elite = parsePercentage(row['Elite']); if (elite !== null) { sumElite += elite; countElite++; if (elite > 0.0001) hasNonZeroElite = true; } const postTrainScore = parseFloat(row['Post Training Score']); if (!isNaN(postTrainScore)) { sumPostTrain += postTrainScore; countPostTrain++; } }); if (isSingleStoreSelected && fullyFilteredData.length === 1) { territoryValueForDisplay = parsePercentage(fullyFilteredData[0]['Territory Rev%']); } else { let sumOfTerritoryRevPercentages = 0; let countValidTerritoryRev = 0; fullyFilteredData.forEach(row => { const terrRev = parsePercentage(row['Territory Rev%']); if (terrRev !== null) { sumOfTerritoryRevPercentages += terrRev; countValidTerritoryRev++; } }); territoryValueForDisplay = countValidTerritoryRev > 0 ? sumOfTerritoryRevPercentages : null; isSingleStoreSelected = false; } } const overallRevenueAchievementFilteredDecimal = totalQtdTargetFiltered > 0 ? (totalRevenueWithDFFiltered / totalQtdTargetFiltered) : null; const overallUnitAchievementFilteredPercent = totalUnitTargetFiltered > 0 ? (totalUnitsWithDFFiltered / totalUnitTargetFiltered) * 100 : null; const avgRepSkillAch = countRepSkill > 0 ? sumRepSkill / countRepSkill : null; const avgVPmrAch = countVPmr > 0 ? sumVPmr / countVPmr : null; const avgElite = countElite > 0 ? sumElite / countElite : null; const avgPostTrainScore = countPostTrain > 0 ? sumPostTrain / countPostTrain : null; displaySummary( totalQtdTargetFiltered, totalRevenueWithDFFiltered, overallRevenueAchievementFilteredDecimal, territoryRevARUnfilteredDecimal, isSingleStoreSelected, territoryValueForDisplay, totalUnitsWithDFFiltered, totalUnitTargetFiltered, overallUnitAchievementFilteredPercent ); displayRepPmrData(avgRepSkillAch, avgVPmrAch, hasRepPmrData); displayTrainingStats(avgElite, avgPostTrainScore, hasNonZeroElite); if (subchannelFilteredData.length > 0) { if (missingTableCols.length === 0) { attachRateColumns.forEach(colName => { attachRateStats[colName] = calculateColumnStats(subchannelFilteredData, colName); }); } else { attachRateColumns.forEach(colName => { attachRateStats[colName] = { average: null, stdDev: null, count: 0 }; }); } subchannelFilteredData.forEach((row, index) => { const storeLabel = (row['Store'] && String(row['Store']).trim()) || (row['STORE ID'] && String(row['STORE ID']).trim()) || `Unknown ${index + 1}`; if(missingLineCols.length === 0) { lineChartLabels.push(storeLabel); lineChartRevARData.push(parsePercentage(row['Rev AR%'])); lineChartUnitAchData.push(parsePercentage(row['Unit Achievement'])); } if(missingTableCols.length === 0) { const tableRow = { store: storeLabel }; attachRateColumns.forEach(col => tableRow[col] = row[col]); attachRateDataForTable.push(tableRow); } }); } const lineChartData = (missingLineCols.length === 0 && lineChartLabels.length > 0) ? { labels: lineChartLabels, datasets: [ { label: 'Rev AR %', data: lineChartRevARData, borderColor: 'rgb(75, 192, 192)', backgroundColor: 'rgba(75, 192, 192, 0.7)', /* Increased opacity */ yAxisID: 'yPercentage' }, { label: 'Unit Achievement %', data: lineChartUnitAchData, borderColor: 'rgb(255, 99, 132)', backgroundColor: 'rgba(255, 99, 132, 0.7)', /* Increased opacity */ yAxisID: 'yPercentage' } ] } : { labels: [], datasets: [] }; displayLineChart(lineChartData); if(missingTableCols.length === 0 && attachRateDataForTable.length > 0){ displayAttachRateTable(attachRateDataForTable, attachRateStats); } else { displayAttachRateTable([], {}); } }


// --- Display Functions ---
function displaySummary( totalQtdTargetFiltered, totalRevenueWithDFFiltered, overallRevenueAchievementFilteredDecimal, territoryRevARUnfilteredDecimal, isSingleStoreSelected, territoryValueDecimal, totalUnitsWithDFFiltered, totalUnitTargetFiltered, overallUnitAchievementFilteredPercent ) { const displayRevenueAchPercent = overallRevenueAchievementFilteredDecimal !== null ? formatPercentage(overallRevenueAchievementFilteredDecimal, 2) : 'N/A'; const achievementBgColor = getAchievementHighlightColor(overallRevenueAchievementFilteredDecimal, territoryRevARUnfilteredDecimal); const territoryLabel = isSingleStoreSelected ? "Territory Rev %" : "Combined Territory Rev %"; const displayTerritoryValue = formatPercentage(territoryValueDecimal, 2); const displayUnitAchPercent = overallUnitAchievementFilteredPercent !== null ? overallUnitAchievementFilteredPercent.toFixed(2) + '%' : 'N/A'; summaryElement.innerHTML = ` <h3>Filtered QTD Summary</h3> <p>Total QTD Revenue Target: ${formatCurrency(totalQtdTargetFiltered)}</p> <p>Total Revenue w/DF (QTD): ${formatCurrency(totalRevenueWithDFFiltered)}</p> <p>Overall QTD Revenue Achievement: <span style="background-color: ${achievementBgColor}; padding: 2px 5px; border-radius: 3px;">${displayRevenueAchPercent}</span> ${territoryRevARUnfilteredDecimal !== null ? ` (Territory Rev AR%: ${formatPercentage(territoryRevARUnfilteredDecimal, 2)})` : ''}</p> <p>${territoryLabel}: ${displayTerritoryValue}</p> <hr style="border-top: 1px dashed #ccc; margin: 10px 0;"> <p>Total Units w/ DF: ${totalUnitsWithDFFiltered.toLocaleString()}</p> <p>Total Unit Target: ${totalUnitTargetFiltered.toLocaleString()}</p> <p>Overall Unit Achievement: ${displayUnitAchPercent}</p> `; const hasData = overallRevenueAchievementFilteredDecimal !== null || territoryValueDecimal !== null || overallUnitAchievementFilteredPercent !== null || totalUnitsWithDFFiltered > 0 || totalUnitTargetFiltered > 0; summaryElement.style.textAlign = hasData ? 'left' : 'center'; if (!hasData && originalData.length > 0) { summaryElement.innerHTML += '<p style="font-style: italic; text-align: center;">No matching data for summary based on current filters and first 61 data rows.</p>'; } else if (originalData.length === 0) { summaryElement.innerHTML = 'Summary data will appear here...'; } }
function displayRepPmrData(avgRepSkillAch, avgVPmrAch, hasData) { if (hasData) { repSkillAchValue.textContent = formatPercentage(avgRepSkillAch, 1); vPmrAchValue.textContent = formatPercentage(avgVPmrAch, 1); repPmrSection.style.display = 'block'; } else { repPmrSection.style.display = 'none'; } }
function displayTrainingStats(avgElite, avgPostTrainScore, hasNonZeroElite) { let showSection = false; if (avgElite !== null && hasNonZeroElite) { eliteValue.textContent = formatPercentage(avgElite, 1); eliteP.style.display = 'block'; showSection = true; } else { eliteP.style.display = 'none'; } if (avgPostTrainScore !== null) { postTrainingScoreValue.textContent = avgPostTrainScore.toFixed(1); postTrainingP.style.display = 'block'; showSection = true; } else { postTrainingScoreValue.textContent = 'N/A'; } if (showSection) { trainingStatsSection.style.display = 'block'; } else { trainingStatsSection.style.display = 'none'; } }

// *** UPDATED displayLineChart function ***
function displayLineChart(chartData) {
    if (lineChartInstance) {
        lineChartInstance.destroy();
    }
    const isEmpty = !chartData || !chartData.labels || chartData.labels.length === 0;

    // Define colors for dark mode chart elements
    const gridColor = 'rgba(255, 255, 255, 0.1)';
    const textColor = '#e0e0e0';
    const titleColor = '#f0f0f0';

    lineChartInstance = new Chart(lineChartCanvas, {
        // *** CHANGE TYPE TO 'bar' ***
        type: 'bar',
        data: chartData, // Data structure remains the same
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    // *** UPDATED TITLE TEXT ***
                    text: isEmpty ? 'Bar Chart - Select Subchannel / No Data' : 'Rev AR % vs Unit Achievement % per Store (Subchannel Filtered)',
                    color: titleColor // Dark mode title color
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            let label = context.dataset.label || '';
                            if (label) {
                                label += ': ';
                            }
                            // Use formattedValue for bars, fallback to parsed if needed
                            if (context.formattedValue !== null) {
                                // Tooltips for bar charts often show the direct value,
                                // but we still want percentages here.
                                label += formatPercentage(context.parsed.y);
                            } else {
                                label += 'N/A';
                            }
                            return label;
                        }
                    }
                },
                legend: { // Dark mode legend color
                    labels: {
                        color: textColor
                    }
                }
            },
            interaction: {
                intersect: false,
                mode: 'index', // Shows tooltips for both bars at the same index
            },
            scales: {
                x: {
                    title: {
                        display: true,
                        text: 'Store',
                        color: textColor // Dark mode axis title color
                    },
                    ticks: {
                         color: textColor // Dark mode axis labels color
                    },
                    grid: {
                        color: gridColor // Dark mode grid lines
                    }
                },
                yPercentage: { // Keep using this scale ID
                    type: 'linear',
                    position: 'left',
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Percentage (%)',
                        color: textColor // Dark mode axis title color
                    },
                    ticks: {
                        // Format Y-axis labels as percentages
                        callback: function(value, index, values) {
                            return formatPercentage(value, 0); // e.g., 50%
                        },
                        color: textColor // Dark mode axis labels color
                    },
                    grid: {
                        color: gridColor // Dark mode grid lines
                    }
                }
            },
            // Remove options specific to line charts:
            // tension: 0.1,
            // spanGaps: false
        }
    });
}
function displayAttachRateTable(tableData, columnStats) { attachRateTableBody.innerHTML = ''; attachTableStatusElement.textContent = ''; if (!tableData || tableData.length === 0) { attachTableStatusElement.textContent = 'No data for selected subchannel or required columns missing.'; return; } if (!columnStats) { console.error("Attach Rate Stats missing!"); attachTableStatusElement.textContent = 'Error calculating stats for highlighting.'; columnStats = {}; } tableData.forEach(rowData => { const row = attachRateTableBody.insertRow(); const addCell = (text) => { const cell = row.insertCell(); cell.textContent = text; return cell; }; const addPercentageCellWithHighlight = (rawValue, colName) => { const cell = row.insertCell(); const decimalValue = parsePercentage(rawValue); cell.textContent = formatPercentage(decimalValue); cell.style.textAlign = 'right'; const stats = columnStats[colName] || { average: null, stdDev: null }; if (decimalValue !== null && stats.average !== null && stats.stdDev !== null && stats.stdDev > 0.0001) { const lowerBound = stats.average - stats.stdDev; const upperBound = stats.average + stats.stdDev; if (decimalValue > upperBound) { cell.classList.add('highlight-green'); } else if (decimalValue < lowerBound) { cell.classList.add('highlight-red'); } } return cell; }; addCell(rowData.store); addPercentageCellWithHighlight(rowData['Tablet Attach Rate'], 'Tablet Attach Rate'); addPercentageCellWithHighlight(rowData['PC Attach Rate'], 'PC Attach Rate'); addPercentageCellWithHighlight(rowData['NC Attach Rate'], 'NC Attach Rate'); addPercentageCellWithHighlight(rowData['TWS Attach Rate'], 'TWS Attach Rate'); addPercentageCellWithHighlight(rowData['WW Attach Rate'], 'WW Attach Rate'); addPercentageCellWithHighlight(rowData['ME Attach Rate'], 'ME Attach Rate'); addPercentageCellWithHighlight(rowData['NCME Attach Rate'], 'NCME Attach Rate'); }); attachTableStatusElement.textContent = `Displaying ${tableData.length} stores in subchannel. Highlight: > avg + 1 std dev (green), < avg - 1 std dev (red).`; }


// --- Email Share Handler --- (No changes)
function handleShareEmail() { const recipient = emailRecipientInput.value.trim(); shareStatusElement.textContent = ''; if (!recipient) { shareStatusElement.textContent = 'Please enter a recipient email address.'; emailRecipientInput.focus(); return; } if (!recipient.includes('@') || !recipient.includes('.')) { shareStatusElement.textContent = 'Please enter a valid email address format.'; emailRecipientInput.focus(); return; } let emailBody = "Dashboard Summary:\n"; emailBody += "--------------------\n"; const summaryParas = summaryElement.querySelectorAll('p'); if (summaryParas.length > 0 && !summaryElement.textContent.includes('No matching data')) { summaryParas.forEach(p => { let text = p.textContent.replace(/\s+/g, ' ').trim(); text = text.replace(' (Filtered)', ''); if (text.includes('Territory Rev AR%')) { emailBody += text + "\n"; } else if (text.startsWith('Combined Territory Rev %') || text.startsWith('Territory Rev %')) { emailBody += "\n" + text + "\n"; } else if (text.startsWith('Total Units w/ DF')) { emailBody += "\n" + text + "\n"; } else if (!text.startsWith('Average Territory Rev %')) { emailBody += text + "\n"; } }); } else { emailBody += "(No summary data available for current filters)\n"; } emailBody += "\n"; if (repPmrSection.style.display !== 'none') { emailBody += "Mysteryshop Results:\n"; emailBody += "--------------------\n"; emailBody += repSkillP.textContent.trim() + "\n"; emailBody += vPmrP.textContent.trim() + "\n\n"; } if (trainingStatsSection.style.display !== 'none') { emailBody += "Training Stats:\n"; emailBody += "--------------------\n"; if (eliteP.style.display !== 'none') { emailBody += eliteP.textContent.trim() + "\n"; } if (postTrainingP.style.display !== 'none') { emailBody += postTrainingP.textContent.trim() + "\n"; } emailBody += "\n"; } emailBody += "--\n"; emailBody += "Christopher Wasney\n"; emailBody += "Samsung Field Sales Manager\n"; emailBody += "(989) 751-9277\n"; emailBody += "c9.wasney@sea.samsung.com\n"; const subject = "Dashboard Results Summary"; const encodedSubject = encodeURIComponent(subject); const encodedBody = encodeURIComponent(emailBody); const mailtoLink = `mailto:${recipient}?subject=${encodedSubject}&body=${encodedBody}`; if (mailtoLink.length > 1800) { console.warn("Generated mailto link is very long:", mailtoLink.length, "characters"); shareStatusElement.textContent = 'Warning: Results data might be too long for email; content may be truncated.'; } else { shareStatusElement.textContent = 'Opening your email client...'; } const mailtoWindow = window.open(mailtoLink, '_blank'); if (!mailtoWindow) { console.log("window.open failed, trying location.href for mailto"); window.location.href = mailtoLink; } setTimeout(() => { shareStatusElement.textContent = ''; }, 5000); }

// --- Clear Functions --- (No changes)
function clearChart() { if (lineChartInstance) { lineChartInstance.destroy(); lineChartInstance = null; } }
function clearDashboard() { clearChart(); summaryElement.innerHTML = 'Summary data will appear here...'; summaryElement.style.textAlign = 'center'; statusElement.textContent = 'No file selected.'; subchannelFilter.innerHTML = '<option value="ALL">-- Load File First --</option>'; subchannelFilter.disabled = true; subchannelFilter.value = 'ALL'; storeFilter.innerHTML = '<option value="ALL">-- Load File First --</option>'; resetMultiSelect(storeFilter, 'ALL'); storeFilter.disabled = true; originalData = []; attachRateTableBody.innerHTML = ''; attachTableStatusElement.textContent = ''; if(repPmrSection) repPmrSection.style.display = 'none'; if(trainingStatsSection) trainingStatsSection.style.display = 'none'; if(eliteP) eliteP.style.display = 'none'; if(eliteValue) eliteValue.textContent = 'N/A'; if(postTrainingScoreValue) postTrainingScoreValue.textContent = 'N/A'; if(repSkillAchValue) repSkillAchValue.textContent = 'N/A'; if(vPmrAchValue) vPmrAchValue.textContent = 'N/A'; if(emailRecipientInput) emailRecipientInput.value = ''; if(shareStatusElement) shareStatusElement.textContent = ''; if(fileInput) fileInput.value = ''; }

// --- Helper Function to Format Currency --- (No changes)
function formatCurrency(number) { if (isNaN(number) || number === null) return "$0.00"; const num = parseFloat(number); if (isNaN(num)) return "$0.00"; return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(num); }


// --- PWA Service Worker Registration --- (No changes)
if ('serviceWorker' in navigator) { window.addEventListener('load', () => { navigator.serviceWorker.register('/sw.js') .then(registration => { console.log('Service Worker registered successfully with scope: ', registration.scope); }) .catch(error => { console.log('Service Worker registration failed: ', error); }); }); } else { console.log('Service Worker is not supported by this browser.'); }
