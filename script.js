//
//    Timestamp: 2025-05-31T20:23:00EDT
//    Summary: Implemented logic for showing and dismissing the data accuracy disclaimer banner with 30-day localStorage persistence.
//
document.addEventListener('DOMContentLoaded', () => {
    // --- Theme Constants and Elements ---
    const LIGHT_THEME_CLASS = 'light-theme';
    const THEME_STORAGE_KEY = 'themePreference';
    const DEFAULT_FILTERS_STORAGE_KEY = 'fsmDashboardDefaultFilters_v1'; 
    const DARK_THEME_ICON = '🌙'; 
    const LIGHT_THEME_ICON = '☀️'; 
    const DARK_THEME_META_COLOR = '#2c2c2c';
    const LIGHT_THEME_META_COLOR = '#f4f4f8'; 

    const themeToggleBtn = document.getElementById('themeToggleBtn');
    const metaThemeColorTag = document.querySelector('meta[name="theme-color"]');

    // --- "What's New" Modal Elements ---
    const whatsNewModal = document.getElementById('whatsNewModal');
    const closeWhatsNewModalBtn = document.getElementById('closeWhatsNewModalBtn');
    const gotItWhatsNewBtn = document.getElementById('gotItWhatsNewBtn');
    const BETA_FEATURES_POPUP_COOKIE = 'betaFeaturesPopupShown_v1.3'; 
    const openWhatsNewBtn = document.getElementById('openWhatsNewBtn'); 

    // --- Filter Modal Elements ---
    const filterModal = document.getElementById('filterModal');
    const openFilterModalBtn = document.getElementById('openFilterModalBtn'); 
    const closeFilterModalBtn = document.getElementById('closeFilterModalBtn');
    const applyFiltersButtonModal = document.getElementById('applyFiltersButtonModal');
    const resetFiltersButtonModal = document.getElementById('resetFiltersButtonModal');
    const saveDefaultFiltersBtn = document.getElementById('saveDefaultFiltersBtn');
    const clearDefaultFiltersBtn = document.getElementById('clearDefaultFiltersBtn');
    const filterLoadingIndicatorModal = document.getElementById('filterLoadingIndicatorModal');

    // --- Disclaimer Banner Elements ---
    const dataAccuracyDisclaimer = document.getElementById('dataAccuracyDisclaimer');
    const dismissDisclaimerBtn = document.getElementById('dismissDisclaimerBtn');
    const DISCLAIMER_STORAGE_KEY = 'dataAccuracyDisclaimerDismissed_v1';
    const DISCLAIMER_EXPIRY_DAYS = 30;


    // --- Configuration ---
    const MICHIGAN_AREA_VIEW = { lat: 43.8, lon: -84.8, zoom: 7 }; 
    const AVERAGE_THRESHOLD_PERCENT = 0.10; 

    const REQUIRED_HEADERS = [ 
        'Store', 'REGION', 'DISTRICT', 'Q2 Territory', 'FSM NAME', 'CHANNEL',
        'SUB_CHANNEL', 'DEALER_NAME', 'Revenue w/DF', 'QTD Revenue Target',
        'Quarterly Revenue Target', 'QTD Gap', '% Quarterly Revenue Target', 'Rev AR%', 
        'Unit w/ DF', 'Unit Target', 'Unit Achievement', 'Visit count', 'Trainings',
        'Retail Mode Connectivity', 'Rep Skill Ach', '(V)PMR Ach', 'Elite', 'Post Training Score',
        'Tablet Attach Rate', 'PC Attach Rate', 'NC Attach Rate', 'TWS Attach Rate',
        'WW Attach Rate', 'ME Attach Rate', 'NCME Attach Rate', 'SUPER STORE', 'GOLDEN RHINO',
        'GCE', 'AI_Zone', 'Hispanic_Market', 'EV ROUTE',
        'STORE ID', 'ADDRESS1', 'CITY', 'STATE', 'ZIPCODE',
        'LATITUDE_ORG', 'LONGITUDE_ORG', 
        'ORG_STORE_ID', 'CV_STORE_ID', 'CINGLEPOINT_ID', 
        'STORE_TYPE_NAME', 'National_Tier', 'Merchandising_Level', 'Combined_Tier', 
        '%Quarterly Territory Rev Target', 'Region Rev%', 'District Rev%', 'Territory Rev%'
    ]; 
    const FLAG_HEADERS = ['SUPER STORE', 'GOLDEN RHINO', 'GCE', 'AI_Zone', 'Hispanic_Market', 'EV ROUTE'];
    const ATTACH_RATE_COLUMNS = [ 
        'Tablet Attach Rate', 'PC Attach Rate', 'NC Attach Rate', 
        'TWS Attach Rate', 'WW Attach Rate', 'ME Attach Rate', 'NCME Attach Rate'
    ];
    const CURRENCY_FORMAT = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' });
    const PERCENT_FORMAT = new Intl.NumberFormat('en-US', { style: 'percent', minimumFractionDigits: 1, maximumFractionDigits: 1 });
    const NUMBER_FORMAT = new Intl.NumberFormat('en-US');
    const CHART_COLORS = ['#58a6ff', '#ffb758', '#86dc86', '#ff7f7f', '#b796e6', '#ffda8a', '#8ad7ff', '#ff9ba6'];
    const TOP_N_CHART = 15; 
    const TOP_N_TABLES = 5;

    // --- DOM Elements ---
    const excelFileInput = document.getElementById('excelFile');
    const statusDiv = document.getElementById('status');
    const loadingIndicator = document.getElementById('loadingIndicator');
    const resultsArea = document.getElementById('resultsArea');
    
    const regionFilter = document.getElementById('regionFilter');
    const districtFilter = document.getElementById('districtFilter');
    const territoryFilter = document.getElementById('territoryFilter');
    const fsmFilter = document.getElementById('fsmFilter');
    const channelFilter = document.getElementById('channelFilter');
    const subchannelFilter = document.getElementById('subchannelFilter');
    const dealerFilter = document.getElementById('dealerFilter');
    const storeFilter = document.getElementById('storeFilter');
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
            default: console.warn(`Unknown flag header encountered during mapping: ${header}`); return acc;
        }
        const element = document.getElementById(expectedId);
        if (element) { acc[header] = element; } 
        else { console.warn(`Flag filter checkbox not found for ID: ${expectedId} (Header: ${header}) upon initial mapping. Check HTML.`);}
        return acc;
    }, {});
    const territorySelectAll = document.getElementById('territorySelectAll');
    const territoryDeselectAll = document.getElementById('territoryDeselectAll');
    const storeSelectAll = document.getElementById('storeSelectAll');
    const storeDeselectAll = document.getElementById('storeDeselectAll');
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
    const attachRateTableBody = document.getElementById('attachRateTableBody');
    const attachRateTableFooter = document.getElementById('attachRateTableFooter');
    const attachTableStatus = document.getElementById('attachTableStatus');
    const attachRateTable = document.getElementById('attachRateTable');
    const exportCsvButton = document.getElementById('exportCsvButton');
    const topBottomSection = document.getElementById('topBottomSection');
    const top5TableBody = document.getElementById('top5TableBody');
    const bottom5TableBody = document.getElementById('bottom5TableBody');
    const mainChartCanvas = document.getElementById('mainChartCanvas')?.getContext('2d');
    const storeDetailsSection = document.getElementById('storeDetailsSection');
    const storeDetailsContent = document.getElementById('storeDetailsContent');
    const closeStoreDetailsButton = document.getElementById('closeStoreDetailsButton');
    
    const printReportButton = document.getElementById('printReportButton');
    const emailShareSection = document.getElementById('emailShareSection');
    const emailShareControls = document.getElementById('emailShareControls');
    const emailRecipientInput = document.getElementById('emailRecipient');
    const shareEmailButton = document.getElementById('shareEmailButton');
    const shareStatus = document.getElementById('shareStatus');
    const emailShareHint = document.getElementById('emailShareHint');

    const showMapViewFilter = document.getElementById('showMapViewFilter'); 
    const focusEliteFilter = document.getElementById('focusEliteFilter');
    const focusConnectivityFilter = document.getElementById('focusConnectivityFilter');
    const focusRepSkillFilter = document.getElementById('focusRepSkillFilter');
    const focusVpmrFilter = document.getElementById('focusVpmrFilter');

    const eliteOpportunitiesSection = document.getElementById('eliteOpportunitiesSection');
    const connectivityOpportunitiesSection = document.getElementById('connectivityOpportunitiesSection');
    const repSkillOpportunitiesSection = document.getElementById('repSkillOpportunitiesSection');
    const vpmrOpportunitiesSection = document.getElementById('vpmrOpportunitiesSection');

    const mapViewContainer = document.getElementById('mapViewContainer');
    const mapStatus = document.getElementById('mapStatus');

    // --- Global State ---
    let rawData = [];
    let filteredData = [];
    let mainChartInstance = null;
    let mapInstance = null; 
    let mapMarkersLayer = null;
    let storeOptions = [];
    let allPossibleStores = [];
    let currentSort = { column: 'Store', ascending: true }; 
    let selectedStoreRow = null;

    // --- "What's New" Modal Logic ---
    const showWhatsNewModal = () => {
        if (whatsNewModal) {
            whatsNewModal.style.display = 'flex';
            setTimeout(() => whatsNewModal.classList.add('active'), 10); 
        }
    };
    const hideWhatsNewModal = () => {
        if (whatsNewModal) {
            whatsNewModal.classList.remove('active');
            setTimeout(() => whatsNewModal.style.display = 'none', 300); 
        }
    };
    const checkAndShowWhatsNew = () => { 
        if (!getCookie(BETA_FEATURES_POPUP_COOKIE)) {
            showWhatsNewModal();
        }
    };
    const setCookie = (name, value, days) => {
        let expires = "";
        if (days) {
            const date = new Date();
            date.setTime(date.getTime() + (days * 24 * 60 * 60 * 1000));
            expires = "; expires=" + date.toUTCString();
        }
        document.cookie = name + "=" + (value || "") + expires + "; path=/; SameSite=Lax";
    };
    const getCookie = (name) => {
        const nameEQ = name + "=";
        const ca = document.cookie.split(';');
        for (let i = 0; i < ca.length; i++) {
            let c = ca[i];
            while (c.charAt(0) === ' ') c = c.substring(1, c.length);
            if (c.indexOf(nameEQ) === 0) return c.substring(nameEQ.length, c.length);
        }
        return null;
    };
    if (closeWhatsNewModalBtn) {
        closeWhatsNewModalBtn.addEventListener('click', () => {
            hideWhatsNewModal();
            setCookie(BETA_FEATURES_POPUP_COOKIE, 'true', 365); 
        });
    }
    if (gotItWhatsNewBtn) {
        gotItWhatsNewBtn.addEventListener('click', () => {
            hideWhatsNewModal();
            setCookie(BETA_FEATURES_POPUP_COOKIE, 'true', 365); 
        });
    }
    if (openWhatsNewBtn) { 
        openWhatsNewBtn.addEventListener('click', showWhatsNewModal); 
    }


    // --- Filter Modal Logic ---
    const openFilterModal = () => {
        if (filterModal) {
            filterModal.style.display = 'flex';
            setTimeout(() => filterModal.classList.add('active'), 10);
            document.body.style.overflow = 'hidden'; 
        }
    };
    const closeFilterModal = () => {
        if (filterModal) {
            filterModal.classList.remove('active');
            setTimeout(() => filterModal.style.display = 'none', 300);
            document.body.style.overflow = ''; 
        }
    };
    const mainFilterModalTrigger = document.getElementById('openFilterModalBtn'); 
    if (mainFilterModalTrigger) {
        mainFilterModalTrigger.addEventListener('click', openFilterModal);
    }
    if (closeFilterModalBtn) {
        closeFilterModalBtn.addEventListener('click', closeFilterModal);
    }

    // --- Theme Management ---
    const applyTheme = (theme) => {
        if (theme === 'light') {
            document.body.classList.add(LIGHT_THEME_CLASS);
            if (themeToggleBtn) themeToggleBtn.textContent = DARK_THEME_ICON;
            if (metaThemeColorTag) metaThemeColorTag.setAttribute('content', LIGHT_THEME_META_COLOR);
        } else { 
            document.body.classList.remove(LIGHT_THEME_CLASS);
            if (themeToggleBtn) themeToggleBtn.textContent = LIGHT_THEME_ICON;
            if (metaThemeColorTag) metaThemeColorTag.setAttribute('content', DARK_THEME_META_COLOR);
        }
        if (mainChartInstance && (filteredData.length > 0 || (rawData.length > 0 && filteredData.length === 0) )) { 
             updateCharts(filteredData); 
        }
        if (mapInstance) { 
            setTimeout(() => {
                if (mapInstance && typeof mapInstance.invalidateSize === 'function') {
                    mapInstance.invalidateSize();
                }
            }, 0); 
        }
    };
    const toggleTheme = () => {
        const isLight = document.body.classList.contains(LIGHT_THEME_CLASS);
        const newTheme = isLight ? 'dark' : 'light';
        applyTheme(newTheme);
        localStorage.setItem(THEME_STORAGE_KEY, newTheme);
    };
    const getChartThemeColors = () => {
        const isLight = document.body.classList.contains(LIGHT_THEME_CLASS);
        return {
            tickColor: isLight ? '#495057' : '#e0e0e0',
            gridColor: isLight ? 'rgba(0, 0, 0, 0.1)' : 'rgba(224, 224, 224, 0.2)',
            legendColor: isLight ? '#333333' : '#e0e0e0'
        };
    };

    // --- Disclaimer Banner Logic ---
    const checkAndShowDisclaimer = () => {
        if (!dataAccuracyDisclaimer) return;
        try {
            const disclaimerStateJSON = localStorage.getItem(DISCLAIMER_STORAGE_KEY);
            if (disclaimerStateJSON) {
                const disclaimerState = JSON.parse(disclaimerStateJSON);
                if (disclaimerState.expires && new Date().getTime() < disclaimerState.expires) {
                    dataAccuracyDisclaimer.style.display = 'none'; // Still within 30-day dismissal period
                    return;
                }
            }
            // Show disclaimer if not dismissed or if expiry has passed
            dataAccuracyDisclaimer.style.display = 'flex';
        } catch (e) {
            console.error("Error accessing disclaimer state from localStorage:", e);
            dataAccuracyDisclaimer.style.display = 'flex'; // Show by default on error
        }
    };

    const handleDismissDisclaimer = () => {
        if (!dataAccuracyDisclaimer) return;
        try {
            const expiryDate = new Date();
            expiryDate.setDate(expiryDate.getDate() + DISCLAIMER_EXPIRY_DAYS);
            localStorage.setItem(DISCLAIMER_STORAGE_KEY, JSON.stringify({ expires: expiryDate.getTime() }));
            dataAccuracyDisclaimer.style.display = 'none';
        } catch (e) {
            console.error("Error saving disclaimer state to localStorage:", e);
            // Still hide it for the current session even if localStorage fails
            dataAccuracyDisclaimer.style.display = 'none';
        }
    };

    if (dismissDisclaimerBtn) {
        dismissDisclaimerBtn.addEventListener('click', handleDismissDisclaimer);
    }


    // --- Helper Functions ---
    const formatCurrency = (value) => isNaN(value) ? 'N/A' : CURRENCY_FORMAT.format(value);
    const formatPercent = (value) => isNaN(value) ? 'N/A' : PERCENT_FORMAT.format(value);
    const formatNumber = (value) => isNaN(value) ? 'N/A' : NUMBER_FORMAT.format(value);
    const parseNumber = (value) => {
        if (value === null || value === undefined || String(value).trim() === '') return NaN;
        if (typeof value === 'number') return value;
        if (typeof value === 'string') { const numStr = value.replace(/[$,%]/g, ''); const num = parseFloat(numStr); return isNaN(num) ? NaN : num; }
        return NaN;
    };
    const parsePercent = (value) => {
         if (value === null || value === undefined || String(value).trim() === '') return NaN;
         if (typeof value === 'number') return value; 
         if (typeof value === 'string') { 
             const numStr = value.replace('%', ''); 
             const num = parseFloat(numStr); 
             if (isNaN(num)) return NaN;
             if (value.includes('%') || (num > 1 && num <= 100) || (num === 0) || (num === 1) ) { 
                 return num / 100;
             }
             return num; 
        }
         return NaN;
    };
    const safeGet = (obj, path, defaultValue = 'N/A') => {
        const value = obj ? obj[path] : undefined;
        return (value !== undefined && value !== null && String(value).trim() !== '') ? value : defaultValue;
    };
    const isValidForAverage = (value) => { 
         if (value === null || value === undefined || String(value).trim() === '') return false;
         const parsed = parseNumber(String(value).replace('%','')); 
         return !isNaN(parsed);
    };
    const isValidNumericForFocus = (value) => { 
        if (value === null || value === undefined || String(value).trim() === '') return false;
        const parsedVal = parsePercent(value); 
        return !isNaN(parsedVal);
    };
    const calculateQtdGap = (row) => {
        const revenue = parseNumber(safeGet(row, 'Revenue w/DF', 0)); 
        const target = parseNumber(safeGet(row, 'QTD Revenue Target', 0));
        if (isNaN(revenue) || isNaN(target)) { return Infinity; } 
        return revenue - target;
    };
    const calculateRevARPercentForRow = (row) => { 
        const revenue = parseNumber(safeGet(row, 'Revenue w/DF', 0));
        const target = parseNumber(safeGet(row, 'QTD Revenue Target', 0));
        if (target === 0 || isNaN(revenue) || isNaN(target)) { return NaN; } 
        return revenue / target;
    };
    const calculateUnitAchievementPercentForRow = (row) => { 
        const units = parseNumber(safeGet(row, 'Unit w/ DF', 0));
        const target = parseNumber(safeGet(row, 'Unit Target', 0));
        if (target === 0 || isNaN(units) || isNaN(target)) { return NaN; } 
        return units / target;
    };
    const getUniqueValues = (data, column) => {
        const values = new Set(data.map(item => safeGet(item, column, '')).filter(val => String(val).trim() !== ''));
        return ['ALL', ...Array.from(values).sort()];
    };
    const setOptions = (selectElement, options, disable = false) => {
        if (!selectElement) return;
        selectElement.innerHTML = '';
        options.forEach(optionValue => { const option = document.createElement('option'); option.value = optionValue; option.textContent = optionValue === 'ALL' ? `-- ${optionValue} --` : optionValue; option.title = optionValue; selectElement.appendChild(option); });
        selectElement.disabled = disable;
    };
    const setMultiSelectOptions = (selectElement, options, disable = false) => {
        if (!selectElement) return;
        selectElement.innerHTML = '';
        options.forEach(optionValue => { if (optionValue === 'ALL') return; const option = document.createElement('option'); option.value = optionValue; option.textContent = optionValue; option.title = optionValue; selectElement.appendChild(option); });
        selectElement.disabled = disable;
    };
    const showLoading = (isLoading, isFiltering = false) => {
        const displayStyle = isLoading ? 'flex' : 'none';
        const indicator = isFiltering ? filterLoadingIndicatorModal : loadingIndicator;
        const primaryButton = isFiltering ? applyFiltersButtonModal : excelFileInput; 
    
        if (indicator) indicator.style.display = displayStyle;
        if (primaryButton) primaryButton.disabled = isLoading;
    
        if (isFiltering) {
            if (resetFiltersButtonModal) resetFiltersButtonModal.disabled = isLoading;
            if (saveDefaultFiltersBtn) saveDefaultFiltersBtn.disabled = isLoading;
            if (clearDefaultFiltersBtn) clearDefaultFiltersBtn.disabled = isLoading || localStorage.getItem(DEFAULT_FILTERS_STORAGE_KEY) === null;
            const mainFilterBtn = document.getElementById('openFilterModalBtn'); 
            if (mainFilterBtn) mainFilterBtn.disabled = isLoading;
        } else {
            if (excelFileInput) excelFileInput.disabled = isLoading;
            const mainFilterBtn = document.getElementById('openFilterModalBtn'); 
            if (mainFilterBtn) mainFilterBtn.disabled = isLoading; 
        }
    };    

    // --- Map Functions ---
    const initMapView = () => {
        if (typeof L === 'undefined' || !L || typeof L.map !== 'function') {
            console.error("Leaflet library (L) or L.map is not available. Map cannot be initialized.");
            if (mapStatus) mapStatus.textContent = 'Map library failed to load.';
            if (mapViewContainer) mapViewContainer.style.display = 'none';
            return;
        }
        const mapElement = document.getElementById('mapid');
        if (mapElement && !mapInstance) {
            try {
                mapInstance = L.map(mapElement).setView([MICHIGAN_AREA_VIEW.lat, MICHIGAN_AREA_VIEW.lon], MICHIGAN_AREA_VIEW.zoom); 
                L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
                    attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
                    maxZoom: 18, tileSize: 512, zoomOffset: -1
                }).addTo(mapInstance);
                mapMarkersLayer = L.layerGroup().addTo(mapInstance);
                if (mapStatus) mapStatus.textContent = 'Map ready. Enable via "Additional Tools" and apply filters to see stores.';
            } catch (e) {
                console.error("Error initializing Leaflet map:", e);
                if (mapStatus) mapStatus.textContent = 'Error initializing map.';
                mapInstance = null; mapMarkersLayer = null;
                if (mapViewContainer) mapViewContainer.style.display = 'none';
            }
        } else if (!mapElement) {
            if (mapStatus) mapStatus.textContent = 'Map container not found.';
        } else if (mapInstance && !mapMarkersLayer) {
            mapMarkersLayer = L.layerGroup().addTo(mapInstance);
        }
    };
    const updateMapView = (data) => {
        if (!showMapViewFilter || !showMapViewFilter.checked) {
            if (mapViewContainer) mapViewContainer.style.display = 'none';
            if (mapInstance && mapMarkersLayer?.clearLayers) mapMarkersLayer.clearLayers(); 
            return; 
        }
        if (!mapInstance || !document.getElementById('mapid')) {
            if (mapStatus) mapStatus.textContent = 'Map is not initialized or container is missing.';
            if (mapViewContainer) mapViewContainer.style.display = 'block'; 
            return;
        }
        if (!mapMarkersLayer || !(mapMarkersLayer instanceof L.LayerGroup)) {
             if (mapMarkersLayer?.remove) mapMarkersLayer.remove();
            mapMarkersLayer = L.layerGroup().addTo(mapInstance);
            if (!(mapMarkersLayer instanceof L.LayerGroup)) { 
                if (mapStatus) mapStatus.textContent = 'Map layer component error. Please refresh.';
                if (mapViewContainer) mapViewContainer.style.display = 'block'; 
                return;
            }
        }
        mapMarkersLayer.clearLayers();
        let storesOnMapCount = 0;
        const validStoresWithCoords = data.filter(row => {
            const lat = parseNumber(safeGet(row, 'LATITUDE_ORG', NaN));
            const lon = parseNumber(safeGet(row, 'LONGITUDE_ORG', NaN));
            return !isNaN(lat) && !isNaN(lon) && lat !== 0 && lon !== 0  && lat >= -90 && lat <= 90 && lon >= -180 && lon <= 180;
        });
        if (mapViewContainer) mapViewContainer.style.display = 'block'; 
        if (validStoresWithCoords.length === 0) {
            if (mapStatus) mapStatus.textContent = 'No stores with valid coordinates in filtered data. Showing default map area.';
            mapInstance.setView([MICHIGAN_AREA_VIEW.lat, MICHIGAN_AREA_VIEW.lon], MICHIGAN_AREA_VIEW.zoom);
            return;
        }
        validStoresWithCoords.forEach(row => {
            const lat = parseNumber(safeGet(row, 'LATITUDE_ORG', NaN));
            const lon = parseNumber(safeGet(row, 'LONGITUDE_ORG', NaN));
            const storeName = safeGet(row, 'Store', 'Unknown Store');
            const revenue = formatCurrency(parseNumber(safeGet(row, 'Revenue w/DF', NaN)));
            const qtdGapVal = calculateQtdGap(row);
            const formattedQtdGap = isNaN(qtdGapVal) || qtdGapVal === Infinity ? 'N/A' : formatCurrency(qtdGapVal);
            const popupContent = `<strong>${storeName}</strong><br>Revenue: ${revenue}<br>QTD Gap: ${formattedQtdGap}`;
            const marker = L.marker([lat, lon]);
            marker.bindPopup(popupContent);
            marker.on('click', () => { showStoreDetails(row); highlightTableRow(storeName); });
            mapMarkersLayer.addLayer(marker);
            storesOnMapCount++;
        });
        if (storesOnMapCount > 0) {
            if (mapMarkersLayer?.getBounds) {
                const bounds = mapMarkersLayer.getBounds();
                if (bounds?.isValid()) {
                    mapInstance.fitBounds(bounds, { padding: [25, 25], maxZoom: 16 }); 
                } else {
                    if (validStoresWithCoords.length > 0) {
                        const firstStoreWithCoords = validStoresWithCoords[0];
                        const lat = parseNumber(safeGet(firstStoreWithCoords, 'LATITUDE_ORG'));
                        const lon = parseNumber(safeGet(firstStoreWithCoords, 'LONGITUDE_ORG'));
                        if(!isNaN(lat) && !isNaN(lon)) mapInstance.setView([lat, lon], 10);
                    } else { mapInstance.setView([MICHIGAN_AREA_VIEW.lat, MICHIGAN_AREA_VIEW.lon], MICHIGAN_AREA_VIEW.zoom); }
                }
            }
            if (mapStatus) mapStatus.textContent = `Displaying ${storesOnMapCount} stores on map.`;
        } else { 
            if (mapStatus) mapStatus.textContent = 'No stores with displayable coordinates in filtered data. Showing default map area.';
             mapInstance.setView([MICHIGAN_AREA_VIEW.lat, MICHIGAN_AREA_VIEW.lon], MICHIGAN_AREA_VIEW.zoom);
        }
        setTimeout(() => { if (mapInstance) mapInstance.invalidateSize(); }, 0);
    };

    // --- Core Functions ---
    const handleFile = async (event) => {
        const file = event.target.files[0];
        if (!file) { if (statusDiv) statusDiv.textContent = 'No file selected.'; return; }
        if (statusDiv) statusDiv.textContent = 'Reading file...';
        showLoading(true, false); 
        if (resultsArea) resultsArea.style.display = 'none'; 
        resetUI(); 
        try {
            const data = await file.arrayBuffer(); const workbook = XLSX.read(data); const firstSheetName = workbook.SheetNames[0]; const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: null });
            if (jsonData.length > 0) { const headers = Object.keys(jsonData[0]); const missingHeaders = REQUIRED_HEADERS.filter(h => !headers.includes(h)); if (missingHeaders.length > 0) { console.warn(`Warning: Missing expected columns: ${missingHeaders.join(', ')}.`); }
            } else { throw new Error("Excel sheet appears to be empty."); }
            
            rawData = jsonData; 
            allPossibleStores = [...new Set(rawData.map(r => safeGet(r, 'Store', null)).filter(s => s && String(s).trim() !== ''))].sort().map(s => ({ value: s, text: s }));
            
            populateFilters(rawData); 
            const mainFilterBtn = document.getElementById('openFilterModalBtn');
            if (mainFilterBtn) mainFilterBtn.disabled = false;
            
            if (loadDefaultFilters()) { 
                if (statusDiv) statusDiv.textContent += ` Loaded ${rawData.length} rows. Default filters applied.`;
                applyFilters(true); // Apply loaded defaults and close modal if open
            } else {
                if (statusDiv) statusDiv.textContent = `Loaded ${rawData.length} rows. Filters opened automatically.`;
                setTimeout(() => { openFilterModal(); }, 100); 
            }

        } catch (error) {
            console.error('Error processing file:', error); if (statusDiv) statusDiv.textContent = `Error: ${error.message}`;
            rawData = []; allPossibleStores = []; filteredData = []; resetUI();
        } finally { 
            showLoading(false, false); 
            if (excelFileInput) excelFileInput.value = ''; 
        }
    };
    const populateFilters = (data) => {
        setOptions(regionFilter, getUniqueValues(data, 'REGION')); setOptions(districtFilter, getUniqueValues(data, 'DISTRICT')); setMultiSelectOptions(territoryFilter, getUniqueValues(data, 'Q2 Territory').slice(1));
        setOptions(fsmFilter, getUniqueValues(data, 'FSM NAME')); setOptions(channelFilter, getUniqueValues(data, 'CHANNEL')); setOptions(subchannelFilter, getUniqueValues(data, 'SUB_CHANNEL')); setOptions(dealerFilter, getUniqueValues(data, 'DEALER_NAME'));
        Object.values(flagFiltersCheckboxes).forEach(input => { if(input) input.disabled = false; });
        
        if(showMapViewFilter) showMapViewFilter.disabled = false;
        if(focusEliteFilter) focusEliteFilter.disabled = false;
        if(focusConnectivityFilter) focusConnectivityFilter.disabled = false;
        if(focusRepSkillFilter) focusRepSkillFilter.disabled = false;
        if(focusVpmrFilter) focusVpmrFilter.disabled = false;

        storeOptions = [...allPossibleStores]; setStoreFilterOptions(storeOptions, false);
        if (territorySelectAll) territorySelectAll.disabled = false; if (territoryDeselectAll) territoryDeselectAll.disabled = false;
        if (storeSelectAll) storeSelectAll.disabled = false; if (storeDeselectAll) storeDeselectAll.disabled = false;
        if (storeSearch) storeSearch.disabled = false; 
        
        if (applyFiltersButtonModal) applyFiltersButtonModal.disabled = false;
        if (resetFiltersButtonModal) resetFiltersButtonModal.disabled = false; 
        if (saveDefaultFiltersBtn) saveDefaultFiltersBtn.disabled = false;
        if (clearDefaultFiltersBtn) clearDefaultFiltersBtn.disabled = localStorage.getItem(DEFAULT_FILTERS_STORAGE_KEY) === null;
        
        addDependencyFilterListeners();
    };

    const saveDefaultFilters = () => {
        const defaultFilters = {
            region: regionFilter?.value,
            district: districtFilter?.value,
            territories: territoryFilter ? Array.from(territoryFilter.selectedOptions).map(opt => opt.value) : [],
            fsm: fsmFilter?.value,
            channel: channelFilter?.value,
            subchannel: subchannelFilter?.value,
            dealer: dealerFilter?.value,
            stores: storeFilter ? Array.from(storeFilter.selectedOptions).map(opt => opt.value) : [],
            flags: FLAG_HEADERS.reduce((acc, header) => {
                if (flagFiltersCheckboxes[header]) {
                    acc[header] = flagFiltersCheckboxes[header].checked;
                }
                return acc;
            }, {}),
            additionalTools: {
                mapView: showMapViewFilter?.checked,
                elite: focusEliteFilter?.checked,
                connectivity: focusConnectivityFilter?.checked,
                repSkill: focusRepSkillFilter?.checked,
                vpmr: focusVpmrFilter?.checked
            }
        };
        try {
            localStorage.setItem(DEFAULT_FILTERS_STORAGE_KEY, JSON.stringify(defaultFilters));
            if(statusDiv) statusDiv.textContent = "⭐ Default filters saved successfully!";
            if(clearDefaultFiltersBtn) clearDefaultFiltersBtn.disabled = false; 
        } catch (e) {
            console.error("Error saving default filters to localStorage:", e);
            if(statusDiv) statusDiv.textContent = "Error saving default filters. Your browser might be blocking localStorage or it's full.";
        }
    };

    const loadDefaultFilters = () => {
        let defaultsApplied = false;
        try {
            const savedFiltersJSON = localStorage.getItem(DEFAULT_FILTERS_STORAGE_KEY);
            if (!savedFiltersJSON) {
                if(clearDefaultFiltersBtn) clearDefaultFiltersBtn.disabled = true;
                return false;
            }

            const defaults = JSON.parse(savedFiltersJSON);
            let partialLoadMessage = "";

            if (regionFilter && defaults.region) regionFilter.value = defaults.region;
            if (districtFilter && defaults.district) districtFilter.value = defaults.district;
            if (fsmFilter && defaults.fsm) fsmFilter.value = defaults.fsm;
            if (channelFilter && defaults.channel) channelFilter.value = defaults.channel;
            if (subchannelFilter && defaults.subchannel) subchannelFilter.value = defaults.subchannel;
            if (dealerFilter && defaults.dealer) dealerFilter.value = defaults.dealer;

            if (territoryFilter && defaults.territories) {
                const availableTerritories = Array.from(territoryFilter.options).map(opt => opt.value);
                let foundCount = 0;
                defaults.territories.forEach(savedTerritory => {
                    const option = Array.from(territoryFilter.options).find(opt => opt.value === savedTerritory);
                    if (option) {
                        option.selected = true;
                        foundCount++;
                    }
                });
                if (foundCount < defaults.territories.length && defaults.territories.length > 0) {
                    partialLoadMessage += ` Some saved territories not found in current file.`;
                }
            } else if (territoryFilter) {
                territoryFilter.selectedIndex = -1;
            }
            
            // Temporarily store saved store selections
            if (storeFilter && defaults.stores) {
                 storeFilter.savedDefaults = defaults.stores; 
            } else if (storeFilter) {
                 delete storeFilter.savedDefaults;
            }

            FLAG_HEADERS.forEach(header => {
                if (flagFiltersCheckboxes[header] && defaults.flags && typeof defaults.flags[header] === 'boolean') {
                    flagFiltersCheckboxes[header].checked = defaults.flags[header];
                }
            });

            if (defaults.additionalTools) {
                if (showMapViewFilter && typeof defaults.additionalTools.mapView === 'boolean') showMapViewFilter.checked = defaults.additionalTools.mapView;
                if (focusEliteFilter && typeof defaults.additionalTools.elite === 'boolean') focusEliteFilter.checked = defaults.additionalTools.elite;
                if (focusConnectivityFilter && typeof defaults.additionalTools.connectivity === 'boolean') focusConnectivityFilter.checked = defaults.additionalTools.connectivity;
                if (focusRepSkillFilter && typeof defaults.additionalTools.repSkill === 'boolean') focusRepSkillFilter.checked = defaults.additionalTools.repSkill;
                if (focusVpmrFilter && typeof defaults.additionalTools.vpmr === 'boolean') focusVpmrFilter.checked = defaults.additionalTools.vpmr;
            }
            
            updateStoreFilterOptionsBasedOnHierarchy(); 

            // Apply saved store selections after store options are updated
            if (storeFilter && storeFilter.savedDefaults) {
                let foundStoreCount = 0;
                // Deselect all first to ensure clean application of defaults
                Array.from(storeFilter.options).forEach(opt => opt.selected = false);

                storeFilter.savedDefaults.forEach(savedStore => {
                    const option = Array.from(storeFilter.options).find(opt => opt.value === savedStore);
                    if (option) {
                        option.selected = true;
                        foundStoreCount++;
                    }
                });
                 if (foundStoreCount < storeFilter.savedDefaults.length && storeFilter.savedDefaults.length > 0) {
                    partialLoadMessage += ` Some saved stores not found in current file.`;
                }
                delete storeFilter.savedDefaults; 
            } else if (storeFilter) {
                 storeFilter.selectedIndex = -1; // Deselect if no saved stores
            }


            if (statusDiv) statusDiv.textContent = "⭐ Default filters loaded." + partialLoadMessage;
            if (clearDefaultFiltersBtn) clearDefaultFiltersBtn.disabled = false;
            defaultsApplied = true;

        } catch (e) {
            console.error("Error loading default filters from localStorage:", e);
            if (statusDiv) statusDiv.textContent = "Error loading default filters. They might be corrupted.";
            localStorage.removeItem(DEFAULT_FILTERS_STORAGE_KEY); 
            if (clearDefaultFiltersBtn) clearDefaultFiltersBtn.disabled = true;
            return false;
        }
        return defaultsApplied;
    };

    const clearDefaultFilters = () => {
        try {
            localStorage.removeItem(DEFAULT_FILTERS_STORAGE_KEY);
            if(statusDiv) statusDiv.textContent = "🗑️ Default filters cleared.";
            if(clearDefaultFiltersBtn) clearDefaultFiltersBtn.disabled = true; 
        } catch (e) {
            console.error("Error clearing default filters from localStorage:", e);
            if(statusDiv) statusDiv.textContent = "Error clearing default filters.";
        }
    };

    const addDependencyFilterListeners = () => {
        const handler = updateStoreFilterOptionsBasedOnHierarchy; const filtersToListen = [regionFilter, districtFilter, territoryFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter];
        filtersToListen.forEach(filter => { if (filter) { filter.removeEventListener('change', handler); filter.addEventListener('change', handler); } });
        Object.values(flagFiltersCheckboxes).forEach(input => { if (input) { input.removeEventListener('change', handler); input.addEventListener('change', handler); } });
    };
    const updateStoreFilterOptionsBasedOnHierarchy = () => {
        if (rawData.length === 0) return;
        const selectedRegion = regionFilter?.value; const selectedDistrict = districtFilter?.value; const selectedTerritories = territoryFilter ? Array.from(territoryFilter.selectedOptions).map(opt => opt.value) : [];
        const selectedFsm = fsmFilter?.value; const selectedChannel = channelFilter?.value; const selectedSubchannel = subchannelFilter?.value; const selectedDealer = dealerFilter?.value;
        const selectedFlags = {}; Object.entries(flagFiltersCheckboxes).forEach(([key, input]) => { if (input?.checked) { selectedFlags[key] = true; } });
        const potentiallyValidStoresData = rawData.filter(row => {
            if (selectedRegion !== 'ALL' && safeGet(row, 'REGION', null) !== selectedRegion) return false; if (selectedDistrict !== 'ALL' && safeGet(row, 'DISTRICT', null) !== selectedDistrict) return false;
            if (selectedTerritories.length > 0 && !selectedTerritories.includes(safeGet(row, 'Q2 Territory', null))) return false; if (selectedFsm !== 'ALL' && safeGet(row, 'FSM NAME', null) !== selectedFsm) return false;
            if (selectedChannel !== 'ALL' && safeGet(row, 'CHANNEL', null) !== selectedChannel) return false; if (selectedSubchannel !== 'ALL' && safeGet(row, 'SUB_CHANNEL', null) !== selectedSubchannel) return false;
            if (selectedDealer !== 'ALL' && safeGet(row, 'DEALER_NAME', null) !== selectedDealer) return false;
            for (const flag in selectedFlags) { const flagValue = safeGet(row, flag, 'NO'); if (!(flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1')) { return false; } }
            return true;
        });
        const validStoreNames = new Set(potentiallyValidStoresData.map(row => safeGet(row, 'Store', null)).filter(s => s && String(s).trim() !== ''));
        storeOptions = Array.from(validStoreNames).sort().map(s => ({ value: s, text: s }));
        const previouslySelectedStores = storeFilter ? new Set(Array.from(storeFilter.selectedOptions).map(opt => opt.value)) : new Set();
        
        const currentSearchTerm = storeSearch?.value || ''; 
        setStoreFilterOptions(storeOptions, storeFilter ? storeFilter.disabled : true); 
        filterStoreOptions(); 

        if (storeFilter) { 
            Array.from(storeFilter.options).forEach(option => { 
                if (previouslySelectedStores.has(option.value)) { 
                    option.selected = true; 
                } 
            }); 
            if (storeFilter.selectedOptions.length === 0 && previouslySelectedStores.size > 0) {
            } else if (storeFilter.selectedOptions.length === 0) {
                storeFilter.selectedIndex = -1;
            }
        }
    };
    const setStoreFilterOptions = (optionsToShow, disable = true) => {
        if (!storeFilter) return; const currentSearchTerm = storeSearch?.value || ''; storeFilter.innerHTML = '';
        optionsToShow.forEach(opt => { const option = document.createElement('option'); option.value = opt.value; option.textContent = opt.text; option.title = opt.text; storeFilter.appendChild(option); });
        storeFilter.disabled = disable; if (storeSearch) storeSearch.disabled = disable;
        if (storeSelectAll) storeSelectAll.disabled = disable || optionsToShow.length === 0; if (storeDeselectAll) storeDeselectAll.disabled = disable || optionsToShow.length === 0;
        if (storeSearch) storeSearch.value = currentSearchTerm;
    };
    const filterStoreOptions = () => {
        if (!storeFilter || !storeSearch) return; const searchTerm = storeSearch.value.toLowerCase();
        const filteredOptions = storeOptions.filter(opt => opt.text.toLowerCase().includes(searchTerm)); const selectedValues = new Set(Array.from(storeFilter.selectedOptions).map(opt => opt.value));
        storeFilter.innerHTML = ''; filteredOptions.forEach(opt => { const option = document.createElement('option'); option.value = opt.value; option.textContent = opt.text; option.title = opt.text; if (selectedValues.has(opt.value)) { option.selected = true; } storeFilter.appendChild(option); });
        if (storeSelectAll) storeSelectAll.disabled = storeFilter.disabled || filteredOptions.length === 0; if (storeDeselectAll) storeDeselectAll.disabled = storeFilter.disabled || filteredOptions.length === 0;
    };
    const updateTopBottomTables = (data) => {
        if (!topBottomSection || !top5TableBody || !bottom5TableBody) { console.warn("Top/Bottom table elements not found."); return; }
        top5TableBody.innerHTML = ''; bottom5TableBody.innerHTML = '';
        const territoriesInData = new Set(data.map(row => safeGet(row, 'Q2 Territory', null)).filter(Boolean));
        const isSingleTerritorySelected = territoriesInData.size === 1;
        if (!isSingleTerritorySelected || data.length === 0) { topBottomSection.style.display = 'none'; return; }
        topBottomSection.style.display = 'flex'; 
        const top5Data = [...data].sort((a, b) => parseNumber(safeGet(b, 'Revenue w/DF', -Infinity)) - parseNumber(safeGet(a, 'Revenue w/DF', -Infinity))).slice(0, TOP_N_TABLES);
        top5Data.forEach(row => {
            const tr = top5TableBody.insertRow(); const storeName = safeGet(row, 'Store', 'N/A'); tr.dataset.storeName = storeName;
            tr.onclick = () => { showStoreDetails(row); highlightTableRow(storeName); };
            const revenue = parseNumber(safeGet(row, 'Revenue w/DF', NaN)); const revAR = calculateRevARPercentForRow(row); const unitAch = calculateUnitAchievementPercentForRow(row); const visits = parseNumber(safeGet(row, 'Visit count', NaN));
            tr.insertCell().textContent = storeName; tr.cells[0].title = storeName; tr.insertCell().textContent = formatCurrency(revenue); tr.cells[1].title = formatCurrency(revenue);
            tr.insertCell().textContent = formatPercent(revAR); tr.cells[2].title = formatPercent(revAR); tr.insertCell().textContent = formatPercent(unitAch); tr.cells[3].title = formatPercent(unitAch);
            tr.insertCell().textContent = formatNumber(visits); tr.cells[4].title = formatNumber(visits);
        });
        const bottom5Data = [...data].sort((a, b) => calculateQtdGap(a) - calculateQtdGap(b)).slice(0, TOP_N_TABLES);
        bottom5Data.forEach(row => {
            const tr = bottom5TableBody.insertRow(); const storeName = safeGet(row, 'Store', 'N/A'); tr.dataset.storeName = storeName;
            tr.onclick = () => { showStoreDetails(row); highlightTableRow(storeName); };
            const qtdGap = calculateQtdGap(row); const revAR = calculateRevARPercentForRow(row); const unitAch = calculateUnitAchievementPercentForRow(row); const visits = parseNumber(safeGet(row, 'Visit count', NaN));
            tr.insertCell().textContent = storeName; tr.cells[0].title = storeName; tr.insertCell().textContent = formatCurrency(qtdGap === Infinity ? NaN : qtdGap); tr.cells[1].title = formatCurrency(qtdGap === Infinity ? NaN : qtdGap);
            tr.insertCell().textContent = formatPercent(revAR); tr.cells[2].title = formatPercent(revAR); tr.insertCell().textContent = formatPercent(unitAch); tr.cells[3].title = formatPercent(unitAch);
            tr.insertCell().textContent = formatNumber(visits); tr.cells[4].title = formatNumber(visits);
        });
    };
    
    const applyFilters = (isFromModalOrDefaults = false) => { 
        showLoading(true, true); 
        if (resultsArea) resultsArea.style.display = 'none';
        
        setTimeout(() => {
            try {
                const selectedRegion = regionFilter?.value; const selectedDistrict = districtFilter?.value; const selectedTerritories = territoryFilter ? Array.from(territoryFilter.selectedOptions).map(opt => opt.value) : [];
                const selectedFsm = fsmFilter?.value; const selectedChannel = channelFilter?.value; const selectedSubchannel = subchannelFilter?.value; const selectedDealer = dealerFilter?.value;
                const selectedStores = storeFilter ? Array.from(storeFilter.selectedOptions).map(opt => opt.value) : [];
                const selectedFlags = {}; Object.entries(flagFiltersCheckboxes).forEach(([key, input]) => { if (input?.checked) { selectedFlags[key] = true; } });
                filteredData = rawData.filter(row => {
                    if (selectedRegion !== 'ALL' && safeGet(row, 'REGION', null) !== selectedRegion) return false; if (selectedDistrict !== 'ALL' && safeGet(row, 'DISTRICT', null) !== selectedDistrict) return false;
                    if (selectedTerritories.length > 0 && !selectedTerritories.includes(safeGet(row, 'Q2 Territory', null))) return false; if (selectedFsm !== 'ALL' && safeGet(row, 'FSM NAME', null) !== selectedFsm) return false;
                    if (selectedChannel !== 'ALL' && safeGet(row, 'CHANNEL', null) !== selectedChannel) return false; if (selectedSubchannel !== 'ALL' && safeGet(row, 'SUB_CHANNEL', null) !== selectedSubchannel) return false;
                    if (selectedDealer !== 'ALL' && safeGet(row, 'DEALER_NAME', null) !== selectedDealer) return false; if (selectedStores.length > 0 && !selectedStores.includes(safeGet(row, 'Store', null))) return false;
                    for (const flag in selectedFlags) { const flagValue = safeGet(row, flag, 'NO'); if (!(flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1')) { return false; } }
                    return true;
                });
                updateSummary(filteredData); updateTopBottomTables(filteredData); updateCharts(filteredData); updateAttachRateTable(filteredData); 
                
                if (showMapViewFilter && showMapViewFilter.checked) {
                    updateMapView(filteredData);
                } else {
                    if (mapViewContainer) mapViewContainer.style.display = 'none';
                    if (mapInstance && mapMarkersLayer?.clearLayers) mapMarkersLayer.clearLayers();
                }
                
                updateFocusPointSections(filteredData);
                updateShareOptions(); 

                if (filteredData.length === 1) { showStoreDetails(filteredData[0]); highlightTableRow(safeGet(filteredData[0], 'Store', null)); } else { hideStoreDetails(); }
                if (statusDiv && !statusDiv.textContent.includes("Default filters loaded")) { 
                     statusDiv.textContent = `Displaying ${filteredData.length} of ${rawData.length} rows based on filters.`;
                } else if (statusDiv && statusDiv.textContent.includes("Default filters loaded") && rawData.length > 0) {
                     statusDiv.textContent = statusDiv.textContent.split('.')[0] + `. Displaying ${filteredData.length} of ${rawData.length} rows.`;
                } else if (statusDiv && rawData.length === 0) {
                    statusDiv.textContent = 'No data to display.';
                }


                if (resultsArea) resultsArea.style.display = 'block'; 
                if (exportCsvButton) exportCsvButton.disabled = filteredData.length === 0;
                if (printReportButton) printReportButton.disabled = filteredData.length === 0;
                
                if (isFromModalOrDefaults && filterModal && filterModal.classList.contains('active')) { 
                    closeFilterModal();
                }

            } catch (error) {
                console.error("Error applying filters:", error); if (statusDiv) statusDiv.textContent = "Error applying filters. Check console for details.";
                filteredData = []; if (resultsArea) resultsArea.style.display = 'none'; 
                if (exportCsvButton) exportCsvButton.disabled = true;
                if (printReportButton) printReportButton.disabled = true;
                if (emailShareSection) emailShareSection.style.display = 'none';
                updateSummary([]); updateTopBottomTables([]); updateCharts([]); updateAttachRateTable([]); updateMapView([]); hideStoreDetails();
                if (eliteOpportunitiesSection) eliteOpportunitiesSection.style.display = 'none';
                if (connectivityOpportunitiesSection) connectivityOpportunitiesSection.style.display = 'none';
                if (repSkillOpportunitiesSection) repSkillOpportunitiesSection.style.display = 'none';
                if (vpmrOpportunitiesSection) vpmrOpportunitiesSection.style.display = 'none';
                if (mapViewContainer) mapViewContainer.style.display = 'none';
            } finally { 
                showLoading(false, true); 
                const mainFilterBtn = document.getElementById('openFilterModalBtn');
                if (mainFilterBtn) mainFilterBtn.disabled = rawData.length === 0;
            }
        }, 10);
    };
    
    const resetFiltersForFullUIReset = () => {
         const allOptionHTML = '<option value="ALL">-- Load File First --</option>';
         [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(sel => { 
             if (sel) { sel.innerHTML = allOptionHTML; sel.value = 'ALL'; sel.disabled = true;}
         });
         if (territoryFilter) { territoryFilter.innerHTML = '<option value="ALL">-- Load File First --</option>'; territoryFilter.selectedIndex = -1; territoryFilter.disabled = true; }
         if (storeFilter) { storeFilter.innerHTML = '<option value="ALL">-- Load File First --</option>'; storeFilter.selectedIndex = -1; storeFilter.disabled = true; }
         if (storeSearch) { storeSearch.value = ''; storeSearch.disabled = true; }
         storeOptions = []; 
         Object.values(flagFiltersCheckboxes).forEach(input => { if(input) {input.checked = false; input.disabled = true;} });
        
        if(showMapViewFilter) { showMapViewFilter.checked = false; showMapViewFilter.disabled = true; }
        if(focusEliteFilter) { focusEliteFilter.checked = false; focusEliteFilter.disabled = true; }
        if(focusConnectivityFilter) { focusConnectivityFilter.checked = false; focusConnectivityFilter.disabled = true; }
        if(focusRepSkillFilter) { focusRepSkillFilter.checked = false; focusRepSkillFilter.disabled = true; }
        if(focusVpmrFilter) { focusVpmrFilter.checked = false; focusVpmrFilter.disabled = true; }

         if (applyFiltersButtonModal) applyFiltersButtonModal.disabled = true;
         if (resetFiltersButtonModal) resetFiltersButtonModal.disabled = true; 
         if (saveDefaultFiltersBtn) saveDefaultFiltersBtn.disabled = true;
         if (clearDefaultFiltersBtn) clearDefaultFiltersBtn.disabled = true;

         const mainFilterBtn = document.getElementById('openFilterModalBtn');
         if (mainFilterBtn) mainFilterBtn.disabled = true;

         if (territorySelectAll) territorySelectAll.disabled = true; if (territoryDeselectAll) territoryDeselectAll.disabled = true;
         if (storeSelectAll) storeSelectAll.disabled = true; if (storeDeselectAll) storeDeselectAll.disabled = true;
         if (exportCsvButton) exportCsvButton.disabled = true;
         if (printReportButton) printReportButton.disabled = true;
         if (emailShareSection) emailShareSection.style.display = 'none';
         const handler = updateStoreFilterOptionsBasedOnHierarchy;
         [regionFilter, districtFilter, territoryFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(filter => { if (filter) filter.removeEventListener('change', handler); });
         Object.values(flagFiltersCheckboxes).forEach(input => { if (input) input.removeEventListener('change', handler); });
    };

     const resetUI = () => {
         resetFiltersForFullUIReset(); 
         if (resultsArea) resultsArea.style.display = 'none';
         if (mainChartInstance) { mainChartInstance.destroy(); mainChartInstance = null; }
         
         if (mapInstance && mapMarkersLayer?.clearLayers) {
             mapMarkersLayer.clearLayers();
             mapInstance.setView([MICHIGAN_AREA_VIEW.lat, MICHIGAN_AREA_VIEW.lon], MICHIGAN_AREA_VIEW.zoom); 
         } else if (mapInstance) { 
             mapInstance.setView([MICHIGAN_AREA_VIEW.lat, MICHIGAN_AREA_VIEW.lon], MICHIGAN_AREA_VIEW.zoom);
         }

         if (mapViewContainer) mapViewContainer.style.display = 'none';
         if (mapStatus) mapStatus.textContent = 'Enable via "Additional Tools" and apply filters to see map.';

         if (attachRateTableBody) attachRateTableBody.innerHTML = ''; 
         if (attachRateTableFooter) attachRateTableFooter.innerHTML = ''; 
         if (attachTableStatus) attachTableStatus.textContent = '';
         if (topBottomSection) topBottomSection.style.display = 'none'; 
         if (top5TableBody) top5TableBody.innerHTML = ''; 
         if (bottom5TableBody) bottom5TableBody.innerHTML = '';
         hideStoreDetails(); 
         updateSummary([]); 
        if (eliteOpportunitiesSection) eliteOpportunitiesSection.style.display = 'none';
        if (connectivityOpportunitiesSection) connectivityOpportunitiesSection.style.display = 'none';
        if (repSkillOpportunitiesSection) repSkillOpportunitiesSection.style.display = 'none';
        if (vpmrOpportunitiesSection) vpmrOpportunitiesSection.style.display = 'none';
        if (emailShareSection) emailShareSection.style.display = 'none';

         if(statusDiv) statusDiv.textContent = 'No file selected.';
         allPossibleStores = []; 
         rawData = []; 
         filteredData = [];
         updateShareOptions(); 
         const mainFilterBtn = document.getElementById('openFilterModalBtn');
         if (mainFilterBtn) mainFilterBtn.disabled = true;
         if (clearDefaultFiltersBtn) clearDefaultFiltersBtn.disabled = localStorage.getItem(DEFAULT_FILTERS_STORAGE_KEY) === null;
         closeFilterModal(); 
     };

    const handleResetFiltersClick = (isFromModal = false) => { 
        [regionFilter, districtFilter, fsmFilter, channelFilter, subchannelFilter, dealerFilter].forEach(sel => { 
            if (sel) sel.value = 'ALL'; 
        });
        if (territoryFilter) territoryFilter.selectedIndex = -1;
        if (storeFilter) storeFilter.selectedIndex = -1; 
        if (storeSearch) storeSearch.value = ''; 
        Object.values(flagFiltersCheckboxes).forEach(input => { if(input) input.checked = false; });
        
        if(showMapViewFilter) showMapViewFilter.checked = false;
        if(focusEliteFilter) focusEliteFilter.checked = false;
        if(focusConnectivityFilter) focusConnectivityFilter.checked = false;
        if(focusRepSkillFilter) focusRepSkillFilter.checked = false;
        if(focusVpmrFilter) focusVpmrFilter.checked = false;

        if (rawData.length > 0) {
            updateStoreFilterOptionsBasedOnHierarchy(); 
        } else {
            if (storeFilter) { storeFilter.innerHTML = '<option value="ALL">-- Load File First --</option>'; storeFilter.disabled = true; }
            if (storeSearch) storeSearch.disabled = true;
            if (storeSelectAll) storeSelectAll.disabled = true;
            if (storeDeselectAll) storeDeselectAll.disabled = true;
        }
        
        if (statusDiv) {
            if (rawData.length > 0) {
                statusDiv.textContent = 'Session filters reset. Click "Apply Filters" to see results, or reload file to use defaults.';
            } else {
                statusDiv.textContent = 'No file selected. Load a file to use filters.';
            }
        }
        if (isFromModal) {
             applyFilters(true); 
        }
    };

    if (applyFiltersButtonModal) {
        applyFiltersButtonModal.addEventListener('click', () => applyFilters(true)); 
    }
    if (resetFiltersButtonModal) {
        resetFiltersButtonModal.addEventListener('click', () => handleResetFiltersClick(true)); 
    }
    if (saveDefaultFiltersBtn) {
        saveDefaultFiltersBtn.addEventListener('click', saveDefaultFilters);
    }
    if (clearDefaultFiltersBtn) {
        clearDefaultFiltersBtn.addEventListener('click', clearDefaultFilters);
    }


    const updateSummary = (data) => {
        const totalCount = data.length;
        const fieldsToClearText = [revenueWithDFValue, qtdRevenueTargetValue, qtdGapValue, quarterlyRevenueTargetValue, percentQuarterlyStoreTargetValue, revARValue, unitsWithDFValue, unitTargetValue, unitAchievementValue, visitCountValue, trainingCountValue, retailModeConnectivityValue, repSkillAchValue, vPmrAchValue, postTrainingScoreValue, eliteValue, percentQuarterlyTerritoryTargetValue, territoryRevPercentValue, districtRevPercentValue, regionRevPercentValue];
        fieldsToClearText.forEach(el => { if (el) el.textContent = 'N/A'; });
        [percentQuarterlyTerritoryTargetP, territoryRevPercentP, districtRevPercentP, regionRevPercentP].forEach(p => { if(p) p.style.display = 'none';});
        if (totalCount === 0) return;
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
            let valStr; const subChannel = safeGet(row, 'SUB_CHANNEL', null); 
            valStr = safeGet(row, 'Retail Mode Connectivity', null); if (isValidForAverage(valStr)) { sumConnectivity += parsePercent(valStr); countConnectivity++; }
            valStr = safeGet(row, 'Rep Skill Ach', null); if (isValidForAverage(valStr)) { sumRepSkill += parsePercent(valStr); countRepSkill++; }
            valStr = safeGet(row, '(V)PMR Ach', null); if (isValidForAverage(valStr)) { sumPmr += parsePercent(valStr); countPmr++; }
            valStr = safeGet(row, 'Post Training Score', null); 
            if (isValidForAverage(valStr)) { 
                const numericScore = parseNumber(valStr); 
                if (numericScore !== 0) { sumPostTraining += numericScore; countPostTraining++; }
            }
            if (subChannel !== "Verizon COR") { valStr = safeGet(row, 'Elite', null); if (isValidForAverage(valStr)) { sumElite += parsePercent(valStr); countElite++; } }
        });
        const calculatedRevAR = sumQtdTarget === 0 ? NaN : sumRevenue / sumQtdTarget;
        const avgConnectivity = countConnectivity > 0 ? sumConnectivity / countConnectivity : NaN; const avgRepSkill = countRepSkill > 0 ? sumRepSkill / countRepSkill : NaN;
        const avgPmr = countPmr > 0 ? sumPmr / countPmr : NaN; const avgPostTraining = countPostTraining > 0 ? sumPostTraining / countPostTraining : NaN;
        const avgElite = countElite > 0 ? sumElite / countElite : NaN; 
        const overallPercentStoreTarget = sumQuarterlyTarget !== 0 ? sumRevenue / sumQuarterlyTarget : NaN; const overallUnitAchievement = sumUnitTarget !== 0 ? sumUnits / sumUnitTarget : NaN;
        if (revenueWithDFValue) { revenueWithDFValue.textContent = formatCurrency(sumRevenue); revenueWithDFValue.title = `Sum of 'Revenue w/DF' for ${totalCount} filtered stores`; }
        if (qtdRevenueTargetValue) { qtdRevenueTargetValue.textContent = formatCurrency(sumQtdTarget); qtdRevenueTargetValue.title = `Sum of 'QTD Revenue Target' for ${totalCount} filtered stores`; }
        if (qtdGapValue) { qtdGapValue.textContent = formatCurrency(sumRevenue - sumQtdTarget); qtdGapValue.title = `Calculated Gap (Total Revenue - QTD Target) for ${totalCount} filtered stores`; }
        if (quarterlyRevenueTargetValue) { quarterlyRevenueTargetValue.textContent = formatCurrency(sumQuarterlyTarget); quarterlyRevenueTargetValue.title = `Sum of 'Quarterly Revenue Target' for ${totalCount} filtered stores`; }
        if (unitsWithDFValue) { unitsWithDFValue.textContent = formatNumber(sumUnits); unitsWithDFValue.title = `Sum of 'Unit w/ DF' for ${totalCount} filtered stores`; }
        if (unitTargetValue) { unitTargetValue.textContent = formatNumber(sumUnitTarget); unitTargetValue.title = `Sum of 'Unit Target' for ${totalCount} filtered stores`; }
        if (visitCountValue) { visitCountValue.textContent = formatNumber(sumVisits); visitCountValue.title = `Sum of 'Visit count' for ${totalCount} filtered stores`; }
        if (trainingCountValue) { trainingCountValue.textContent = formatNumber(sumTrainings); trainingCountValue.title = `Sum of 'Trainings' for ${totalCount} filtered stores`; }
        if (revARValue) { revARValue.textContent = formatPercent(calculatedRevAR); revARValue.title = "Rev AR% for selected stores with data"; }
        if (percentQuarterlyStoreTargetValue) { percentQuarterlyStoreTargetValue.textContent = formatPercent(overallPercentStoreTarget); percentQuarterlyStoreTargetValue.title = `Overall % Quarterly Target (Total Revenue / Total Quarterly Target)`; }
        if (unitAchievementValue) { unitAchievementValue.textContent = formatPercent(overallUnitAchievement); unitAchievementValue.title = `Overall Unit Achievement % (Total Units / Total Unit Target)`; }
        if (retailModeConnectivityValue) { retailModeConnectivityValue.textContent = formatPercent(avgConnectivity); retailModeConnectivityValue.title = `Average 'Retail Mode Connectivity' across ${countConnectivity} stores with data`; }
        if (repSkillAchValue) { repSkillAchValue.textContent = formatPercent(avgRepSkill); repSkillAchValue.title = `Average 'Rep Skill Ach' across ${countRepSkill} stores with data`; }
        if (vPmrAchValue) { vPmrAchValue.textContent = formatPercent(avgPmr); vPmrAchValue.title = `Average '(V)PMR Ach' across ${countPmr} stores with data`; }
        if (postTrainingScoreValue) { postTrainingScoreValue.textContent = isNaN(avgPostTraining) ? 'N/A' : avgPostTraining.toFixed(1); postTrainingScoreValue.title = `Average 'Post Training Score' across ${countPostTraining} stores with data (excluding 0s)`; }
        if (eliteValue) { eliteValue.textContent = formatPercent(avgElite); eliteValue.title = `Average 'Elite' score % across ${countElite} stores with data (excluding Verizon COR sub-channel)`;}
        updateContextualSummary(data);
    };
    const updateContextualSummary = (data) => {
        [percentQuarterlyTerritoryTargetP, territoryRevPercentP, districtRevPercentP, regionRevPercentP].forEach(p => {if (p) p.style.display = 'none'});
        if (data.length === 0) return;
        const singleRegion = regionFilter?.value !== 'ALL'; const singleDistrict = districtFilter?.value !== 'ALL'; const singleTerritory = territoryFilter && territoryFilter.selectedOptions.length === 1;
        const calculateAverageExcludeBlanks = (column) => { let sum = 0, count = 0; data.forEach(row => { const valStr = safeGet(row, column, null); if (isValidForAverage(valStr)) { sum += parsePercent(valStr); count++; } }); return count > 0 ? sum / count : NaN; };
        const avgPercentTerritoryTarget = calculateAverageExcludeBlanks('%Quarterly Territory Rev Target'); const avgTerritoryRevPercent = calculateAverageExcludeBlanks('Territory Rev%');
        const avgDistrictRevPercent = calculateAverageExcludeBlanks('District Rev%'); const avgRegionRevPercent = calculateAverageExcludeBlanks('Region Rev%');
        if (percentQuarterlyTerritoryTargetValue) percentQuarterlyTerritoryTargetValue.textContent = formatPercent(avgPercentTerritoryTarget); if (percentQuarterlyTerritoryTargetP && !isNaN(avgPercentTerritoryTarget)) percentQuarterlyTerritoryTargetP.style.display = 'block';
        if (singleTerritory || singleDistrict || singleRegion) { if (territoryRevPercentValue) territoryRevPercentValue.textContent = formatPercent(avgTerritoryRevPercent); if (territoryRevPercentP && !isNaN(avgTerritoryRevPercent)) territoryRevPercentP.style.display = 'block'; }
        if (singleDistrict || singleRegion) { if (districtRevPercentValue) districtRevPercentValue.textContent = formatPercent(avgDistrictRevPercent); if (districtRevPercentP && !isNaN(avgDistrictRevPercent)) districtRevPercentP.style.display = 'block'; }
        if (singleRegion) { if (regionRevPercentValue) regionRevPercentValue.textContent = formatPercent(avgRegionRevPercent); if (regionRevPercentP && !isNaN(avgRegionRevPercent)) regionRevPercentP.style.display = 'block'; }
    };
    const updateCharts = (data) => {
        if (mainChartInstance) { mainChartInstance.destroy(); mainChartInstance = null; }
        if (!mainChartCanvas || (data.length === 0 && rawData.length === 0) ) { if (mainChartCanvas && mainChartInstance) { mainChartInstance = new Chart(mainChartCanvas, {type: 'bar', data: {labels:[], datasets:[]}}); mainChartInstance.destroy(); mainChartInstance = null;} return; }
        const chartThemeColors = getChartThemeColors(); 
        const sortedData = [...data].sort((a, b) => parseNumber(safeGet(b, 'Revenue w/DF', 0)) - parseNumber(safeGet(a, 'Revenue w/DF', 0)));
        const chartData = sortedData.slice(0, TOP_N_CHART);
        const labels = chartData.map(row => safeGet(row, 'Store', 'Unknown Store'));
        const revenueDataSet = chartData.map(row => parseNumber(safeGet(row, 'Revenue w/DF', 0)));
        const targetDataSet = chartData.map(row => parseNumber(safeGet(row, 'QTD Revenue Target', 0)));
        const backgroundColors = chartData.map((_, index) => revenueDataSet[index] >= targetDataSet[index] ? 'rgba(75, 192, 192, 0.6)' : 'rgba(255, 99, 132, 0.6)'); 
        const borderColors = chartData.map((_, index) => revenueDataSet[index] >= targetDataSet[index] ? 'rgba(75, 192, 192, 1)' : 'rgba(255, 99, 132, 1)');
        mainChartInstance = new Chart(mainChartCanvas, {
            type: 'bar', data: { labels: labels, datasets: [ { label: 'Total Revenue (incl. DF)', data: revenueDataSet, backgroundColor: backgroundColors, borderColor: borderColors, borderWidth: 1 }, { label: 'QTD Revenue Target', data: targetDataSet, type: 'line', borderColor: 'rgba(255, 206, 86, 1)', backgroundColor: 'rgba(255, 206, 86, 0.2)', fill: false, tension: 0.1, borderWidth: 2, pointRadius: 3, pointHoverRadius: 5 } ] },
            options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, ticks: { color: chartThemeColors.tickColor, callback: value => formatCurrency(value) }, grid: { color: chartThemeColors.gridColor } }, x: { ticks: { color: chartThemeColors.tickColor }, grid: { display: false } } }, plugins: { legend: { labels: { color: chartThemeColors.legendColor } }, tooltip: { callbacks: { label: function(context) { let label = context.dataset.label || ''; if (label) label += ': '; if (context.parsed.y !== null) { label += formatCurrency(context.parsed.y); if (context.dataset.type !== 'line' && chartData[context.dataIndex]) { const storeData = chartData[context.dataIndex]; const percentQtrTarget = parsePercent(safeGet(storeData, '% Quarterly Revenue Target', 0)); if (!isNaN(percentQtrTarget)) { label += ` (${formatPercent(percentQtrTarget)} of Qtr Target)`; } } } return label; } } } }, onClick: (_, elements) => { if (elements.length > 0) { const index = elements[0].index; const storeName = labels[index]; const storeData = filteredData.find(row => safeGet(row, 'Store', null) === storeName); if (storeData) { showStoreDetails(storeData); highlightTableRow(storeName); } } } }
        });
    };
    const updateAttachRateTable = (data) => {
        if (!attachRateTableBody || !attachRateTableFooter || !attachRateTable) return;
        attachRateTableBody.innerHTML = ''; attachRateTableFooter.innerHTML = '';
        const dataForTable = data.filter(row => ATTACH_RATE_COLUMNS.every(colKey => isValidNumericForFocus(safeGet(row, colKey, null))));
        const tableHead = attachRateTable.querySelector('thead');
        if (!tableHead) { console.error("Attach rate table head not found!"); return; }
        let headerRow = tableHead.querySelector('tr');
        if (!headerRow) { headerRow = tableHead.insertRow(); }
        headerRow.innerHTML = ''; 
        if (dataForTable.length === 0) {
            if (attachTableStatus) attachTableStatus.textContent = 'No stores with complete & valid attach rate data based on current filters.';
            return;
        }
        const uniqueTerritories = new Set(dataForTable.map(row => safeGet(row, 'Q2 Territory', 'N/A_TERRITORY')));
        const showTerritoryColumn = uniqueTerritories.size > 1;
        const baseHeaderConfig = [
            { label: 'Store', sortKey: 'Store', title: 'Store Name' },
            { label: 'Tablet', sortKey: 'Tablet Attach Rate', title: 'Tablet Attach Rate' },
            { label: 'PC', sortKey: 'PC Attach Rate', title: 'PC Attach Rate' },
            { label: 'NC', sortKey: 'NC Attach Rate', title: 'NC = Tablet + PC Attach Rate' },
            { label: 'TWS', sortKey: 'TWS Attach Rate', title: 'True Wireless Stereo (Buds) Attach Rate' },
            { label: 'WW', sortKey: 'WW Attach Rate', title: 'Wearable Watch Attach Rate' },
            { label: 'ME', sortKey: 'ME Attach Rate', title: 'ME = TWS + WW Attach Rate' },
            { label: 'NCME', sortKey: 'NCME Attach Rate', title: 'NCME = Total Attach Rate' }
        ];
        let actualTableHeaders = [...baseHeaderConfig];
        if (showTerritoryColumn) {
            actualTableHeaders.splice(1, 0, { label: 'Territory', sortKey: 'Q2 Territory', title: 'Q2 Territory' });
        }
        actualTableHeaders.forEach(headerInfo => {
            const th = document.createElement('th');
            th.textContent = headerInfo.label; th.dataset.sort = headerInfo.sortKey; th.title = headerInfo.title;
            th.classList.add('sortable'); th.innerHTML += ' <span class="sort-arrow"></span>';
            headerRow.appendChild(th);
        });
        const sortedData = [...dataForTable].sort((a, b) => {
            let valA = safeGet(a, currentSort.column, null); let valB = safeGet(b, currentSort.column, null);
            if (valA === null && valB === null) return 0; if (valA === null) return currentSort.ascending ? -1 : 1;
            if (valB === null) return currentSort.ascending ? 1 : -1;
            const isAttachRateKey = ATTACH_RATE_COLUMNS.includes(currentSort.column);
            const isStoreOrTerritoryKey = currentSort.column === 'Store' || currentSort.column === 'Q2 Territory';
            let numA, numB;
            if(isAttachRateKey) { numA = parsePercent(valA); numB = parsePercent(valB); } 
            else if (!isStoreOrTerritoryKey) { numA = parseNumber(valA); numB = parseNumber(valB); } 
            else { valA = String(valA).toLowerCase(); valB = String(valB).toLowerCase(); return currentSort.ascending ? valA.localeCompare(valB) : valB.localeCompare(valA); }
            if (typeof numA === 'number' && typeof numB === 'number' && !isNaN(numA) && !isNaN(numB)) { return currentSort.ascending ? numA - numB : numB - numA; }
            valA = String(valA).toLowerCase(); valB = String(valB).toLowerCase();
            return currentSort.ascending ? valA.localeCompare(valB) : valB.localeCompare(valA);
        });
        const averages = {};
        ATTACH_RATE_COLUMNS.forEach(key => {
            let sum = 0, count = 0;
            dataForTable.forEach(row => { const valStr = safeGet(row, key, null); if (isValidNumericForFocus(valStr)) { sum += parsePercent(valStr); count++; } });
            averages[key] = count > 0 ? sum / count : NaN;
        });
        sortedData.forEach(row => {
            const tr = attachRateTableBody.insertRow(); const storeName = safeGet(row, 'Store', 'N/A');
            tr.dataset.storeName = storeName; tr.onclick = () => { showStoreDetails(row); highlightTableRow(storeName); };
            actualTableHeaders.forEach(headerInfo => {
                const td = tr.insertCell(); let cellValue; let rawValueForMetric = safeGet(row, headerInfo.sortKey, null);
                if (headerInfo.sortKey === 'Store') { cellValue = storeName; } 
                else if (headerInfo.sortKey === 'Q2 Territory') { cellValue = safeGet(row, 'Q2 Territory', 'N/A'); } 
                else { 
                    const numericValue = parsePercent(rawValueForMetric); cellValue = isNaN(numericValue) ? 'N/A' : formatPercent(numericValue); td.style.textAlign = "right";
                    td.classList.remove('highlight-green', 'highlight-red', 'highlight-yellow'); 
                    if (!isNaN(averages[headerInfo.sortKey]) && typeof numericValue === 'number' && !isNaN(numericValue)) {
                        const avg = averages[headerInfo.sortKey]; const lowerBound = avg * (1 - AVERAGE_THRESHOLD_PERCENT); const upperBound = avg * (1 + AVERAGE_THRESHOLD_PERCENT);
                        if (numericValue > upperBound) td.classList.add('highlight-green');
                        else if (numericValue < lowerBound) td.classList.add('highlight-red');
                        else td.classList.add('highlight-yellow');
                    }
                }
                td.textContent = cellValue; td.title = cellValue; 
            });
        });
        if (dataForTable.length > 0) {
            const footerRowNew = attachRateTableFooter.insertRow();
            actualTableHeaders.forEach((headerInfo, index) => {
                const td = footerRowNew.insertCell();
                 if (index === 0) { td.textContent = 'Filtered Avg*'; td.style.fontWeight = "bold"; td.title = 'Average calculated only using stores with complete and valid attach rate data'; } 
                 else if (showTerritoryColumn && index === 1 && headerInfo.sortKey === 'Q2 Territory') { td.textContent = ''; } 
                 else if (ATTACH_RATE_COLUMNS.includes(headerInfo.sortKey)) {
                    const avgValue = averages[headerInfo.sortKey]; td.textContent = formatPercent(avgValue);
                    let validCount = dataForTable.filter(r => isValidNumericForFocus(safeGet(r, headerInfo.sortKey, null))).length;
                    td.title = `Average ${headerInfo.label}: ${formatPercent(avgValue)} (from ${validCount} stores)`; td.style.textAlign = "right";
                }
            });
        }
        if (attachTableStatus) attachTableStatus.textContent = `Showing ${attachRateTableBody.rows.length} stores with complete attach rate data. Click row for details. Click headers to sort.`;
        updateSortArrows();
    };

    const updateFocusTableSortArrows = (tableElement) => {
        if (!tableElement || !tableElement.sortConfig) return;
        const sortConfig = tableElement.sortConfig;
        tableElement.querySelectorAll('thead th.sortable .sort-arrow').forEach(arrow => {
            arrow.className = 'sort-arrow'; 
            arrow.textContent = '';
        });
        const currentHeader = tableElement.querySelector(`thead th[data-sort="${CSS.escape(sortConfig.column)}"]`);
        if (currentHeader) {
            const arrowSpan = currentHeader.querySelector('.sort-arrow');
            if (arrowSpan) {
                arrowSpan.classList.add(sortConfig.ascending ? 'asc' : 'desc');
            }
        }
    };
    
    const populateFocusPointTable = (tableId, sectionElement, data, valueKey, valueLabel) => {
        const table = document.getElementById(tableId);
        if (!table) { console.error(`Table with ID ${tableId} not found.`); return; }
    
        const thead = table.querySelector('thead');
        const tbody = table.querySelector('tbody');
        if (!thead || !tbody) { console.error(`Thead or Tbody not found for table ${tableId}.`); return; }
    
        thead.innerHTML = ''; 
        const headerRow = thead.insertRow();
        const headers = [
            { label: 'Store', sortKey: 'Store' },
            { label: 'Territory', sortKey: 'Q2 Territory' },
            { label: valueLabel, sortKey: valueKey }
        ];
    
        headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header.label;
            th.dataset.sort = header.sortKey;
            th.classList.add('sortable');
            th.innerHTML += ' <span class="sort-arrow"></span>';
            headerRow.appendChild(th);
        });
    
        if (!table.sortConfig) {
            table.sortConfig = { column: 'Store', ascending: true }; 
        }
    
        const sortConfig = table.sortConfig;
        const sortedData = [...data].sort((a, b) => {
            let valA = safeGet(a, sortConfig.column, null);
            let valB = safeGet(b, sortConfig.column, null);
    
            if (valA === null || String(valA).trim() === '') valA = sortConfig.ascending ? Infinity : -Infinity;
            if (valB === null || String(valB).trim() === '') valB = sortConfig.ascending ? Infinity : -Infinity;
    
            let numA, numB;
            if (sortConfig.column === valueKey) { 
                numA = parsePercent(String(valA).replace('%',''));
                numB = parsePercent(String(valB).replace('%',''));
            } else if (sortConfig.column === 'Store' || sortConfig.column === 'Q2 Territory') {
                valA = String(valA).toLowerCase();
                valB = String(valB).toLowerCase();
                return sortConfig.ascending ? valA.localeCompare(valB) : valB.localeCompare(valA);
            } else { 
                numA = parseNumber(valA);
                numB = parseNumber(valB);
            }
    
            if (typeof numA === 'number' && typeof numB === 'number' && !isNaN(numA) && !isNaN(numB)) {
                return sortConfig.ascending ? numA - numB : numB - numA;
            }
            valA = String(valA).toLowerCase();
            valB = String(valB).toLowerCase();
            return sortConfig.ascending ? valA.localeCompare(valB) : valB.localeCompare(valA);
        });
    
        tbody.innerHTML = '';
        sortedData.forEach(row => {
            const tr = tbody.insertRow();
            const storeName = safeGet(row, 'Store', 'N/A');
            tr.dataset.storeName = storeName;
            tr.onclick = () => { showStoreDetails(row); highlightTableRow(storeName); };
    
            tr.insertCell().textContent = storeName;
            tr.cells[0].title = storeName;
    
            const territoryName = safeGet(row, 'Q2 Territory', 'N/A');
            tr.insertCell().textContent = territoryName;
            tr.cells[1].title = territoryName;
    
            const metricValue = parsePercent(safeGet(row, valueKey, NaN));
            const cellMetric = tr.insertCell();
            cellMetric.textContent = formatPercent(metricValue);
            cellMetric.title = formatPercent(metricValue);
            cellMetric.style.textAlign = "right";
        });
    
        updateFocusTableSortArrows(table);
    
        const newThead = thead.cloneNode(true); 
        thead.parentNode.replaceChild(newThead, thead);
        newThead.addEventListener('click', (event) => {
            const headerCell = event.target.closest('th');
            if (!headerCell || !headerCell.classList.contains('sortable')) return;
            
            const sortKey = headerCell.dataset.sort;
            if (!sortKey) return;
    
            if (table.sortConfig.column === sortKey) {
                table.sortConfig.ascending = !table.sortConfig.ascending;
            } else {
                table.sortConfig.column = sortKey;
                table.sortConfig.ascending = true;
            }
            populateFocusPointTable(tableId, sectionElement, data, valueKey, valueLabel); 
        });
    
        const statusP = sectionElement.querySelector('.focus-point-status');
        if (statusP) {
            if (data.length > 0) {
                statusP.textContent = `Displaying ${data.length} stores. Click headers to sort.`;
            } else {
                statusP.textContent = 'No stores meet this criteria based on current filters.';
            }
        }
        sectionElement.style.display = 'block';
    };

    const updateFocusPointSections = (baseData) => {
        if (focusEliteFilter?.checked) {
            const eliteOpps = baseData.filter(row => { const eliteVal = parsePercent(safeGet(row, 'Elite', null)); return !isNaN(eliteVal) && eliteVal > 0.01 && eliteVal < 1.0; });
            populateFocusPointTable('eliteOpportunitiesTable', eliteOpportunitiesSection, eliteOpps, 'Elite', 'Elite %');
        } else { if (eliteOpportunitiesSection) eliteOpportunitiesSection.style.display = 'none'; }

        if (focusConnectivityFilter?.checked) {
            const connOpps = baseData.filter(row => { const connVal = parsePercent(safeGet(row, 'Retail Mode Connectivity', null)); return !isNaN(connVal) && connVal < 1.0; });
            populateFocusPointTable('connectivityOpportunitiesTable', connectivityOpportunitiesSection, connOpps, 'Retail Mode Connectivity', 'Connectivity %');
        } else { if (connectivityOpportunitiesSection) connectivityOpportunitiesSection.style.display = 'none'; }
        
        if (focusRepSkillFilter?.checked) {
            const repSkillOpps = baseData.filter(row => { const repSkillVal = parsePercent(safeGet(row, 'Rep Skill Ach', null)); return isValidNumericForFocus(safeGet(row, 'Rep Skill Ach', null)) && repSkillVal < 1.0; });
            populateFocusPointTable('repSkillOpportunitiesTable', repSkillOpportunitiesSection, repSkillOpps, 'Rep Skill Ach', 'Rep Skill Ach %');
        } else { if (repSkillOpportunitiesSection) repSkillOpportunitiesSection.style.display = 'none'; }
        
        if (focusVpmrFilter?.checked) {
            const vpmrOpps = baseData.filter(row => { const vpmrVal = parsePercent(safeGet(row, '(V)PMR Ach', null)); return isValidNumericForFocus(safeGet(row, '(V)PMR Ach', null)) && vpmrVal < 1.0; });
            populateFocusPointTable('vpmrOpportunitiesTable', vpmrOpportunitiesSection, vpmrOpps, '(V)PMR Ach', '(V)PMR Ach %');
        } else { if (vpmrOpportunitiesSection) vpmrOpportunitiesSection.style.display = 'none'; }
    };

    const handleSort = (event) => { 
         const headerCell = event.target.closest('th'); if (!headerCell?.classList.contains('sortable')) return;
         const sortKey = headerCell.dataset.sort; if (!sortKey) return;
         if (currentSort.column === sortKey) { currentSort.ascending = !currentSort.ascending; } else { currentSort.column = sortKey; currentSort.ascending = true; }
         updateAttachRateTable(filteredData); 
    };
    const updateSortArrows = () => { 
        if (!attachRateTable) return;
        attachRateTable.querySelectorAll('thead th.sortable .sort-arrow').forEach(arrow => { arrow.className = 'sort-arrow'; arrow.textContent = ''; });
        const currentHeaderArrow = attachRateTable.querySelector(`thead th[data-sort="${CSS.escape(currentSort.column)}"] .sort-arrow`);
        if (currentHeaderArrow) { currentHeaderArrow.classList.add(currentSort.ascending ? 'asc' : 'desc'); }
    };
    const showStoreDetails = (storeData) => {
        if (!storeDetailsContent || !storeDetailsSection || !closeStoreDetailsButton) return;
        const addressParts = [ safeGet(storeData, 'ADDRESS1', null), safeGet(storeData, 'CITY', null), safeGet(storeData, 'STATE', null), safeGet(storeData, 'ZIPCODE', null) ].filter(part => part && String(part).trim() !== '');
        const addressString = addressParts.length > 0 ? addressParts.join(', ') : 'N/A';
        const latitude = parseNumber(safeGet(storeData, 'LATITUDE_ORG', NaN)); const longitude = parseNumber(safeGet(storeData, 'LONGITUDE_ORG', NaN));
        let mapsLinkHtml = `<p style="color: #aaa; font-style: italic;">(Map coordinates not available)</p>`;
        if (!isNaN(latitude) && !isNaN(longitude) && latitude !== 0 && longitude !==0) { const mapsUrl = `https://www.google.com/maps/search/?api=1&query=${latitude},${longitude}`; mapsLinkHtml = `<p><a href="${mapsUrl}" target="_blank" title="Open in Google Maps">View on Google Maps</a></p>`; }
        let flagSummaryHtml = FLAG_HEADERS.map(flag => { const flagValue = safeGet(storeData, flag, 'NO'); const isTrue = (flagValue === true || String(flagValue).toUpperCase() === 'YES' || String(flagValue) === 'Y' || flagValue === 1 || String(flagValue) === '1'); return `<span title="${flag.replace(/_/g, ' ')}" data-flag="${isTrue}">${flag.replace(/_/g, ' ')} ${isTrue ? '✔' : '✘'}</span>`; }).join(' | ');
        storeDetailsContent.innerHTML = ` <p><strong>Store:</strong> ${safeGet(storeData, 'Store')}</p> <p><strong>Address:</strong> ${addressString}</p> ${mapsLinkHtml} <hr> <p><strong>IDs:</strong> Store: ${safeGet(storeData, 'STORE ID')} | Org: ${safeGet(storeData, 'ORG_STORE_ID')} | CV: ${safeGet(storeData, 'CV_STORE_ID')} | CinglePoint: ${safeGet(storeData, 'CINGLEPOINT_ID')}</p> <p><strong>Type:</strong> ${safeGet(storeData, 'STORE_TYPE_NAME')} | Nat Tier: ${safeGet(storeData, 'National_Tier')} | Merch Lvl: ${safeGet(storeData, 'Merchandising_Level')} | Comb Tier: ${safeGet(storeData, 'Combined_Tier')}</p> <hr> <p><strong>Hierarchy:</strong> ${safeGet(storeData, 'REGION')} > ${safeGet(storeData, 'DISTRICT')} > ${safeGet(storeData, 'Q2 Territory')}</p> <p><strong>FSM:</strong> ${safeGet(storeData, 'FSM NAME')}</p> <p><strong>Channel:</strong> ${safeGet(storeData, 'CHANNEL')} / ${safeGet(storeData, 'SUB_CHANNEL')}</p> <p><strong>Dealer:</strong> ${safeGet(storeData, 'DEALER_NAME')}</p> <hr> <p><strong>Visits:</strong> ${formatNumber(parseNumber(safeGet(storeData, 'Visit count', 0)))} | <strong>Trainings:</strong> ${formatNumber(parseNumber(safeGet(storeData, 'Trainings', 0)))}</p> <p><strong>Connectivity:</strong> ${formatPercent(parsePercent(safeGet(storeData, 'Retail Mode Connectivity', 0)))}</p> <hr> <p><strong>Flags:</strong> ${flagSummaryHtml}</p> `;
        storeDetailsSection.style.display = 'block'; closeStoreDetailsButton.style.display = 'inline-block';
    };
    const hideStoreDetails = () => {
        if (!storeDetailsContent || !storeDetailsSection || !closeStoreDetailsButton) return;
        storeDetailsContent.innerHTML = 'Select a store from the table or chart for details, or apply filters resulting in a single store.';
        storeDetailsSection.style.display = 'none'; closeStoreDetailsButton.style.display = 'none'; highlightTableRow(null);
    };
    const highlightTableRow = (storeName) => {
        if (selectedStoreRow) { selectedStoreRow.classList.remove('selected-row'); selectedStoreRow = null; }
        if (storeName) {
            const tablesToSearch = [ 
                document.getElementById('attachRateTableBody'), 
                document.getElementById('top5TableBody'), 
                document.getElementById('bottom5TableBody'), 
                document.getElementById('eliteOpportunitiesTableBody'), 
                document.getElementById('connectivityOpportunitiesTableBody'), 
                document.getElementById('repSkillOpportunitiesTableBody'), 
                document.getElementById('vpmrOpportunitiesTableBody') 
            ];
            for (const tableBody of tablesToSearch) {
                if (tableBody) { 
                    try { const rowToHighlight = tableBody.querySelector(`tr[data-store-name="${CSS.escape(storeName)}"]`); if (rowToHighlight) { rowToHighlight.classList.add('selected-row'); selectedStoreRow = rowToHighlight; break; }
                    } catch (e) { console.error("Error selecting table row:", e, "StoreName:", storeName); }
                }
            }
        }
    };
    const updateShareOptions = () => {
        if (!emailShareSection || !shareEmailButton || !emailShareHint || !printReportButton) return;
        if (filteredData.length === 0) { printReportButton.disabled = true; emailShareSection.style.display = 'none'; return; }
        printReportButton.disabled = false; const emailBody = generateEmailBody(); 
        if (emailBody.length < 2000) {
            emailShareSection.style.display = 'block'; if(emailShareControls) emailShareControls.style.display = 'flex'; 
            shareEmailButton.disabled = false;
            emailShareHint.textContent = 'Note: This will open your default desktop email client. Available for summaries under 2000 characters.';
        } else {
            emailShareSection.style.display = 'block'; if(emailShareControls) emailShareControls.style.display = 'none';
            shareEmailButton.disabled = true;
            emailShareHint.textContent = 'Email summary is too long for direct sharing. Please use "Print / Save as PDF" for a full report.';
        }
    };
    const generateReportHTML = () => {
        let html = `<h1>FSM Performance Dashboard Report</h1><p><strong>Report Generated:</strong> ${new Date().toLocaleString()}</p><div class="report-filter-summary"><p><strong>Filters Applied:</strong> ${getFilterSummary() || 'None'}</p></div>`;
        const summaryDataEl = document.getElementById('summaryData');
        if (summaryDataEl) {
            html += `<h2>Performance Summary</h2>`; const summaryGrid = summaryDataEl.querySelector('.summary-grid');
            if (summaryGrid) {
                html += '<ul style="list-style-type: none; padding-left: 0;">';
                summaryGrid.querySelectorAll('p').forEach(p => {
                    if (p.style.display !== 'none') {
                         const labelNode = p.childNodes[0]; const label = labelNode ? labelNode.nodeValue.trim() : '';
                         const value = p.querySelector('strong')?.textContent || '';
                         if(label) html += `<li style="margin-bottom: 5px;">${label} <strong>${value}</strong></li>`;
                    }
                }); html += '</ul>';
            }
        }
        if (mainChartInstance) {
            try { const chartImageURL = mainChartInstance.toBase64Image(); html += `<h2>Revenue Performance Chart</h2><div class="report-chart-container"><img src="${chartImageURL}" alt="Revenue Performance Chart" style="max-width: 100%; height: auto; border: 1px solid #ddd;"></div>`;
            } catch(e) { html += `<p><em>Main chart could not be included.</em></p>`; }
        }
        const generateTableHTMLFromDOM = (tableElementId, title) => {
            const tableElement = document.getElementById(tableElementId);
            let parentSection = tableElement ? tableElement.closest('.card') : null;
            if (tableElementId === 'top5Table' || tableElementId === 'bottom5Table') parentSection = document.getElementById('topBottomSection');
            else if (tableElementId.includes('OpportunitiesTable')) parentSection = document.getElementById(tableElementId)?.closest('.focus-point-card');
            else if (tableElementId === 'attachRateTable') parentSection = document.getElementById('attachRateTableContainer');
            if (!tableElement || (parentSection && parentSection.style.display === 'none') ) return '';
            const tableBody = tableElement.querySelector('tbody');
            if (!tableBody || tableBody.children.length === 0) return '';
            let tableHTML = `<h2>${title}</h2><table>`; const header = tableElement.querySelector('thead');
            if (header) tableHTML += `<thead>${header.innerHTML}</thead>`; 
            tableHTML += `<tbody>`;
            Array.from(tableBody.rows).forEach(row => {
                tableHTML += `<tr>`; Array.from(row.cells).forEach(cell => { tableHTML += `<td style="text-align: ${cell.style.textAlign || 'left'};">${cell.textContent}</td>`; });
                tableHTML += `</tr>`;
            }); tableHTML += `</tbody>`;
            const footer = tableElement.querySelector('tfoot');
            if (footer && footer.innerHTML.trim() !== '') tableHTML += `<tfoot>${footer.innerHTML}</tfoot>`;
            tableHTML += `</table>`; return tableHTML;
        };
        if (topBottomSection && topBottomSection.style.display !== 'none') {
            html += generateTableHTMLFromDOM('top5Table', 'Top 5 (Revenue)');
            html += generateTableHTMLFromDOM('bottom5Table', 'Bottom 5 (Opportunities by QTD Gap)');
        }
        html += generateTableHTMLFromDOM('attachRateTable', 'Attach Rates');
        const focusSections = [
            { id: 'eliteOpportunitiesTable', title: 'Elite Opportunities (>1% <100%)', sectionId: 'eliteOpportunitiesSection' },
            { id: 'connectivityOpportunitiesTable', title: 'Connectivity Opportunities (<100%)', sectionId: 'connectivityOpportunitiesSection' },
            { id: 'repSkillOpportunitiesTable', title: 'Rep Skill Opportunities (<100%, valid data)', sectionId: 'repSkillOpportunitiesSection' },
            { id: 'vpmrOpportunitiesTable', title: 'VPMR Opportunities (<100%, valid data)', sectionId: 'vpmrOpportunitiesSection' } ];
        focusSections.forEach(fs => { html += generateTableHTMLFromDOM(fs.id, fs.title); });
        return html;
    };
    const handlePrintReport = () => {
        if (filteredData.length === 0) { alert("No data to generate a report. Please apply filters first."); return; }
        const reportHTML = generateReportHTML(); const reportWindow = window.open('', '_blank');
        if (!reportWindow) { alert("Please allow popups for this site to generate the report."); return; }
        const printStyles = `body{font-family:Arial,sans-serif;margin:20px;color:#000}h1{font-size:20pt;text-align:center;margin-bottom:10px;color:#000}h2{font-size:16pt;margin-top:25px;margin-bottom:10px;border-bottom:1px solid #ccc;padding-bottom:5px;color:#000;page-break-after:avoid}p,li{font-size:10pt;color:#000}ul{padding-left:20px;margin-top:5px;list-style-type:none}ul li{margin-bottom:5px}.report-filter-summary{margin-bottom:20px;padding:10px;border:1px solid #eee;background-color:#f9f9f9!important;page-break-inside:avoid}table{width:100%;border-collapse:collapse;margin-bottom:20px;font-size:9pt;page-break-inside:auto}th,td{border:1px solid #ccc!important;padding:6px;text-align:left;color:#000!important;background-color:#fff!important;page-break-inside:avoid}th{background-color:#f2f2f2!important;font-weight:bold}.report-chart-container img{max-width:100%!important;height:auto!important;display:block;margin:15px auto;border:1px solid #eee;page-break-inside:avoid}@media print{body{margin:.5in;font-size:9pt;-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}h1{font-size:18pt}h2{font-size:14pt}table{font-size:8pt}th,td{padding:4px}}`;
        reportWindow.document.write(`<html><head><title>FSM Dashboard Report - ${new Date().toLocaleDateString()}</title><style>${printStyles}</style></head><body>${reportHTML}</body></html>`);
        reportWindow.document.close(); setTimeout(() => { reportWindow.focus(); reportWindow.print(); }, 500);
    };
    const exportData = () => {
        if (filteredData.length === 0) { alert("No filtered data to export."); return; }
        try {
            const currentHeaders = Array.from(attachRateTable.querySelectorAll('thead th')).map(th => th.dataset.sort || th.textContent.replace(/ [▲▼]$/, '').trim());
            const dataForExport = filteredData.filter(row => ATTACH_RATE_COLUMNS.every(colKey => isValidNumericForFocus(safeGet(row, colKey, null)))).map(row => currentHeaders.map(headerKey => { 
                const dataKey = headerKey === 'Territory' ? 'Q2 Territory' : headerKey; let value = safeGet(row, dataKey, ''); 
                const isPercentLike = ATTACH_RATE_COLUMNS.includes(dataKey) || dataKey.includes('%') || dataKey.includes('Ach') || dataKey.includes('Connectivity') || dataKey.includes('Elite');
                if (isPercentLike) { const numVal = parsePercent(value); return isNaN(numVal) ? '' : numVal; } 
                else { const numVal = parseNumber(value); if (!isNaN(numVal) && typeof value !== 'boolean' && String(value).trim() !== '') return numVal; if (typeof value === 'string' && (value.includes(',') || value.includes('"') || value.includes('\n'))) return `"${value.replace(/"/g, '""')}"`; return value; }
            }));
            let csvContent = "data:text/csv;charset=utf-8," + currentHeaders.join(",") + "\n" + dataForExport.map(e => e.join(",")).join("\n");
            const encodedUri = encodeURI(csvContent); const link = document.createElement("a"); link.setAttribute("href", encodedUri); link.setAttribute("download", "fsm_dashboard_export.csv"); document.body.appendChild(link); link.click(); document.body.removeChild(link);
        } catch (error) { console.error("Error exporting CSV:", error); alert("Error generating CSV export. See console for details."); }
    };
    const generateEmailBody = () => {
        if (filteredData.length === 0) return "No data available based on current filters.";
        let body = "FSM Dashboard Summary:\n---------------------------------\n"; body += `Filters Applied: ${getFilterSummary()}\nStores Found: ${filteredData.length}\n---------------------------------\n\n`;
        body += "Performance Summary:\n"; body += `- Total Revenue (incl. DF): ${revenueWithDFValue?.textContent || 'N/A'}\n`; body += `- QTD Revenue Target: ${qtdRevenueTargetValue?.textContent || 'N/A'}\n`;
        body += `- QTD Gap: ${qtdGapValue?.textContent || 'N/A'}\n`; body += `- Rev AR%: ${revARValue?.textContent || 'N/A'} (Calculated as: Total Revenue w/DF / Total QTD Revenue Target)\n`;
        body += `- % Store Quarterly Target: ${percentQuarterlyStoreTargetValue?.textContent || 'N/A'}\n`; body += `- Total Units (incl. DF): ${unitsWithDFValue?.textContent || 'N/A'}\n`;
        body += `- Unit Achievement %: ${unitAchievementValue?.textContent || 'N/A'}\n`; body += `- Total Visits: ${visitCountValue?.textContent || 'N/A'}\n`; body += `- Avg. Connectivity: ${retailModeConnectivityValue?.textContent || 'N/A'}\n\n`;
        body += "Mysteryshop & Training (Avg*):\n"; body += `- Rep Skill Ach: ${repSkillAchValue?.textContent || 'N/A'}\n`; body += `- (V)PMR Ach: ${vPmrAchValue?.textContent || 'N/A'}\n`;
        body += `- Post Training Score: ${postTrainingScoreValue?.textContent || 'N/A'} (Excludes 0s)\n`; 
        body += `- Elite Score %: ${eliteValue?.textContent || 'N/A'}\n\n`;
        body += "*Averages calculated only using stores with valid data for each metric.\n\n";
        const territoriesInData = new Set(filteredData.map(row => safeGet(row, 'Q2 Territory', null)).filter(Boolean));
        if (territoriesInData.size === 1) {
            const territoryName = territoriesInData.values().next().value; body += `--- Key Performers for Territory: ${territoryName} ---\n`;
            const top5ForEmail = [...filteredData].sort((a, b) => parseNumber(safeGet(b, 'Revenue w/DF', -Infinity)) - parseNumber(safeGet(a, 'Revenue w/DF', -Infinity))).slice(0, TOP_N_TABLES);
            if (top5ForEmail.length > 0) { body += `Top ${top5ForEmail.length} (Revenue):\n`; top5ForEmail.forEach((store, i) => { body += `${i+1}. ${safeGet(store, 'Store')} - Rev: ${formatCurrency(parseNumber(safeGet(store, 'Revenue w/DF')))}\n`; }); body += "\n"; }
            const bottom5ForEmail = [...filteredData].sort((a, b) => calculateQtdGap(a) - calculateQtdGap(b)).slice(0, TOP_N_TABLES);
            if (bottom5ForEmail.length > 0) { body += `Bottom ${bottom5ForEmail.length} (Opportunities by QTD Gap):\n`; bottom5ForEmail.forEach((store, i) => { body += `${i+1}. ${safeGet(store, 'Store')} - Gap: ${formatCurrency(calculateQtdGap(store) === Infinity ? NaN : calculateQtdGap(store))}\n`; }); body += "\n"; }
        }
        body += "---------------------------------\nGenerated by FSM Dashboard\n"; return body;
    };
    const getFilterSummary = () => {
        let summary = [];
        if (regionFilter?.value !== 'ALL') summary.push(`Region: ${regionFilter.value}`); if (districtFilter?.value !== 'ALL') summary.push(`District: ${districtFilter.value}`);
        const territories = territoryFilter ? Array.from(territoryFilter.selectedOptions).map(o => o.value) : []; if (territories.length > 0) summary.push(`Territories: ${territories.length === 1 ? territories[0] : `${territories.length} selected`}`);
        if (fsmFilter?.value !== 'ALL') summary.push(`FSM: ${fsmFilter.value}`); if (channelFilter?.value !== 'ALL') summary.push(`Channel: ${channelFilter.value}`);
        if (subchannelFilter?.value !== 'ALL') summary.push(`Subchannel: ${subchannelFilter.value}`); if (dealerFilter?.value !== 'ALL') summary.push(`Dealer: ${dealerFilter.value}`);
        const stores = storeFilter ? Array.from(storeFilter.selectedOptions).map(o => o.value) : []; if (stores.length > 0) summary.push(`Stores: ${stores.length === 1 ? stores[0] : `${stores.length} selected`}`);
        const flags = Object.entries(flagFiltersCheckboxes).filter(([, input]) => input?.checked).map(([key])=> key.replace(/_/g, ' ')); if (flags.length > 0) summary.push(`Attributes: ${flags.join(', ')}`);
        const additionalToolsSummary = [];
        if(showMapViewFilter?.checked) additionalToolsSummary.push("Map View");
        if(focusEliteFilter?.checked) additionalToolsSummary.push("Elite Opps");
        if(focusConnectivityFilter?.checked) additionalToolsSummary.push("Connectivity Opps");
        if(focusRepSkillFilter?.checked) additionalToolsSummary.push("Rep Skill Opps");
        if(focusVpmrFilter?.checked) additionalToolsSummary.push("VPMR Opps");
        if(additionalToolsSummary.length > 0) summary.push(`Tools: ${additionalToolsSummary.join(', ')}`);
        return summary.length > 0 ? summary.join('; ') : 'None';
    };
    const handleShareEmail = () => {
        if (!emailRecipientInput || !shareStatus) return; const recipient = emailRecipientInput.value;
        if (!recipient || !/\S+@\S+\.\S+/.test(recipient)) { if(shareStatus) shareStatus.textContent = "Please enter a valid recipient email address."; return; }
        try {
            const subject = `FSM Dashboard Summary - ${new Date().toLocaleDateString()}`; const bodyContent = generateEmailBody();
            const mailtoLink = `mailto:${recipient}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(bodyContent)}`;
            if (mailtoLink.length > 2000) { if(shareStatus) shareStatus.textContent = "Generated email body is too long. Please use Print/Save as PDF."; console.warn("Mailto link > 2000 chars"); return; }
            window.location.href = mailtoLink; if(shareStatus) shareStatus.textContent = "Your email client should open.";
        } catch (error) { console.error("Error generating mailto link:", error); if(shareStatus) shareStatus.textContent = "Error generating email. Check console."; }
    };
     const selectAllOptions = (selectElement) => {
         if (!selectElement) return; Array.from(selectElement.options).forEach(option => option.selected = true);
         if (selectElement === territoryFilter) updateStoreFilterOptionsBasedOnHierarchy();
    };
     const deselectAllOptions = (selectElement) => {
         if (!selectElement) return; selectElement.selectedIndex = -1;
         if (selectElement === territoryFilter) updateStoreFilterOptionsBasedOnHierarchy();
    };

    // --- Event Listeners ---
    if (themeToggleBtn) themeToggleBtn.addEventListener('click', toggleTheme);
    excelFileInput?.addEventListener('change', handleFile);
    storeSearch?.addEventListener('input', filterStoreOptions);
    exportCsvButton?.addEventListener('click', exportData);
    if (printReportButton) printReportButton.addEventListener('click', handlePrintReport);
    if (shareEmailButton) shareEmailButton.addEventListener('click', handleShareEmail);
    closeStoreDetailsButton?.addEventListener('click', hideStoreDetails);
    territorySelectAll?.addEventListener('click', () => selectAllOptions(territoryFilter));
    territoryDeselectAll?.addEventListener('click', () => deselectAllOptions(territoryFilter));
    storeSelectAll?.addEventListener('click', () => selectAllOptions(storeFilter));
    storeDeselectAll?.addEventListener('click', () => deselectAllOptions(storeFilter));
    attachRateTable?.querySelector('thead')?.addEventListener('click', handleSort);

    // --- Initial Setup ---
    const savedTheme = localStorage.getItem(THEME_STORAGE_KEY); 
    applyTheme(savedTheme || 'dark'); 
    initMapView(); 
    resetUI(); 
    if (!mainChartCanvas) console.warn("Main chart canvas context not found on load. Chart will not render.");
    updateShareOptions(); 
    checkAndShowWhatsNew(); 
    checkAndShowDisclaimer(); // Call to show disclaimer on page load

});
