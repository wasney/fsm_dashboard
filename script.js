/*
    Timestamp: 2025-05-08T21:15:26EDT
    Summary: Proofread for consistency and best practices. No major functional changes.
*/
/* --- Global Resets & Body (Dark Mode) --- */
*, *::before, *::after { box-sizing: border-box; }
body {
    font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif, "Apple Color Emoji", "Segoe UI Emoji", "Segoe UI Symbol";
    margin: 0;
    padding: 0;
    background-color: #1e1e1e; /* Dark background */
    color: #e0e0e0; /* Light text */
    line-height: 1.5;
    -webkit-font-smoothing: antialiased;
    -moz-osx-font-smoothing: grayscale;
}
.container { max-width: 1400px; margin: 0 auto; padding: 10px; } /* Wider container */

/* --- Headings (Dark Mode) --- */
h1 { color: #58a6ff; text-align: center; margin-top: 1rem; margin-bottom: 1.5rem; font-weight: 600; font-size: 1.8rem; }
h2 { color: #77b6ff; text-align: center; margin-top: 0; margin-bottom: 1rem; font-weight: 500; font-size: 1.4rem; border-bottom: 1px solid #444; padding-bottom: 0.5rem; }
h3 { margin-top: 0; margin-bottom: 1rem; color: #58a6ff; font-weight: 500; font-size: 1.15rem; text-align: center; }

/* --- Upload Instructions Style --- */
.upload-instructions {
    text-align: center;
    color: #a0a0a0;
    margin-top: -0.5rem;
    margin-bottom: 1.5rem;
    padding: 0 10px;
}

/* --- Card Styling (Dark Mode) --- */
.card {
    background-color: #2c2c2c;
    border: 1px solid #444;
    border-radius: 0.25rem;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.4);
    padding: 1rem;
    margin: 1rem auto;
    width: auto;
    margin-left: 10px;
    margin-right: 10px;
}

/* --- Input Area & Loading Indicator --- */
.input-area { text-align: center; max-width: 600px; }
.input-area label { display: block; margin-bottom: 0.5rem; margin-right: 0; font-weight: 600; color: #c0c0c0; }
#excelFile { border: 1px solid #555; padding: 0.5rem 0.75rem; border-radius: 0.25rem; cursor: pointer; width: 100%; max-width: 350px; background-color: #333; color: #e0e0e0; }
#excelFile::file-selector-button { padding: 0.5rem 0.75rem; border: 1px solid #555; border-radius: 0.25rem; background-color: #444; color: #e0e0e0; cursor: pointer; margin-right: 0.5rem; }
#status.status-message { /* More specific selector for the status paragraph */
    margin-top: 1rem; /* Increased top margin for better separation */
    margin-bottom: 1rem; /* Added bottom margin */
    font-style: italic;
    color: #a0a0a0;
    text-align: center;
    min-height: 1.5em;
    padding: 0 10px;
}
#loadingIndicator, #filterLoadingIndicator { color: #99ccff; margin-top: 0.75rem; font-size: 0.9em; display: flex; align-items: center; justify-content: center; gap: 8px; }
.spinner { border: 4px solid rgba(153, 204, 255, 0.3); border-radius: 50%; border-top: 4px solid #99ccff; width: 20px; height: 20px; animation: spin 1s linear infinite; }
.spinner-small { border: 3px solid rgba(153, 204, 255, 0.3); border-radius: 50%; border-top: 3px solid #99ccff; width: 16px; height: 16px; animation: spin 1s linear infinite; }
@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }

/* --- Filter Controls (Dark Mode) --- */
#filterArea { max-width: 1200px; }
.filter-controls { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 1rem 1.5rem; margin-bottom: 1.5rem; }
.filter-group { display: flex; flex-direction: column; gap: 0.3rem; }
.filter-group label, .flags-label { font-weight: 600; color: #c0c0c0; margin-bottom: 0; width: auto; text-align: left; font-size: 0.95em; }
.filter-group select, .filter-group input[type="text"] { padding: 0.6rem 0.75rem; border: 1px solid #555; border-radius: 0.25rem; min-width: 0; width: 100%; background-color: #333; color: #e0e0e0; font-size: 0.9rem; }
.filter-group select[multiple] { padding: 0.5rem; height: auto; min-height: 100px; font-size: 0.85rem; }
.filter-group select:disabled, .filter-group input:disabled, .filter-group button:disabled, .filter-group-flags input:disabled { background-color: #444 !important; cursor: not-allowed !important; color: #888 !important; border-color: #555 !important; opacity: 0.6; }
.filter-group select option { background-color: #333; color: #e0e0e0; }
.multi-select-info { font-size: 0.75em; color: #999; margin-top: -2px; margin-bottom: 2px; display: block; text-align: left; }
.multi-select-controls { display: flex; gap: 0.5rem; margin-bottom: 0.3rem; }
.select-button { padding: 0.2rem 0.5rem; font-size: 0.75em; background-color: #4a4a4a; color: #ccc; border: 1px solid #666; border-radius: 3px; cursor: pointer; }
.select-button:hover:not(:disabled) { background-color: #5a5a5a; color: #eee; }
#storeSearch { margin-bottom: 0.3rem; }
.filter-group-store { grid-column: span 2; /* Make store filter wider if needed */ }

/* --- Flag Filters --- */
.filter-group-flags { grid-column: 1 / -1; /* Span full width */ margin-top: 0.5rem; }
.flags-label { margin-bottom: 0.5rem; display: block; }
.flag-toggles { display: flex; flex-wrap: wrap; gap: 0.5rem 1.5rem; background-color: #333; padding: 0.75rem; border-radius: 4px; border: 1px solid #444; }
.flag-toggles label { display: flex; align-items: center; gap: 0.4rem; font-size: 0.9em; color: #ccc; cursor: pointer; }
.flag-toggles input[type="checkbox"] { width: 16px; height: 16px; accent-color: #58a6ff; cursor: pointer; margin: 0; }
.flag-toggles label:has(input:disabled) { cursor: not-allowed; color: #888; }


.apply-filters-button {
    display: block;
    width: 100%;
    max-width: 250px;
    margin: 1rem auto 0 auto;
    padding: 0.7rem 1rem;
    background-color: #3081d2;
    color: #ffffff;
    border: none;
    border-radius: 0.25rem;
    font-weight: 600;
    font-size: 1rem;
    cursor: pointer;
    transition: background-color 0.2s ease;
}
.apply-filters-button:hover:not(:disabled) { background-color: #58a6ff; }

/* --- Results Area --- */
.results-container { margin-top: 1.5rem; }

/* --- Store Details Container --- */
.store-details-container { max-width: 95%; position: relative; } /* Allow more width */
#storeDetailsContent { font-size: 0.9rem; text-align: left; color: #d0d0d0; line-height: 1.6; max-height: 400px; overflow-y: auto; padding-right: 10px; /* For scrollbar */}
#storeDetailsContent p { margin-bottom: 0.5rem; }
#storeDetailsContent strong { color: #e5e5e5; margin-right: 5px; }
#storeDetailsContent hr { border: none; border-top: 1px solid #444; margin: 0.75rem 0; }
#storeDetailsContent span[data-flag="true"] { color: #86dc86; font-weight: bold; }
#storeDetailsContent span[data-flag="false"] { color: #aaa; }
.close-button {
    position: absolute;
    top: 10px;
    right: 10px;
    background: #555;
    color: #eee;
    border: none;
    border-radius: 50%;
    width: 24px;
    height: 24px;
    font-size: 16px;
    line-height: 22px;
    text-align: center;
    cursor: pointer;
    font-weight: bold;
}
.close-button:hover { background: #777; }


/* --- Summary & Stat Containers (Dark Mode) --- */
.summary-container, .rep-pmr-container, .training-stats-container { max-width: 100%; }
.summary-container { margin-top: 1rem; } /* Adjust spacing */
.summary-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); /* Responsive grid */
    gap: 0.6rem 1.5rem; /* Row and column gap */
    text-align: left;
    padding-left: 0;
}
.summary-grid p { margin: 0; font-size: 0.95rem; color: #d0d0d0; }
.summary-grid p strong { font-weight: 600; color: #f0f0f0; margin-left: 5px; }
.summary-grid p[style*="display: none"] { display: none !important; } /* Ensure hidden items stay hidden */
.summary-container span[style*="background-color"] { color: #111; padding: 2px 5px; border-radius: 3px; font-weight: 500; } /* For potential future use */

.stats-row {
    display: flex;
    flex-wrap: wrap;
    gap: 1rem;
    margin-top: 1rem;
}
.rep-pmr-container, .training-stats-container {
    flex: 1 1 300px; /* Allow flex grow/shrink, base width 300px */
    max-width: none; /* Remove max-width from individual cards */
}
.rep-pmr-container p, .training-stats-container p { margin-bottom: 0.5rem; font-size: 0.95rem; text-align: left; padding-left: 0; color: #d0d0d0; }
.rep-pmr-container p:last-child, .training-stats-container p:last-child { margin-bottom: 0; }
.rep-pmr-container strong, .training-stats-container strong { font-weight: 600; color: #f0f0f0; }


/* --- Chart Container (Dark Mode) --- */
.chart-container { height: 350px; max-width: 95%; /* Wider */ background-color: #2c2c2c; }
#mainChartCanvas, #secondaryChartCanvas { width: 100% !important; height: 100% !important; }

/* --- Table Container & Table (Dark Mode) --- */
.table-container { max-width: 95%; /* Wider */ }
.table-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.75rem; flex-wrap: wrap; gap: 10px; }
.table-header h2 { margin: 0; font-size: 1.2rem; text-align: left; }
.export-button { padding: 0.4rem 0.8rem; background-color: #4a4a4a; color: #ccc; border: 1px solid #666; border-radius: 4px; font-size: 0.85em; cursor: pointer; }
.export-button:hover:not(:disabled) { background-color: #5a5a5a; color: #eee; }
.table-wrapper { overflow-x: auto; margin-top: 0.75rem; }
#attachRateTable { width: 100%; border-collapse: collapse; margin-top: 0; font-size: 0.85rem; }
#attachRateTable th, #attachRateTable td { border: 1px solid #555; padding: 0.5rem 0.75rem; text-align: left; white-space: nowrap; color: #d0d0d0; }
#attachRateTable th { background-color: #3a3a3a; font-weight: 600; color: #f0f0f0; position: sticky; top: 0; z-index: 1; text-transform: uppercase; letter-spacing: 0.5px; font-size: 0.8rem; cursor: pointer; user-select: none; }
#attachRateTable th:hover { background-color: #4a4a4a; }
#attachRateTable tbody tr { background-color: #2c2c2c; cursor: pointer; } /* Add cursor pointer */
#attachRateTable tbody tr:hover { background-color: #383838; }
#attachRateTable tbody tr.selected-row { background-color: #405d7a !important; /* Highlight selected row */ color: #fff; }
#attachRateTable td:not(:first-child):not(:nth-child(2)) { text-align: right; } /* Right-align numerical columns, except % target */
#attachRateTable tfoot tr { background-color: #333; font-weight: bold; color: #eee; }
#attachRateTable tfoot td { border-top: 2px solid #666; }

.highlight-green { background-color: rgba(42, 74, 42, 0.7) !important; color: #b0ffb0; }
.highlight-red { background-color: rgba(90, 42, 42, 0.7) !important; color: #ffb0b0; }
.sort-arrow { display: inline-block; width: 12px; height: 12px; margin-left: 5px; opacity: 0.5; vertical-align: middle; }
.sort-arrow.asc::after { content: '▲'; font-size: 0.8em; }
.sort-arrow.desc::after { content: '▼'; font-size: 0.8em; }
#attachTableStatus { margin-top: 0.75rem; color: #a0a0a0; font-size: 0.85rem; text-align: center; }

/* --- Share Section (Dark Mode) --- */
.share-container { max-width: 600px; }
.share-controls { display: flex; flex-direction: column; gap: 0.75rem; margin-bottom: 1rem; }
.share-controls label { font-weight: 600; color: #c0c0c0; margin-bottom: 0.25rem; }
.share-controls input[type="email"] { padding: 0.6rem 0.75rem; border: 1px solid #555; border-radius: 0.25rem; flex-grow: 1; background-color: #333; color: #e0e0e0; }
.share-controls button { padding: 0.6rem 1rem; background-color: #3081d2; color: #ffffff; border: none; border-radius: 0.25rem; font-weight: 600; cursor: pointer; transition: background-color 0.2s ease; }
.share-controls button:hover { background-color: #58a6ff; }
.share-status { margin-top: 0.5rem; font-style: italic; color: #77b6ff; min-height: 1.2em; }
.share-note { font-size: 0.85rem; color: #a0a0a0; margin-top: 1rem; text-align: center; }


/* --- MEDIA QUERIES --- */
@media (min-width: 768px) { /* Adjusted breakpoint */
    .container { padding: 15px; }
    .input-area label { display: inline-block; margin-bottom: 0; }
    #excelFile { width: auto; max-width: 350px; }
    .summary-grid p { font-size: 1rem; }
    .rep-pmr-container p, .training-stats-container p { font-size: 1rem; }
    .chart-container { height: 400px; }
    #attachRateTable { font-size: 0.9rem; }
    #attachRateTable th, #attachRateTable td { padding: 0.6rem 0.8rem; } /* Slightly adjust padding */
    #attachRateTable th { font-size: 0.85rem; }
    .share-controls { flex-direction: row; align-items: center; }
    .share-controls label { margin-bottom: 0; }
    .filter-group-store { grid-column: span 1; /* Reset span */ }
}

@media (min-width: 1200px) { /* Adjusted breakpoint */
    h1 { font-size: 2rem; margin-bottom: 2rem; }
    h2 { font-size: 1.5rem; margin-bottom: 1.5rem; }
    h3 { font-size: 1.25rem; }
    .card { padding: 1.5rem; margin: 1.5rem auto; }
    .filter-controls { gap: 1.2rem 2rem; }
    .chart-container { height: 450px; }
    #attachRateTable { font-size: 0.9rem; }
    #attachRateTable th { font-size: 0.9rem; }
    #attachRateTable th, #attachRateTable td { padding: 0.75rem 1rem; }
    .filter-group-store { grid-column: span 2; /* Span 2 columns on larger screens */ }
}
