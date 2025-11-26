
// Check if jsPDF loaded successfully, if not try alternative CDN
if (typeof window.jspdf === 'undefined') {
    const script = document.createElement('script');
    script.src = 'https://unpkg.com/jspdf@2.5.1/dist/jspdf.umd.min.js';
    document.head.appendChild(script);
}

// Check if html2canvas loaded successfully, if not try alternative CDN
if (typeof window.html2canvas === 'undefined') {
    const script = document.createElement('script');
    script.src = 'https://unpkg.com/html2canvas@1.4.1/dist/html2canvas.min.js';
    document.head.appendChild(script);
}

const SHEET_ID = '1_-ZPknnR0N_-iIekrmQE9nKU7Lu2FAA2SQqf2ZXCpYY';
const RANGE = 'Form Responses 1'; // or 'Sheet1!A1:H'
const API_KEY = 'AIzaSyBkZ9IVxc5e0s_PJb0ViznAVWUAP-byDWg';

// Configuration for columns: which to show and their display names
// The order in this array determines the display order
const COLUMN_CONFIG = [
    { key: 'Serial', label: 'Serial Number', show: true },
    { key: '👤 Choose Your Name ', label: 'Name', show: true },
    { key: '📆 Choose Date', label: 'Date', show: true },
    { key: '📁 Choose Project Code', label: 'Project Code', show: true },
    { key: '⏰ Choose Task Start Time', label: 'Start time', show: true },
    { key: '⏰ Choose Task End Time', label: 'End Time', show: true },
    { key: 'Hours worked', label: 'Hours worked', show: true },
    { key: '🏢 Work Mode', label: 'Mode', show: true },
    // Add more columns here if needed
];

// Remove the fetch for spreadsheet metadata and setting of sheetTitle

const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent(RANGE)}?key=${API_KEY}`;

let allRows = [];
let filters = {};
let globalFilters = {
    name: '',
    fromDate: '',
    toDate: ''
};
let currentView = 'table'; // Default view

// Helper to normalize date to YYYY-MM-DD (for filtering)
function normalizeDate(str) {
    if (!str) return '';
    const parts = str.split('/');
    if (parts.length === 3) {
        let [m, d, y] = parts; // Excel format is MM/DD/YYYY
        if (y.length === 4) {
            return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
        }
    }
    return str;
}

// Helper to format date for display as DD-MM-YYYY
function formatDateForDisplay(str) {
    if (!str) return '';
    const parts = str.split('/');
    if (parts.length === 3) {
        let [m, d, y] = parts; // Excel format is MM/DD/YYYY
        if (y.length === 4) {
            // Convert MM/DD/YYYY to DD-MM-YYYY for display
            return `${d.padStart(2, '0')}-${m.padStart(2, '0')}-${y}`;
        }
    }
    return str;
}

// Helper to get filtered rows
function getFilteredRows() {
    if (!allRows || allRows.length < 2) return [];

    const headerRow = allRows[0];
    const nameColIdx = headerRow.indexOf('👤 Choose Your Name ');
    const dateColIdx = headerRow.indexOf('📆 Choose Date');

    return allRows.slice(1).filter(row => {
        // Apply name filter
        if (globalFilters.name && row[nameColIdx] !== globalFilters.name) {
            return false;
        }

        // Apply from date filter
        if (globalFilters.fromDate) {
            const dateStr = row[dateColIdx];
            if (dateStr) {
                const normalizedDate = normalizeDate(dateStr);
                if (normalizedDate < globalFilters.fromDate) {
                    return false;
                }
            }
        }

        // Apply to date filter
        if (globalFilters.toDate) {
            const dateStr = row[dateColIdx];
            if (dateStr) {
                const normalizedDate = normalizeDate(dateStr);
                if (normalizedDate > globalFilters.toDate) {
                    return false;
                }
            }
        }

        // Apply month filter
        const monthValEl = document.getElementById('monthFilter');
        if (monthValEl && monthValEl.value) {
            const parts = monthValEl.value.split('-'); // "YYYY-MM"
            if (parts.length === 2) {
                const yy = parseInt(parts[0], 10);
                const mm = parseInt(parts[1], 10);
                const dStr = normalizeDate(row[dateColIdx] || '');
                const d = new Date(dStr);
                if (isNaN(d) || d.getFullYear() !== yy || (d.getMonth() + 1) !== mm) {
                    return false;
                }
            }
        }

        return true;
    });
}

// Filter Functions
function applyNameFilter() {
    try {
        // console.log('Name filter changed'); 
        const select = document.getElementById('nameFilter');
        globalFilters.name = select.value;
        updateCurrentView();
    } catch (e) {
        alert('Error in applyNameFilter: ' + e.message);
    }
}

function applyFromDateFilter() {
    try {
        const input = document.getElementById('fromDateFilter');
        globalFilters.fromDate = input.value;
        updateCurrentView();
    } catch (e) {
        alert('Error in applyFromDateFilter: ' + e.message);
    }
}

function applyToDateFilter() {
    try {
        const input = document.getElementById('toDateFilter');
        globalFilters.toDate = input.value;
        updateCurrentView();
    } catch (e) {
        alert('Error in applyToDateFilter: ' + e.message);
    }
}

function updateCurrentView() {
    try {
        console.log('updateCurrentView called. currentView:', currentView);

        const dashboardContainer = document.getElementById('dashboardContainer');
        // offsetParent is null if display is none
        const isDashboardVisible = dashboardContainer && dashboardContainer.offsetParent !== null;

        if (isDashboardVisible) {
            console.log('Dashboard is visible. Forcing update.');
            // Sync state just in case
            currentView = 'dashboard';
            updateDashboard();
            return;
        }

        if (currentView === 'projectSummary') {
            console.log('Updating Project Summary');
            showProjectSummary();
        } else if (currentView === 'simpleView') {
            console.log('Updating Simple View');
            showSimpleView();
        } else {
            console.log('Rendering Table');
            renderTable(allRows);
        }
    } catch (e) {
        alert('Error in updateCurrentView: ' + e.message);
    }
}

// Expose functions to window for inline onchange handlers
window.applyNameFilter = applyNameFilter;
window.applyFromDateFilter = applyFromDateFilter;
window.applyToDateFilter = applyToDateFilter;


function renderTable(rows, filters = {}) {
    if (!rows || rows.length === 0) {
        document.getElementById('tableContainer').innerHTML = 'No data found or data format incorrect.';
        return;
    }

    // Find indices of columns to show (excluding 'Timestamp')
    const headerRow = rows[0];
    const colIndices = COLUMN_CONFIG.filter(col => col.key !== 'Serial' && col.show)
        .map(col => headerRow.indexOf(col.key));

    // Apply global filters first
    let filteredRows = rows.slice(1).filter((row, rowIdx) => {
        // Apply name filter
        if (globalFilters.name && row[headerRow.indexOf('👤 Choose Your Name ')] !== globalFilters.name) {
            return false;
        }

        // Apply from date filter
        if (globalFilters.fromDate) {
            const dateStr = row[headerRow.indexOf('📆 Choose Date')];
            if (dateStr) {
                const normalizedDate = normalizeDate(dateStr);
                if (normalizedDate < globalFilters.fromDate) {
                    return false;
                }
            }
        }

        // Apply to date filter
        if (globalFilters.toDate) {
            const dateStr = row[headerRow.indexOf('📆 Choose Date')];
            if (dateStr) {
                const normalizedDate = normalizeDate(dateStr);
                if (normalizedDate > globalFilters.toDate) {
                    return false;
                }
            }
        }
        return true;
    });

    // Apply column-specific filters
    filteredRows = filteredRows.filter((row, rowIdx) =>
        colIndices.every((colIdx, i) => {
            const filterVal = filters[colIdx];
            const startDateFilter = filters[colIdx + '_start'];
            const endDateFilter = filters[colIdx + '_end'];

            if (headerRow[colIdx] === '📆 Choose Date') {
                // Handle date range filtering
                if (startDateFilter || endDateFilter) {
                    const rowDate = normalizeDate(row[colIdx]);
                    if (!rowDate) return false;

                    const rowDateObj = new Date(rowDate);
                    if (isNaN(rowDateObj.getTime())) return false;

                    if (startDateFilter) {
                        const startDate = new Date(startDateFilter);
                        if (rowDateObj < startDate) return false;
                    }

                    if (endDateFilter) {
                        const endDate = new Date(endDateFilter);
                        if (rowDateObj > endDate) return false;
                    }

                    return true;
                } else if (filterVal) {
                    // Handle single date filter (backward compatibility)
                    return normalizeDate(row[colIdx]) === filterVal;
                }
                return true;
            }

            // Handle other column filters
            if (!filterVal) return true;
            return row[colIdx] === filterVal;
        })
    );

    // Sort filtered rows by date in ascending order
    const dateColIdx = headerRow.indexOf('📆 Choose Date');
    filteredRows.sort((a, b) => {
        const dateA = normalizeDate(a[dateColIdx]);
        const dateB = normalizeDate(b[dateColIdx]);

        if (!dateA && !dateB) return 0;
        if (!dateA) return 1;
        if (!dateB) return -1;

        return dateA.localeCompare(dateB);
    });

    let html = '<table><thead><tr>';
    // Render dropdown filters and headers
    COLUMN_CONFIG.forEach((col, i) => {
        if (!col.show) return;
        if (col.key === 'Serial') {
            html += `<th><div style="display: flex; justify-content: space-between; align-items: center;">
                        <span>${col.label}</span>
                        <span></span>
                     </div></th>`;
        } else if (col.key === '📆 Choose Date') {
            html += `<th><div class="filter-header"><span>${col.label}</span></div></th>`;
        } else if (col.key === '⏰ Choose Task Start Time' || col.key === '⏰ Choose Task End Time' || col.key === 'Hours worked') {
            html += `<th><div class="filter-header"><span>${col.label}</span></div></th>`;
        } else if (col.key === '👤 Choose Your Name ') {
            html += `<th><div class="filter-header"><span>${col.label}</span></div></th>`;
        } else {
            const colIdx = headerRow.indexOf(col.key);
            const uniqueVals = [...new Set(rows.slice(1).map(row => row[colIdx] || ''))];
            html += `<th>
                <div class="filter-header">
                    <span>${col.label}</span>
                    <select class="filter-dropdown" onchange="applyFilter(${colIdx}, this.value)">
                        <option value=""${!filters[colIdx] ? ' selected' : ''}>All</option>
                        ${uniqueVals.map(val =>
                `<option value="${val}"${filters[colIdx] === val ? ' selected' : ''}>${val}</option>`
            ).join('')}
                    </select>
                </div>
            </th>`;
        }
    });
    html += '</tr></thead><tbody>';

    // Calculate total hours worked
    let totalHours = 0;
    let totalMinutes = 0;

    // Render filtered rows
    filteredRows.forEach((row, i) => {
        html += '<tr class="table-row-anim" style="animation-delay: ' + (i * 0.05) + 's">';
        COLUMN_CONFIG.forEach(col => {
            if (!col.show) return;
            if (col.key === 'Serial') {
                html += `<td>${i + 1}</td>`;
            } else if (col.key === 'Hours worked') {
                const startIdx = headerRow.indexOf('⏰ Choose Task Start Time');
                const endIdx = headerRow.indexOf('⏰ Choose Task End Time');
                const start = row[startIdx] || '';
                const end = row[endIdx] || '';

                function parseTime12h(t) {
                    // Parse time in 12-hour format (e.g., 9:00:00 AM)
                    const [time, modifier] = t.split(' ');
                    let [hours, minutes, seconds] = time.split(':').map(Number);
                    if (modifier === 'PM' && hours !== 12) hours += 12;
                    if (modifier === 'AM' && hours === 12) hours = 0;
                    return hours * 60 + minutes + (seconds ? seconds / 60 : 0);
                }

                let total = '';
                if (start && end) {
                    const startMins = parseTime12h(start);
                    const endMins = parseTime12h(end);
                    let diffMins = endMins - startMins;
                    if (diffMins < 0) diffMins += 24 * 60; // handle overnight
                    const hours = Math.floor(diffMins / 60);
                    const mins = Math.round(diffMins % 60);
                    total = hours + (mins > 0 ? (':' + mins.toString().padStart(2, '0')) : '');

                    // Add to total calculation
                    totalHours += hours;
                    totalMinutes += mins;
                }
                html += `<td>${total}</td>`;
            } else if (col.key === '📆 Choose Date') {
                const colIdx = headerRow.indexOf(col.key);
                html += `<td>${formatDateForDisplay(row[colIdx])}</td>`;
            } else if (col.key === '⏰ Choose Task Start Time' || col.key === '⏰ Choose Task End Time') {
                const colIdx = headerRow.indexOf(col.key);
                const timeValue = row[colIdx] || '';
                // Add leading zero only to single-digit hours (1-9)
                const formattedTime = timeValue.replace(/^(\d):/g, '0$1:');
                html += `<td>${formattedTime}</td>`;
            } else {
                const colIdx = headerRow.indexOf(col.key);
                html += `<td>${row[colIdx]}</td>`;
            }
        });
        html += '</tr>';
    });

    // Add total row
    if (filteredRows.length > 0) {
        // Convert excess minutes to hours
        totalHours += Math.floor(totalMinutes / 60);
        totalMinutes = totalMinutes % 60;

        html += '<tr style="background-color: #e8f4fd; font-weight: bold;">';
        COLUMN_CONFIG.forEach(col => {
            if (!col.show) return;
            if (col.key === 'Serial') {
                html += `<td>Total</td>`;
            } else if (col.key === 'Hours worked') {
                const totalDisplay = totalHours + (totalMinutes > 0 ? (':' + totalMinutes.toString().padStart(2, '0')) : '');
                html += `<td>${totalDisplay}</td>`;
            } else {
                html += `<td></td>`;
            }
        });
        html += '</tr>';
    }

    html += '</tbody></table>';
    document.getElementById('tableContainer').innerHTML = html;
}

window.applyFilter = function (colIdx, value) {
    filters[colIdx] = value;
    renderTable(allRows, filters);
};

window.applyDateRangeFilter = function (colIdx, value, type) {
    if (type === 'start') {
        filters[colIdx + '_start'] = value;
    } else if (type === 'end') {
        filters[colIdx + '_end'] = value;
    }
    renderTable(allRows, filters);
};

// Global filter functions
window.applyNameFilter = function () {
    globalFilters.name = document.getElementById('nameFilter').value;
    renderTable(allRows, filters);
};

window.applyFromDateFilter = function () {
    globalFilters.fromDate = document.getElementById('fromDateFilter').value;
    renderTable(allRows, filters);
};

window.applyToDateFilter = function () {
    globalFilters.toDate = document.getElementById('toDateFilter').value;
    renderTable(allRows, filters);
};

// Function to show simple view (serial, name, date, hours worked)
function showSimpleView() {
    if (!allRows || allRows.length < 2) return;

    const headerRow = allRows[0];
    const nameColIdx = headerRow.indexOf('👤 Choose Your Name ');
    const dateColIdx = headerRow.indexOf('📆 Choose Date');
    const startTimeColIdx = headerRow.indexOf('⏰ Choose Task Start Time');
    const endTimeColIdx = headerRow.indexOf('⏰ Choose Task End Time');

    // Apply the same filters as the main table
    let filteredRows = getFilteredRows();

    // Group by date and collect all projects with their hours for each date
    const dateSummary = {};
    // Compute Days Worked per employee (unique dates in filtered rows)
    const daysWorkedByEmployee = {};
    filteredRows.forEach(row => {
        const empName = row[nameColIdx] || '';
        const dStr = normalizeDate(row[dateColIdx] || '');
        if (!empName || !dStr) return;
        if (!daysWorkedByEmployee[empName]) {
            daysWorkedByEmployee[empName] = new Set();
        }
        // Consider a row a "worked" record if both start and end times are present/valid
        const st = row[startTimeColIdx] || '';
        const et = row[endTimeColIdx] || '';
        if (st && et && st !== 'Invalid Time' && et !== 'Invalid Time') {
            daysWorkedByEmployee[empName].add(dStr);
        }
    });


    filteredRows.forEach(row => {
        const name = row[nameColIdx] || '';
        const originalDate = row[dateColIdx] || '';
        const normalizedDate = normalizeDate(originalDate); // Use normalized date for grouping
        const displayDate = formatDateForDisplay(originalDate); // Use display format for showing
        const projectCode = row[headerRow.indexOf('📁 Choose Project Code')] || 'Unknown Project';
        const start = row[startTimeColIdx] || '';
        const end = row[endTimeColIdx] || '';

        function parseTime12h(t) {
            const [time, modifier] = t.split(' ');
            let [hours, minutes, seconds] = time.split(':').map(Number);
            if (modifier === 'PM' && hours !== 12) hours += 12;
            if (modifier === 'AM' && hours === 12) hours = 0;
            return hours * 60 + minutes + (seconds ? seconds / 60 : 0);
        }

        if (start && end && normalizedDate) {
            const startMins = parseTime12h(start);
            const endMins = parseTime12h(end);
            let diffMins = endMins - startMins;
            if (diffMins < 0) diffMins += 24 * 60;

            // When "All Names" is selected, group by date + name to show individual employee hours
            const groupKey = !globalFilters.name || globalFilters.name === '' ? `${normalizedDate}_${name}` : normalizedDate;

            if (!dateSummary[groupKey]) {
                dateSummary[groupKey] = {
                    name: name,
                    displayDate: displayDate,
                    projects: {},
                    totalMinutes: 0
                };
            }

            // Add project hours
            if (!dateSummary[groupKey].projects[projectCode]) {
                dateSummary[groupKey].projects[projectCode] = 0;
            }
            dateSummary[groupKey].projects[projectCode] += diffMins;
            dateSummary[groupKey].totalMinutes += diffMins;
        }
    });

    // Create consolidated report table
    let simpleHtml = '<div style="margin: 20px 0; padding: 20px; background: #f8f9fa; border-radius: 8px; border: 1px solid #e9ecef;">';
    simpleHtml += '<h3 style="margin-top: 0';
    // Days Worked summary for the selected month
    (function () {
        const monthValEl = document.getElementById('monthFilter');
        if (!monthValEl || !monthValEl.value) return;
        const monthLabel = new Date(monthValEl.value + '-01').toLocaleString('en-GB', { month: 'long', year: 'numeric' });

        if (globalFilters.name) {
            const cnt = (daysWorkedByEmployee[globalFilters.name]?.size) || 0;
            simpleHtml += `
                <div style="margin: 10px 0 14px 0; padding: 8px 12px; display:inline-block; border-radius: 8px; background:#e8f4fd; border:1px solid #cfe8ff;">
                  <strong>Days Worked in ${monthLabel}:</strong> ${cnt}
                </div>
            `;
        } else {
            const rows = Object.keys(daysWorkedByEmployee).sort().map(emp => {
                const cnt = daysWorkedByEmployee[emp]?.size || 0;
                return `<tr style="background:#fff;">
                          <td style="padding:8px; border:1px solid #ccc;">${emp}</td>
                          <td style="padding:8px; border:1px solid #ccc; text-align:center;">${cnt}</td>
                        </tr>`;
            }).join('');
            simpleHtml += `
                <div style="margin: 10px 0 14px 0;">
                  <h4 style="margin: 0 0 8px 0; color:#4285f4;">Days Worked in ${monthLabel}</h4>
                  <table style="width:100%; border-collapse:collapse; margin-top:4px;">
                    <thead>
                      <tr style="background:#4285f4; color:#fff;">
                        <th style="padding:8px; border:1px solid #ccc; text-align:center;">Name</th>
                        <th style="padding:8px; border:1px solid #ccc; text-align:center;">Days Worked</th>
                      </tr>
                    </thead>
                    <tbody>${rows}</tbody>
                  </table>
                </div>
            `;
        }
    })();
    simpleHtml += '<h3 style="margin-top: 15px; margin-bottom: 10px; color: #333;">Consolidate Report</h3>';

    if (Object.keys(dateSummary).length === 0) {
        simpleHtml += '<p>No data found for the selected filters.</p>';
    } else {
        // Check if "All Names" is selected (empty value means all names)
        const showNameColumn = !globalFilters.name || globalFilters.name === '';

        // Only show employee name above table if a specific name is selected
        if (!showNameColumn) {
            const employeeName = Object.values(dateSummary)[0]?.name || 'Unknown';
            simpleHtml += `<h4 style="margin: 15px 0 10px 0; color: #4285f4; font-size: 18px;">Employee: ${employeeName}</h4>`;
        }

        simpleHtml += '<table style="width: 100%; border-collapse: collapse; margin-top: 10px;">';
        simpleHtml += '<thead><tr style="background: #4285f4; color: white;">';
        simpleHtml += '<th style="padding: 10px; border: 1px solid #ccc; text-align: center;">Serial</th>';
        if (showNameColumn) {
            simpleHtml += '<th style="padding: 10px; border: 1px solid #ccc; text-align: center;">Name</th>';
        }
        simpleHtml += '<th style="padding: 10px; border: 1px solid #ccc; text-align: center;">Date</th>';
        simpleHtml += '<th style="padding: 10px; border: 1px solid #ccc; text-align: center;">Projects & Hours</th>';
        simpleHtml += '</tr></thead><tbody>';

        let grandTotal = 0;
        const sortedDates = Object.keys(dateSummary).sort();

        sortedDates.forEach((date, i) => {
            const summary = dateSummary[date];
            grandTotal += summary.totalMinutes;

            // Create project details HTML with table structure for alignment
            let projectsHtml = '<table style="width: 50%; border-collapse: collapse; margin: 0 auto;">';

            // Sort projects in the specified order: VES, VEB, INDIRECT, NO-WORK
            const sortedProjects = Object.entries(summary.projects).sort(([a], [b]) => {
                const aTrimmed = a.trim().toUpperCase();
                const bTrimmed = b.trim().toUpperCase();

                // Check if project names start with our target strings
                const aStartsWithVES = aTrimmed.startsWith('VES');
                const aStartsWithVEB = aTrimmed.startsWith('VEB');
                const aStartsWithINDIRECT = aTrimmed.startsWith('INDIRECT');
                const aStartsWithNOWORK = aTrimmed.startsWith('NO-WORK');

                const bStartsWithVES = bTrimmed.startsWith('VES');
                const bStartsWithVEB = bTrimmed.startsWith('VEB');
                const bStartsWithINDIRECT = bTrimmed.startsWith('INDIRECT');
                const bStartsWithNOWORK = bTrimmed.startsWith('NO-WORK');

                // Assign priorities
                let aOrder = 5;
                if (aStartsWithVES) aOrder = 1;
                else if (aStartsWithVEB) aOrder = 2;
                else if (aStartsWithINDIRECT) aOrder = 3;
                else if (aStartsWithNOWORK) aOrder = 4;

                let bOrder = 5;
                if (bStartsWithVES) bOrder = 1;
                else if (bStartsWithVEB) bOrder = 2;
                else if (bStartsWithINDIRECT) bOrder = 3;
                else if (bStartsWithNOWORK) bOrder = 4;

                // If both are in the same priority group, sort by number if applicable
                if (aOrder === bOrder && aOrder <= 2) { // VES and VEB groups
                    const aNumber = parseInt(aTrimmed.match(/\d+$/)?.[0] || '0');
                    const bNumber = parseInt(bTrimmed.match(/\d+$/)?.[0] || '0');
                    if (aNumber !== bNumber) {
                        return aNumber - bNumber;
                    }
                }

                return aOrder - bOrder;
            });

            sortedProjects.forEach(([projectCode, minutes], projectIndex) => {
                const hours = Math.floor(minutes / 60);
                const mins = Math.round(minutes % 60);
                const timeDisplay = hours + (mins > 0 ? (':' + minutes.toString().padStart(2, '0')) : '');
                // Alternate background colors for project rows
                const projectRowBackground = projectIndex % 2 === 0 ? '#e9ecef' : '#f8f9fa';
                projectsHtml += `<tr style="background: ${projectRowBackground};">
                    <td style="padding: 2px 4px 2px 0; text-align: left; font-weight: bold; color: #4285f4; border: none;">${projectCode} -</td>
                    <td style="padding: 2px 0; text-align: right; font-size: 16px; font-weight: 500; border: none;">${timeDisplay}</td>
                </tr>`;
            });

            // Add total for this date
            const dateHours = Math.floor(summary.totalMinutes / 60);
            const dateMinutes = Math.round(summary.totalMinutes % 60);
            const dateTotal = dateHours + (dateMinutes > 0 ? (':' + dateMinutes.toString().padStart(2, '0')) : '');

            projectsHtml += `<tr style="border-top: 1px solid #ddd;">
                <td style="padding: 4px 4px 2px 0; text-align: left; font-weight: bold; color: #333; border: none;">Total:</td>
                <td style="padding: 4px 0 2px 0; text-align: right; font-weight: bold; color: #333; border: none;">${dateTotal}</td>
            </tr>`;
            projectsHtml += '</table>';

            // Alternate row colors - even rows are darker
            const rowBackground = i % 2 === 0 ? '#fff' : '#e6e6e6';
            simpleHtml += `<tr class="table-row-anim" style="background: ${rowBackground}; animation-delay: ${i * 0.05}s;">`;
            simpleHtml += `<td style="padding: 10px; border: 1px solid #ccc;">${i + 1}</td>`;
            if (showNameColumn) {
                simpleHtml += `<td style="padding: 10px; border: 1px solid #ccc;">${summary.name}</td>`;
            }
            simpleHtml += `<td style="padding: 10px; border: 1px solid #ccc;">${summary.displayDate}</td>`;
            simpleHtml += `<td style="padding: 10px; border: 1px solid #ccc;">${projectsHtml}</td>`;
            simpleHtml += '</tr>';
        });

        // Add grand total row
        if (sortedDates.length > 0) {
            const grandTotalHours = Math.floor(grandTotal / 60);
            const grandTotalMinutes = Math.round(grandTotal % 60);
            const grandTotalDisplay = grandTotalHours + (grandTotalMinutes > 0 ? (':' + grandTotalMinutes.toString().padStart(2, '0')) : '');

            simpleHtml += '<tr style="background: #e8f4fd; font-weight: bold;">';
            simpleHtml += '<td style="padding: 10px; border: 1px solid #ccc;">Total</td>';
            if (showNameColumn) {
                simpleHtml += '<td style="padding: 10px; border: 1px solid #ccc;"></td>';
            }
            simpleHtml += '<td style="padding: 10px; border: 1px solid #ccc;"></td>';
            simpleHtml += `<td style="padding: 10px; border: 1px solid #ccc;">${grandTotalDisplay}</td>`;
            simpleHtml += '</tr>';
        }

        simpleHtml += '</tbody></table>';
    }

    simpleHtml += '</div>';

    document.getElementById('tableContainer').innerHTML = simpleHtml;
}

// Function to show project summary
function showProjectSummary() {
    if (!allRows || allRows.length < 2) return;

    const headerRow = allRows[0];
    const nameColIdx = headerRow.indexOf('👤 Choose Your Name ');
    const dateColIdx = headerRow.indexOf('📆 Choose Date');
    const projectColIdx = headerRow.indexOf('📁 Choose Project Code');
    const startTimeColIdx = headerRow.indexOf('⏰ Choose Task Start Time');
    const endTimeColIdx = headerRow.indexOf('⏰ Choose Task End Time');

    // Apply the same filters as the main table
    let filteredRows = getFilteredRows();

    // Group by project code and calculate total hours
    const projectSummary = {};

    filteredRows.forEach(row => {
        const projectCode = row[projectColIdx] || 'Unknown Project';
        const start = row[startTimeColIdx] || '';
        const end = row[endTimeColIdx] || '';

        function parseTime12h(t) {
            const [time, modifier] = t.split(' ');
            let [hours, minutes, seconds] = time.split(':').map(Number);
            if (modifier === 'PM' && hours !== 12) hours += 12;
            if (modifier === 'AM' && hours === 12) hours = 0;
            return hours * 60 + minutes + (seconds ? seconds / 60 : 0);
        }

        if (start && end) {
            const startMins = parseTime12h(start);
            const endMins = parseTime12h(end);
            let diffMins = endMins - startMins;
            if (diffMins < 0) diffMins += 24 * 60;

            if (!projectSummary[projectCode]) {
                projectSummary[projectCode] = 0;
            }
            projectSummary[projectCode] += diffMins;
        }
    });

    // Create summary table
    let summaryHtml = '<div style="margin: 20px 0; padding: 20px; background: #f8f9fa; border-radius: 8px; border: 1px solid #e9ecef;">';
    summaryHtml += '<h3 style="margin-top: 0; color: #333;">Project Summary</h3>';

    if (Object.keys(projectSummary).length === 0) {
        summaryHtml += '<p>No data found for the selected filters.</p>';
    } else {
        summaryHtml += '<table style="width: 100%; border-collapse: collapse; margin-top: 10px;">';
        summaryHtml += '<thead><tr style="background: #4285f4; color: white;">';
        summaryHtml += '<th style="padding: 10px; border: 1px solid #ccc; text-align: center;">Project Code</th>';
        summaryHtml += '<th style="padding: 10px; border: 1px solid #ccc; text-align: center;">Total Hours</th>';
        summaryHtml += '</tr></thead><tbody>';

        let grandTotal = 0;

        // Sort projects in the specified order: VES, VEB, INDIRECT, NO-WORK
        const sortedProjects = Object.entries(projectSummary).sort(([a], [b]) => {
            const order = { 'VES': 1, 'VEB': 2, 'INDIRECT': 3, 'NO-WORK': 4 };
            const aTrimmed = a.trim().toUpperCase();
            const bTrimmed = b.trim().toUpperCase();

            // Check if project names start with our target strings
            const aStartsWithVES = aTrimmed.startsWith('VES');
            const aStartsWithVEB = aTrimmed.startsWith('VEB');
            const aStartsWithINDIRECT = aTrimmed.startsWith('INDIRECT');
            const aStartsWithNOWORK = aTrimmed.startsWith('NO-WORK');

            const bStartsWithVES = bTrimmed.startsWith('VES');
            const bStartsWithVEB = bTrimmed.startsWith('VEB');
            const bStartsWithINDIRECT = bTrimmed.startsWith('INDIRECT');
            const bStartsWithNOWORK = bTrimmed.startsWith('NO-WORK');

            // Assign priorities
            let aOrder = 5;
            if (aStartsWithVES) aOrder = 1;
            else if (aStartsWithVEB) aOrder = 2;
            else if (aStartsWithINDIRECT) aOrder = 3;
            else if (aStartsWithNOWORK) aOrder = 4;

            let bOrder = 5;
            if (bStartsWithVES) bOrder = 1;
            else if (bStartsWithVEB) bOrder = 2;
            else if (bStartsWithINDIRECT) bOrder = 3;
            else if (bStartsWithNOWORK) bOrder = 4;

            // If both are in the same priority group, sort by number if applicable
            if (aOrder === bOrder && aOrder <= 2) { // VES and VEB groups
                const aNumber = parseInt(aTrimmed.match(/\d+$/)?.[0] || '0');
                const bNumber = parseInt(bTrimmed.match(/\d+$/)?.[0] || '0');
                if (aNumber !== bNumber) {
                    return aNumber - bNumber;
                }
            }

            return aOrder - bOrder;
        });

        sortedProjects.forEach(([project, totalMinutes], i) => {
            const hours = Math.floor(totalMinutes / 60);
            const minutes = Math.round(totalMinutes % 60);
            const timeDisplay = hours + (minutes > 0 ? (':' + minutes.toString().padStart(2, '0')) : '');
            grandTotal += totalMinutes;

            summaryHtml += `<tr class="table-row-anim" style="background: #fff; animation-delay: ${i * 0.05}s;">`;
            summaryHtml += `<td style="padding: 10px; border: 1px solid #ccc;">${project}</td>`;
            summaryHtml += `<td style="padding: 10px; border: 1px solid #ccc;">${timeDisplay}</td>`;
            summaryHtml += '</tr>';
        });

        // Add grand total row
        const grandTotalHours = Math.floor(grandTotal / 60);
        const grandTotalMinutes = Math.round(grandTotal % 60);
        const grandTotalDisplay = grandTotalHours + (grandTotalMinutes > 0 ? (':' + grandTotalMinutes.toString().padStart(2, '0')) : '');

        summaryHtml += '<tr style="background: #e8f4fd; font-weight: bold;">';
        summaryHtml += '<td style="padding: 10px; border: 1px solid #ccc;">Total</td>';
        summaryHtml += `<td style="padding: 10px; border: 1px solid #ccc;">${grandTotalDisplay}</td>`;
        summaryHtml += '</tr>';

        summaryHtml += '</tbody></table>';
    }

    summaryHtml += '</div>';

    // Show the summary in a modal or replace the table
    document.getElementById('tableContainer').innerHTML = summaryHtml;
}

// Function to populate filter dropdowns
function populateFilterDropdowns() {
    if (!allRows || allRows.length < 2) return;

    const headerRow = allRows[0];
    const nameColIdx = headerRow.indexOf('👤 Choose Your Name ');

    // Populate name filter
    const nameFilter = document.getElementById('nameFilter');
    const uniqueNames = [...new Set(allRows.slice(1).map(row => row[nameColIdx] || '').filter(name => name))];
    uniqueNames.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        nameFilter.appendChild(option);
    });
}

function setDateLimits() {
    const today = new Date().toISOString().split('T')[0]; // Get current date in YYYY-MM-DD format
    document.getElementById('fromDateFilter').max = today;
    document.getElementById('toDateFilter').max = today;
}

// Add timeout to fetch request
const controller = new AbortController();
const timeoutId = setTimeout(() => controller.abort(), 10000); // 10 second timeout

fetch(url, { signal: controller.signal })
    .then(response => {
        clearTimeout(timeoutId);
        if (!response.ok) {
            throw new Error('Network response was not ok');
        }
        return response.json();
    })
    .then(data => {
        if (!data.values || data.values.length === 0) {
            document.getElementById('tableContainer').innerHTML = 'No data found in the sheet.';
            return;
        }

        allRows = data.values;

        // Initial render
        renderTable(allRows);
        populateFilterDropdowns();
        setDateLimits();

        // Initialize Dashboard
        updateDashboard();
    })
    .catch(err => {
        console.error('Error fetching data:', err);
        let errorMsg = 'Failed to load data. ';
        if (err.name === 'AbortError') {
            errorMsg = 'Request timed out. Please check your internet connection.';
        } else {
            errorMsg += err.message;
        }
        document.getElementById('tableContainer').innerHTML = `<div style="color: red; padding: 20px; text-align: center;">${errorMsg}<br><button onclick="location.reload()" style="margin-top: 10px; padding: 8px 16px; cursor: pointer;">Retry</button></div>`;
    });

// Event Listeners for Filters
document.getElementById('nameFilter').addEventListener('change', applyNameFilter);
document.getElementById('fromDateFilter').addEventListener('change', applyFromDateFilter);
document.getElementById('toDateFilter').addEventListener('change', applyToDateFilter);

// Event Listeners for Buttons
document.getElementById('menuToggleBtn').addEventListener('click', function () {
    const dropdown = document.getElementById('menuDropdown');
    dropdown.classList.toggle('show');
    // Change arrow direction
    const btn = this;
    if (dropdown.classList.contains('show')) {
        btn.textContent = btn.textContent.replace('▼', '▲');
    } else {
        btn.textContent = btn.textContent.replace('▲', '▼');
    }
});

// Close dropdown when clicking outside
document.addEventListener('click', function (event) {
    const dropdown = document.getElementById('menuDropdown');
    const toggleBtn = document.getElementById('menuToggleBtn');
    if (!toggleBtn.contains(event.target) && !dropdown.contains(event.target)) {
        dropdown.classList.remove('show');
        toggleBtn.textContent = toggleBtn.textContent.replace('▲', '▼');
    }
});

document.getElementById('projectSummaryBtn').addEventListener('click', function () {
    currentView = 'projectSummary';
    showProjectSummary();
    document.getElementById('dashboardContainer').style.display = 'none';
    document.getElementById('tableContainer').style.display = 'block';
    // Close dropdown after selection
    document.getElementById('menuDropdown').classList.remove('show');
    document.getElementById('menuToggleBtn').textContent = document.getElementById('menuToggleBtn').textContent.replace(/[▼▲]/, '▼');
});

// Simple view button functionality
document.getElementById('simpleViewBtn').addEventListener('click', function () {
    currentView = 'simpleView';
    showSimpleView();
    document.getElementById('dashboardContainer').style.display = 'none';
    document.getElementById('tableContainer').style.display = 'block';
    // Close dropdown after selection
    document.getElementById('menuDropdown').classList.remove('show');
    document.getElementById('menuToggleBtn').textContent = document.getElementById('menuToggleBtn').textContent.replace(/[▼▲]/, '▼');
});

// Dashboard button functionality
document.getElementById('dashboardBtn').addEventListener('click', function () {
    currentView = 'dashboard';
    document.getElementById('tableContainer').style.display = 'none';
    document.getElementById('dashboardContainer').style.display = 'block';
    updateDashboard();
    // Close dropdown after selection
    document.getElementById('menuDropdown').classList.remove('show');
    document.getElementById('menuToggleBtn').textContent = document.getElementById('menuToggleBtn').textContent.replace(/[▼▲]/, '▼');
});

// Home button functionality
document.getElementById('homeBtn').addEventListener('click', function () {
    location.reload();
});

document.getElementById('downloadPdfBtn').addEventListener('click', function () {
    window.print();
});

// Month filter functionality
const monthFilter = document.getElementById('monthFilter');
if (monthFilter) {
    monthFilter.addEventListener('change', function () {
        updateCurrentView();
    });
}

// --- New Features Logic ---

// 1. Dark Mode Toggle
const themeToggleBtn = document.getElementById('themeToggleBtn');
const body = document.body;

// Check local storage for theme preference
const currentTheme = localStorage.getItem('theme');
if (currentTheme === 'dark') {
    body.setAttribute('data-theme', 'dark');
    themeToggleBtn.textContent = '☀️';
}

if (themeToggleBtn) {
    themeToggleBtn.addEventListener('click', () => {
        if (body.getAttribute('data-theme') === 'dark') {
            body.removeAttribute('data-theme');
            themeToggleBtn.textContent = '🌙';
            localStorage.setItem('theme', 'light');
        } else {
            body.setAttribute('data-theme', 'dark');
            themeToggleBtn.textContent = '☀️';
            localStorage.setItem('theme', 'dark');
        }
        // Update charts if dashboard is visible
        if (document.getElementById('dashboardContainer').style.display === 'block') {
            updateDashboard();
        }
    });
}

// 2. Export to CSV
const exportCsvBtn = document.getElementById('exportCsvBtn');
if (exportCsvBtn) {
    exportCsvBtn.addEventListener('click', function () {
        exportToCSV(allRows);
    });
}

function exportToCSV(rows) {
    if (!rows || rows.length === 0) return;

    let csvContent = "data:text/csv;charset=utf-8,";

    // Add header
    csvContent += rows[0].join(",") + "\r\n";

    rows.slice(1).forEach(row => {
        const rowStr = row.map(cell => {
            // Escape quotes and wrap in quotes if contains comma
            let cellStr = String(cell || '');
            if (cellStr.includes(',') || cellStr.includes('"')) {
                return `"${cellStr.replace(/"/g, '""')}"`;
            }
            return cellStr;
        }).join(",");
        csvContent += rowStr + "\r\n";
    });

    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "timesheet_export.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// 3. Analytics Dashboard
let projectChartInstance = null;
let dailyChartInstance = null;

function updateDashboard() {
    if (!allRows || allRows.length < 2) return;

    // Debug: Update title to show it's refreshing
    const title = document.querySelector('#dashboardContainer .chart-title');
    if (title) {
        const rowsToProcess = getFilteredRows();
        title.innerHTML = `Analytics Dashboard <span style="font-size: 12px; font-weight: normal;">(Showing ${rowsToProcess.length} rows)</span>`;
    }

    const headerRow = allRows[0];
    const projectColIdx = headerRow.indexOf('📁 Choose Project Code');
    const dateColIdx = headerRow.indexOf('📆 Choose Date');
    const startTimeColIdx = headerRow.indexOf('⏰ Choose Task Start Time');
    const endTimeColIdx = headerRow.indexOf('⏰ Choose Task End Time');

    // Aggregate Data
    const projectHours = {};
    const dailyHours = {};

    const rowsToProcess = getFilteredRows();

    rowsToProcess.forEach(row => {
        const project = row[projectColIdx] || 'Unknown';
        const date = normalizeDate(row[dateColIdx]);
        const start = row[startTimeColIdx];
        const end = row[endTimeColIdx];

        if (start && end) {
            const startMins = parseTime12h(start);
            const endMins = parseTime12h(end);
            let diffMins = endMins - startMins;
            if (diffMins < 0) diffMins += 24 * 60;

            const hours = diffMins / 60;

            // Project Aggregation
            projectHours[project] = (projectHours[project] || 0) + hours;

            // Daily Aggregation
            if (date) {
                dailyHours[date] = (dailyHours[date] || 0) + hours;
            }
        }
    });

    // Helper for time parsing (duplicated from renderTable, should be utility)
    function parseTime12h(t) {
        if (!t) return 0;
        const [time, modifier] = t.split(' ');
        let [hours, minutes, seconds] = time.split(':').map(Number);
        if (modifier === 'PM' && hours !== 12) hours += 12;
        if (modifier === 'AM' && hours === 12) hours = 0;
        return hours * 60 + minutes + (seconds ? seconds / 60 : 0);
    }

    // Prepare Chart Data
    const projectLabels = Object.keys(projectHours);
    const projectData = Object.values(projectHours);

    const sortedDates = Object.keys(dailyHours).sort();
    const dailyLabels = sortedDates.map(d => formatDateForDisplay(d)); // Use display format
    const dailyData = sortedDates.map(d => dailyHours[d]);

    // Chart Colors (Dynamic based on theme)
    const isDark = document.body.getAttribute('data-theme') === 'dark';
    const textColor = isDark ? '#e0e0e0' : '#333';
    const gridColor = isDark ? '#444' : '#ddd';

    // Render Project Chart
    const ctxProject = document.getElementById('projectChart').getContext('2d');
    if (projectChartInstance) projectChartInstance.destroy();

    projectChartInstance = new Chart(ctxProject, {
        type: 'pie',
        data: {
            labels: projectLabels,
            datasets: [{
                data: projectData,
                backgroundColor: [
                    '#4285f4', '#34a853', '#fbbc05', '#ea4335', '#ab47bc', '#00acc1'
                ]
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { labels: { color: textColor } },
                title: { display: false }
            }
        }
    });

    // Render Daily Chart
    const ctxDaily = document.getElementById('dailyChart').getContext('2d');
    if (dailyChartInstance) dailyChartInstance.destroy();

    dailyChartInstance = new Chart(ctxDaily, {
        type: 'bar',
        data: {
            labels: dailyLabels,
            datasets: [{
                label: 'Hours Worked',
                data: dailyData,
                backgroundColor: '#4285f4'
            }]
        },
        options: {
            responsive: true,
            scales: {
                y: {
                    beginAtZero: true,
                    grid: { color: gridColor },
                    ticks: { color: textColor }
                },
                x: {
                    grid: { display: false },
                    ticks: { color: textColor }
                }
            },
            plugins: {
                legend: { labels: { color: textColor } }
            }
        }
    });
}
