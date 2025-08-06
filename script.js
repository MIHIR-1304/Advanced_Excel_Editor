async function saveSpreadsheet() {
  try {

    const table = document.getElementById('dataTable');
    const headers = Array.from(table.querySelectorAll('th')).map(th => th.textContent);
    const rows = Array.from(table.querySelectorAll('tbody tr')).map(row => 
      Array.from(row.querySelectorAll('td')).map(td => td.textContent)
    );

    
    const payload = {
      filename: `spreadsheet_${new Date().getTime()}`,
      data: { headers, rows }
    };

    console.log('Saving as Excel:', payload);

    
    const response = await fetch('http://localhost:3000/api/save-excel', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

   
    if (!response.ok) {
      const error = await response.text();
      throw new Error(error || 'Excel save failed');
    }

    const result = await response.json();
    alert(`Success! Excel file saved as ${result.filename}`);
    
    
    window.open(`http://localhost:3000/downloads/${result.filename}`, '_blank');
    
  } catch (error) {
    console.error('Save error:', error);
    alert(`Save failed: ${error.message}`);
  }
}


document.getElementById('save-btn').addEventListener('click', saveSpreadsheet);
async function testConnection() {
  try {
    const response = await axios.get('http://localhost:3000/api/test');
    console.log("Connection test:", response.data);
  } catch (error) {
    console.error("Connection failed:", error);
  }
}
testConnection();



function getCellData() {
  
  const cells = {};
  document.querySelectorAll('.cell').forEach(cell => {
    cells[cell.id] = cell.value;
  });
  return cells;
}

const API_URL = 'http://localhost:3000/api';


async function testBackendConnection() {
  try {
    const response = await axios.get(`${API_URL}/test`);
    console.log("Backend says:", response.data.message);
    alert("Backend connection successful!");
  } catch (error) {
    console.error("Connection failed:", error);
    alert("Backend connection failed. Check console for details.");
  }
}


window.addEventListener('DOMContentLoaded', () => {
  testBackendConnection();
  

});
document.addEventListener('DOMContentLoaded', function() {
   
    const excelFile = document.getElementById('excelFile');
    const loadBtn = document.getElementById('loadBtn');
    const saveBtn = document.getElementById('saveBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    const downloadPdfBtn = document.getElementById('downloadPdfBtn');
    const resetBtn = document.getElementById('resetBtn');
    const newSheetBtn = document.getElementById('newSheetBtn');
    const searchInput = document.getElementById('searchInput');
    const clearSearch = document.getElementById('clearSearch');
    const dataTable = document.getElementById('dataTable');
    const rowCount = document.getElementById('rowCount');
    const matchCount = document.getElementById('matchCount');
    const loading = document.getElementById('loading');
    const loadingText = document.getElementById('loadingText');
    const currentSheet = document.getElementById('currentSheet');
    const prevSheetBtn = document.getElementById('prevSheetBtn');
    const nextSheetBtn = document.getElementById('nextSheetBtn');
    const statusMessage = document.getElementById('statusMessage');
    const selectedCount = document.getElementById('selectedCount');
    const currentCell = document.getElementById('currentCell');
    const cellFormula = document.getElementById('cellFormula');
    const sortColumn = document.getElementById('sortColumn');
    const sortDirection = document.getElementById('sortDirection');
    const applySortBtn = document.getElementById('applySortBtn');
    const chartType = document.getElementById('chartType');
    const xAxisColumn = document.getElementById('xAxisColumn');
    const yAxisColumn = document.getElementById('yAxisColumn');
    const generateChartBtn = document.getElementById('generateChartBtn');
    const chartContainer = document.getElementById('chartContainer');
    const dataChart = document.getElementById('dataChart');
    const columnFilters = document.getElementById('columnFilters');
    const applyFilterBtn = document.getElementById('applyFilterBtn');
    const clearFilterBtn = document.getElementById('clearFilterBtn');

    
    let workbook = null;
    let sheets = [];
    let currentSheetIndex = 0;
    let originalData = [];
    let currentData = [];
    let originalHeaders = [];
    let currentMatches = 0;
    let isModified = false;
    let selectedCells = new Set();
    let currentChart = null;
    let activeFilterValues = {};

    
    function init() {
        
        loadBtn.addEventListener('click', loadExcelFile);
        saveBtn.addEventListener('click', saveChanges);
        downloadBtn.addEventListener('click', downloadExcel);
        downloadPdfBtn.addEventListener('click', exportToPdf);
        resetBtn.addEventListener('click', resetChanges);
        newSheetBtn.addEventListener('click', createNewSheet);
        searchInput.addEventListener('input', handleSearch);
        clearSearch.addEventListener('click', clearSearchInput);
        prevSheetBtn.addEventListener('click', showPreviousSheet);
        nextSheetBtn.addEventListener('click', showNextSheet);
        applySortBtn.addEventListener('click', applySorting);
        generateChartBtn.addEventListener('click', generateChart);
        applyFilterBtn.addEventListener('click', applyFilters);
        clearFilterBtn.addEventListener('click', clearFilters);

        
        const tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
        tooltipTriggerList.map(function (tooltipTriggerEl) {
            return new bootstrap.Tooltip(tooltipTriggerEl);
        });

        
        const tabEls = document.querySelectorAll('button[data-bs-toggle="tab"]');
        tabEls.forEach(tabEl => {
            tabEl.addEventListener('shown.bs.tab', function (event) {
                if (event.target.id === 'chart-tab' && currentChart) {
                    currentChart.resize();
                }
            });
        });
    }

   
    function loadExcelFile() {
        if (!excelFile.files.length) {
            showStatus('Please select an Excel file first.', 'error');
            return;
        }

        showLoading('Processing Excel file...');
        
        setTimeout(() => {
            try {
                const file = excelFile.files[0];
                const reader = new FileReader();

                reader.onload = function(e) {
                    try {
                        const data = new Uint8Array(e.target.result);
                        workbook = XLSX.read(data, { type: 'array', cellDates: true });
                        
                        // Get all sheet names
                        sheets = workbook.SheetNames;
                        currentSheetIndex = 0;
                        
                        // Load first sheet
                        loadSheetData(0);
                        
                        
                        saveBtn.disabled = false;
                        downloadBtn.disabled = false;
                        downloadPdfBtn.disabled = false;
                        resetBtn.disabled = false;
                        isModified = false;
                        
                        // Update sheet navigation
                        updateSheetNavigation();
                        
                        showStatus('File loaded successfully!', 'success');
                    } catch (error) {
                        console.error('Error processing file:', error);
                        showStatus('Error processing Excel file. Please check the console.', 'error');
                    } finally {
                        hideLoading();
                    }
                };

                reader.onerror = function() {
                    hideLoading();
                    showStatus('Error reading file. Please try again.', 'error');
                };

                reader.readAsArrayBuffer(file);
            } catch (error) {
                hideLoading();
                console.error('Error:', error);
                showStatus('An error occurred. Please check the console.', 'error');
            }
        }, 100);
    }

    // Load data from a specific sheet
    function loadSheetData(sheetIndex) {
        if (!workbook || sheetIndex < 0 || sheetIndex >= sheets.length) return;
        
        currentSheetIndex = sheetIndex;
        const sheetName = sheets[sheetIndex];
        currentSheet.textContent = sheetName;
        
        const worksheet = workbook.Sheets[sheetName];
        
        // Get raw data including formatting
        const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: "" });
        
        // First row contains headers
        originalHeaders = rawData.length > 0 ? rawData[0] : [];
        
        // Remove header row from data
        rawData.shift();
        
        // Store original and current data
        originalData = rawData.map(row => [...row]);
        currentData = rawData.map(row => [...row]);
        
        // Display data with original formatting
        displayData(currentData);
        
        // Initialize sorting and charting options
        initializeDataTools();
    }

    // Display data with original formatting
    function displayData(data, searchTerm = '') {
        // Clear previous data
        dataTable.innerHTML = '';
        selectedCells.clear();
        updateSelectedCount();
        
        // Create table header with original headers
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        
        originalHeaders.forEach((header, colIndex) => {
            const th = document.createElement('th');
            th.textContent = header;
            
            // Add column selection
            th.addEventListener('click', () => {
                selectColumn(colIndex);
            });
            
            headerRow.appendChild(th);
        });
        
        thead.appendChild(headerRow);
        dataTable.appendChild(thead);

        // Create table body
        const tbody = document.createElement('tbody');
        currentMatches = 0;

        data.forEach((row, rowIndex) => {
            const tr = document.createElement('tr');
            let rowHasMatch = false;
            
            row.forEach((cellValue, colIndex) => {
                const td = document.createElement('td');
                td.classList.add('editable');
                
                // Apply special formatting classes based on header
                const header = originalHeaders[colIndex];
                if (header && (header.toLowerCase().includes('date') || isExcelDate(cellValue))) {
                    td.classList.add('excel-date');
                }
                if (header && header.toLowerCase().includes('%')) {
                    td.classList.add('excel-percent');
                }
                
                // Check if this cell matches search term
                const displayValue = String(cellValue);
                if (searchTerm && displayValue.toLowerCase().includes(searchTerm.toLowerCase())) {
                    rowHasMatch = true;
                    // Highlight matching text
                    const highlightedValue = displayValue.replace(
                        new RegExp(searchTerm, 'gi'),
                        match => `<span class="highlight">${match}</span>`
                    );
                    td.innerHTML = highlightedValue;
                } else {
                    td.textContent = displayValue;
                }
                
               
                td.addEventListener('click', function(e) {
                    if (e.shiftKey) {
                        
                        toggleCellSelection(td, rowIndex, colIndex);
                    } else {
                        
                        clearSelection();
                        toggleCellSelection(td, rowIndex, colIndex);
                        makeCellEditable(td, rowIndex, colIndex);
                    }
                });
                
                tr.appendChild(td);
            });
            
            
            if (!searchTerm || rowHasMatch) {
                tbody.appendChild(tr);
                if (rowHasMatch) currentMatches++;
            }
        });

        dataTable.appendChild(tbody);
        rowCount.textContent = `${data.length} rows`;
        
        if (searchTerm) {
            matchCount.textContent = `${currentMatches} matches found`;
            matchCount.style.color = currentMatches > 0 ? 'green' : 'red';
        } else {
            matchCount.textContent = '0 matches found';
            matchCount.style.color = '#6c757d';
        }
    }

    // Make a cell editable
    function makeCellEditable(td, rowIndex, colIndex) {
        const originalValue = td.textContent;
        
        // Update formula bar
        currentCell.textContent = `${getColumnLetter(colIndex)}${rowIndex + 1}`;
        cellFormula.textContent = originalValue;
        
        // Create input element
        const input = document.createElement('input');
        input.type = 'text';
        input.value = originalValue;
        input.className = 'form-control';
        
        // Style input to fit cell
        input.style.width = '100%';
        input.style.height = '100%';
        input.style.border = 'none';
        input.style.padding = '8px';
        
        // Replace cell content with input
        td.innerHTML = '';
        td.appendChild(input);
        input.focus();
        
        // Handle saving the edited value
        const saveEdit = () => {
            const newValue = input.value;
            td.innerHTML = newValue;
            cellFormula.textContent = newValue;
            
            // Update the data if value changed
            if (newValue !== originalValue) {
                currentData[rowIndex][colIndex] = newValue;
                isModified = true;
                td.classList.add('modified-cell');
                
                // Reapply formatting classes
                const header = originalHeaders[colIndex];
                if (header && (header.toLowerCase().includes('date') || isExcelDate(newValue))) {
                    td.classList.add('excel-date');
                }
                if (header && header.toLowerCase().includes('%')) {
                    td.classList.add('excel-percent');
                }
                
                // Make cell editable again
                td.addEventListener('click', function(e) {
                    if (e.shiftKey) {
                        toggleCellSelection(td, rowIndex, colIndex);
                    } else {
                        clearSelection();
                        toggleCellSelection(td, rowIndex, colIndex);
                        makeCellEditable(td, rowIndex, colIndex);
                    }
                });
            }
        };
        
        // Save on Enter or click outside
        input.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                saveEdit();
            }
        });
        
        input.addEventListener('blur', saveEdit);
    }

    // Select/deselect a cell
    function toggleCellSelection(td, rowIndex, colIndex) {
        const cellId = `${rowIndex}-${colIndex}`;
        if (selectedCells.has(cellId)) {
            selectedCells.delete(cellId);
            td.classList.remove('cell-highlight');
        } else {
            selectedCells.add(cellId);
            td.classList.add('cell-highlight');
        }
        updateSelectedCount();
    }

    // Select entire column
    function selectColumn(colIndex) {
        clearSelection();
        const cells = dataTable.querySelectorAll(`td:nth-child(${colIndex + 1})`);
        cells.forEach((td, rowIndex) => {
            const cellId = `${rowIndex}-${colIndex}`;
            selectedCells.add(cellId);
            td.classList.add('cell-highlight');
        });
        updateSelectedCount();
    }

    // Clear all cell selections
    function clearSelection() {
        selectedCells.forEach(cellId => {
            const [rowIndex, colIndex] = cellId.split('-');
            const td = dataTable.querySelector(`tr:nth-child(${parseInt(rowIndex) + 1}) td:nth-child(${parseInt(colIndex) + 1})`);
            if (td) td.classList.remove('cell-highlight');
        });
        selectedCells.clear();
        updateSelectedCount();
    }

    // Update selected cell count display
    function updateSelectedCount() {
        selectedCount.textContent = `${selectedCells.size} ${selectedCells.size === 1 ? 'cell' : 'cells'} selected`;
    }

    // Helper function to get column letter from index
    function getColumnLetter(colIndex) {
        let letter = '';
        while (colIndex >= 0) {
            letter = String.fromCharCode(65 + (colIndex % 26)) + letter;
            colIndex = Math.floor(colIndex / 26) - 1;
        }
        return letter;
    }

    // Helper function to detect Excel dates
    function isExcelDate(value) {
        return typeof value === 'string' && 
              (value.match(/\d{1,2}-\d{1,2}-\d{4}/) || 
               value.match(/\d{4}-\d{1,2}-\d{1,2}/));
    }

    // Search functionality
    function handleSearch() {
        const searchTerm = this.value.trim();
        if (searchTerm) {
            displayData(currentData, searchTerm);
        } else {
            displayData(currentData);
        }
    }

    // Clear search input
    function clearSearchInput() {
        searchInput.value = '';
        displayData(currentData);
        searchInput.focus();
    }

    // Save changes back to the workbook
    function saveChanges() {
        if (!isModified) {
            showStatus('No changes to save.', 'info');
            return;
        }

        try {
            // Create new sheet with current data
            const newSheet = XLSX.utils.aoa_to_sheet([originalHeaders, ...currentData]);
            
            // Update workbook
            workbook.Sheets[sheets[currentSheetIndex]] = newSheet;
            
            // Update original data
            originalData = currentData.map(row => [...row]);
            isModified = false;
            
            // Remove modified cell highlights
            dataTable.querySelectorAll('.modified-cell').forEach(td => {
                td.classList.remove('modified-cell');
            });
            
            showStatus('Changes saved successfully!', 'success');
        } catch (error) {
            console.error('Error saving changes:', error);
            showStatus('Error saving changes. Please check the console.', 'error');
        }
    }

    // Download Excel file
    function downloadExcel() {
        if (!workbook || currentData.length === 0) {
            showStatus('No data to download.', 'error');
            return;
        }

        try {
            // Create new sheet with current data
            const newSheet = XLSX.utils.aoa_to_sheet([originalHeaders, ...currentData]);
            
            // Update workbook
            const newWorkbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWorkbook, newSheet, sheets[currentSheetIndex]);
            
            // Download
            XLSX.writeFile(newWorkbook, `${sheets[currentSheetIndex]}_data.xlsx`);
            showStatus('Excel file downloaded successfully!', 'success');
        } catch (error) {
            console.error('Error downloading file:', error);
            showStatus('Error downloading file. Please check the console.', 'error');
        }
    }

    // Export to PDF
    function exportToPdf() {
        if (!workbook || currentData.length === 0) {
            showStatus('No data to export.', 'error');
            return;
        }

        showLoading('Generating PDF...');
        
        setTimeout(() => {
            try {
                const { jsPDF } = window.jspdf;
                const doc = new jsPDF();
                
                // Add title
                doc.setFontSize(18);
                doc.text(`Project Management Data - ${sheets[currentSheetIndex]}`, 14, 15);
                
                // Prepare data for PDF
                const pdfData = [originalHeaders, ...currentData];
                
                // Generate table
                doc.autoTable({
                    head: [originalHeaders],
                    body: currentData,
                    startY: 20,
                    styles: {
                        fontSize: 8,
                        cellPadding: 2
                    },
                    headerStyles: {
                        fillColor: [52, 152, 219],
                        textColor: 255,
                        fontStyle: 'bold'
                    },
                    alternateRowStyles: {
                        fillColor: [245, 245, 245]
                    }
                });
                
                // Save PDF
                doc.save(`${sheets[currentSheetIndex]}_data.pdf`);
                showStatus('PDF exported successfully!', 'success');
            } catch (error) {
                console.error('Error generating PDF:', error);
                showStatus('Error generating PDF. Please check the console.', 'error');
            } finally {
                hideLoading();
            }
        }, 100);
    }

    // Reset changes
    function resetChanges() {
        if (!isModified) {
            showStatus('No changes to reset.', 'info');
            return;
        }

        if (confirm('Are you sure you want to discard all changes?')) {
            currentData = originalData.map(row => [...row]);
            displayData(currentData);
            isModified = false;
            showStatus('Changes reset successfully.', 'success');
        }
    }

    // Create new sheet
    function createNewSheet() {
        const sheetName = prompt('Enter name for new sheet:', `Sheet${sheets.length + 1}`);
        if (!sheetName) return;
        
        if (sheets.includes(sheetName)) {
            showStatus('Sheet name already exists!', 'error');
            return;
        }
        
        // Add new sheet to workbook
        const newSheet = XLSX.utils.aoa_to_sheet([[]]);
        workbook.Sheets[sheetName] = newSheet;
        sheets.push(sheetName);
        currentSheetIndex = sheets.length - 1;
        
        // Load new sheet
        loadSheetData(currentSheetIndex);
        updateSheetNavigation();
        showStatus(`New sheet "${sheetName}" created.`, 'success');
    }

    // Show previous sheet
    function showPreviousSheet() {
        if (currentSheetIndex > 0) {
            loadSheetData(currentSheetIndex - 1);
            updateSheetNavigation();
        }
    }

    // Show next sheet
    function showNextSheet() {
        if (currentSheetIndex < sheets.length - 1) {
            loadSheetData(currentSheetIndex + 1);
            updateSheetNavigation();
        }
    }

    // Update sheet navigation buttons
    function updateSheetNavigation() {
        prevSheetBtn.disabled = currentSheetIndex <= 0;
        nextSheetBtn.disabled = currentSheetIndex >= sheets.length - 1;
    }

    // Initialize data tools (sort, chart, filter)
    function initializeDataTools() {
        // Clear existing options
        sortColumn.innerHTML = '<option value="">Select column to sort</option>';
        xAxisColumn.innerHTML = '<option value="">Select X-Axis</option>';
        yAxisColumn.innerHTML = '<option value="">Select Y-Axis</option>';
        columnFilters.innerHTML = '';
        activeFilterValues = {};
        
        // Add options for each column
        originalHeaders.forEach((header, index) => {
            // Sort dropdown
            const sortOption = document.createElement('option');
            sortOption.value = index;
            sortOption.textContent = header;
            sortColumn.appendChild(sortOption);
            
            // Chart axis dropdowns
            const xOption = document.createElement('option');
            xOption.value = index;
            xOption.textContent = header;
            xAxisColumn.appendChild(xOption);
            
            const yOption = document.createElement('option');
            yOption.value = index;
            yOption.textContent = header;
            yAxisColumn.appendChild(yOption);
            
            // Column filters
            const filterDiv = document.createElement('div');
            filterDiv.className = 'col-md-4 mb-3';
            
            const filterLabel = document.createElement('label');
            filterLabel.className = 'form-label';
            filterLabel.textContent = header;
            
            const filterInput = document.createElement('input');
            filterInput.type = 'text';
            filterInput.className = 'form-control';
            filterInput.placeholder = `Filter ${header}`;
            filterInput.dataset.column = index;
            
            filterDiv.appendChild(filterLabel);
            filterDiv.appendChild(filterInput);
            columnFilters.appendChild(filterDiv);
            
            // Store filter value when changed
            filterInput.addEventListener('input', function() {
                activeFilterValues[index] = this.value;
            });
        });
    }

    // Apply sorting
    function applySorting() {
        const columnIndex = parseInt(sortColumn.value);
        const direction = sortDirection.value;
        
        if (isNaN(columnIndex)) {
            showStatus('Please select a column to sort.', 'error');
            return;
        }
        
        try {
            const sortedData = [...currentData];
            sortedData.sort((a, b) => {
                const valA = a[columnIndex];
                const valB = b[columnIndex];
                
                // Handle empty values
                if (valA === undefined || valA === '') return direction === 'asc' ? 1 : -1;
                if (valB === undefined || valB === '') return direction === 'asc' ? -1 : 1;
                
                // Try numeric comparison first
                const numA = parseFloat(valA);
                const numB = parseFloat(valB);
                if (!isNaN(numA) && !isNaN(numB)) {
                    return direction === 'asc' ? numA - numB : numB - numA;
                }
                
                // Fall back to string comparison
                return direction === 'asc' 
                    ? String(valA).localeCompare(String(valB))
                    : String(valB).localeCompare(String(valA));
            });
            
            currentData = sortedData;
            displayData(currentData);
            isModified = true;
            showStatus(`Data sorted by ${originalHeaders[columnIndex]} (${direction}).`, 'success');
        } catch (error) {
            console.error('Error sorting data:', error);
            showStatus('Error sorting data. Please check the console.', 'error');
        }
    }

    // Apply filters
    function applyFilters() {
        const activeFilters = Object.entries(activeFilterValues)
            .filter(([_, value]) => value && value.trim() !== '');
        
        if (activeFilters.length === 0) {
            displayData(originalData);
            showStatus('No active filters applied.', 'info');
            return;
        }
        
        const filteredData = originalData.filter(row => {
            return activeFilters.every(([colIndex, filterValue]) => {
                const cellValue = String(row[colIndex] || '').toLowerCase();
                return cellValue.includes(filterValue.toLowerCase());
            });
        });
        
        currentData = filteredData;
        displayData(currentData);
        showStatus(`Applied ${activeFilters.length} filter(s).`, 'success');
    }

    // Clear all filters
    function clearFilters() {
        document.querySelectorAll('#columnFilters input').forEach(input => {
            input.value = '';
        });
        activeFilterValues = {};
        currentData = [...originalData];
        displayData(currentData);
        showStatus('All filters cleared.', 'success');
    }

    // Generate chart
    function generateChart() {
        const xColIndex = parseInt(xAxisColumn.value);
        const yColIndex = parseInt(yAxisColumn.value);
        const type = chartType.value;
        
        if (isNaN(xColIndex) || isNaN(yColIndex)) {
            showStatus('Please select both X and Y axes.', 'error');
            return;
        }
        
        try {
            // Destroy previous chart if exists
            if (currentChart) {
                currentChart.destroy();
            }
            
            // Prepare chart data
            const labels = currentData.map(row => row[xColIndex]);
            const dataValues = currentData.map(row => {
                const val = row[yColIndex];
                return typeof val === 'number' ? val : parseFloat(val) || 0;
            });
            
            // Create chart
            const ctx = dataChart.getContext('2d');
            currentChart = new Chart(ctx, {
                type: type,
                data: {
                    labels: labels,
                    datasets: [{
                        label: originalHeaders[yColIndex],
                        data: dataValues,
                        backgroundColor: getChartColors(type, labels.length),
                        borderColor: '#2c3e50',
                        borderWidth: 1
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        title: {
                            display: true,
                            text: `${originalHeaders[yColIndex]} by ${originalHeaders[xColIndex]}`,
                            font: {
                                size: 16
                            }
                        },
                        legend: {
                            position: type === 'pie' || type === 'doughnut' ? 'right' : 'top'
                        }
                    },
                    scales: type === 'pie' || type === 'doughnut' ? {} : {
                        y: {
                            beginAtZero: true
                        }
                    }
                }
            });
            
            showStatus('Chart generated successfully!', 'success');
        } catch (error) {
            console.error('Error generating chart:', error);
            showStatus('Error generating chart. Please check the console.', 'error');
        }
    }

    // Helper function to generate chart colors
    function getChartColors(type, count) {
        if (type === 'pie' || type === 'doughnut') {
            return [
                '#3498db', '#2ecc71', '#e74c3c', '#f39c12', '#9b59b6',
                '#1abc9c', '#d35400', '#34495e', '#16a085', '#c0392b'
            ].slice(0, count);
        } else {
            return ['rgba(52, 152, 219, 0.7)'];
        }
    }

    // Show loading indicator
    function showLoading(message) {
        loadingText.textContent = message;
        loading.style.display = 'flex';
    }

    // Hide loading indicator
    function hideLoading() {
        loading.style.display = 'none';
    }

    // Show status message
    function showStatus(message, type = 'info') {
        statusMessage.textContent = message;
        statusMessage.style.color = {
            'info': '#17a2b8',
            'success': '#28a745',
            'error': '#dc3545',
            'warning': '#ffc107'
        }[type];
    }
    

    // Initialize the application
    init();
});