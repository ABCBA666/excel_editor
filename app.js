// Excel数据存储
let workbook = null;
let currentSheetIndex = 0;
let sheetData = [];
let columnWidths = {}; // 存储每列的宽度 {sheetIndex: {colIndex: width}}
let dragState = { type: null, fromIndex: null };
let selectedRows = new Set();

// DOM元素
const fileInput = document.getElementById('fileInput');
const uploadSection = document.getElementById('uploadSection');
const editorSection = document.getElementById('editorSection');
const uploadBox = document.getElementById('uploadBox');
const sheetTabs = document.getElementById('sheetTabs');
const dataTable = document.getElementById('dataTable');
const tableHead = document.getElementById('tableHead');
const tableBody = document.getElementById('tableBody');
const downloadBtn = document.getElementById('downloadBtn');
const newFileBtn = document.getElementById('newFileBtn');
const cellInfo = document.getElementById('cellInfo');
const sheetInfo = document.getElementById('sheetInfo');

// 初始化
fileInput.addEventListener('change', handleFileSelect);
downloadBtn.addEventListener('click', downloadExcel);
newFileBtn.addEventListener('click', resetApp);
document.getElementById('addRowBtn').addEventListener('click', addNewRow);
document.getElementById('addColBtn').addEventListener('click', addNewColumn);

// 拖拽上传
uploadBox.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadBox.classList.add('dragover');
});

uploadBox.addEventListener('dragleave', () => {
    uploadBox.classList.remove('dragover');
});

uploadBox.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadBox.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0 && isExcelFile(files[0])) {
        processFile(files[0]);
    }
});

uploadBox.addEventListener('click', () => {
    fileInput.click();
});

// 处理文件选择
function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file && isExcelFile(file)) {
        processFile(file);
    }
}

// 检查是否为Excel文件
function isExcelFile(file) {
    const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel'
    ];
    return validTypes.includes(file.type) || 
           file.name.endsWith('.xlsx') || 
           file.name.endsWith('.xls');
}

// 处理Excel文件
function processFile(file) {
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, { type: 'array' });
            
            // 解析所有sheet
            sheetData = workbook.SheetNames.map((name, index) => {
                const worksheet = workbook.Sheets[name];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                    header: 1, 
                    defval: '',
                    raw: false 
                });
                return {
                    name: name,
                    index: index,
                    data: jsonData
                };
            });
            
            // 显示编辑器
            showEditor();
            // 显示第一个sheet
            switchSheet(0);
        } catch (error) {
            alert('文件读取失败：' + error.message);
            console.error(error);
        }
    };
    
    reader.readAsArrayBuffer(file);
}

// 显示编辑器
function showEditor() {
    uploadSection.style.display = 'none';
    editorSection.style.display = 'flex';
    renderSheetTabs();
}

// 渲染Sheet标签
function renderSheetTabs() {
    sheetTabs.innerHTML = '';
    sheetData.forEach((sheet, index) => {
        const tab = document.createElement('div');
        tab.className = `sheet-tab ${index === currentSheetIndex ? 'active' : ''}`;
        tab.textContent = sheet.name;
        tab.addEventListener('click', () => switchSheet(index));
        sheetTabs.appendChild(tab);
    });
    updateSheetInfo();
}

// 切换Sheet
function switchSheet(index) {
    currentSheetIndex = index;
    renderSheetTabs();
    renderTable();
    // 恢复该Sheet的列宽设置
    applyColumnWidths();
}

// 应用列宽设置
function applyColumnWidths() {
    if (!columnWidths[currentSheetIndex]) return;
    
    const widths = columnWidths[currentSheetIndex];
    Object.keys(widths).forEach(colIndex => {
        const width = widths[colIndex];
        const allCells = document.querySelectorAll(`th[data-col="${colIndex}"], td[data-col="${colIndex}"]`);
        allCells.forEach(cell => {
            if (!cell.classList.contains('row-number')) {
                cell.style.width = width + 'px';
                cell.style.minWidth = width + 'px';
            }
        });
    });
}

// 检查列是否有数据
function hasColumnData(sheet, colIndex) {
    for (let row = 0; row < sheet.data.length; row++) {
        if (sheet.data[row] && sheet.data[row][colIndex] !== undefined && 
            sheet.data[row][colIndex] !== null && 
            String(sheet.data[row][colIndex]).trim() !== '') {
            return true;
        }
    }
    return false;
}

// 检查行是否有数据
function hasRowData(sheet, rowIndex) {
    if (!sheet.data[rowIndex]) return false;
    for (let col = 0; col < sheet.data[rowIndex].length; col++) {
        if (sheet.data[rowIndex][col] !== undefined && 
            sheet.data[rowIndex][col] !== null && 
            String(sheet.data[rowIndex][col]).trim() !== '') {
            return true;
        }
    }
    return false;
}

// 获取有数据的列索引（包括空列，如果它们存在）
function getValidColumns(sheet) {
    const maxCols = sheet.data.length > 0 
        ? Math.max(...sheet.data.map(row => row.length || 0))
        : 0;
    const validCols = [];
    // 先添加有数据的列
    for (let i = 0; i < maxCols; i++) {
        if (hasColumnData(sheet, i)) {
            validCols.push(i);
        }
    }
    // 如果没有任何有数据的列，至少显示第一列（如果存在）
    if (validCols.length === 0 && maxCols > 0) {
        validCols.push(0);
    }
    return validCols;
}

// 获取有数据的行索引（包括空行，如果它们存在）
function getValidRows(sheet) {
    const validRows = [];
    // 先添加有数据的行
    for (let i = 0; i < sheet.data.length; i++) {
        if (hasRowData(sheet, i)) {
            validRows.push(i);
        }
    }
    // 如果没有任何有数据的行，至少显示第一行（表头行）
    if (validRows.length === 0 && sheet.data.length > 0) {
        validRows.push(0);
    }
    return validRows;
}

function moveRowInSheet(sheet, fromIndex, toIndex) {
    if (!sheet || !sheet.data) return;
    if (fromIndex === toIndex) return;
    if (fromIndex < 0 || toIndex < 0) return;
    if (fromIndex >= sheet.data.length || toIndex >= sheet.data.length) return;
    const rows = sheet.data;
    const row = rows.splice(fromIndex, 1)[0];
    rows.splice(toIndex, 0, row);
}

function moveColumnInSheet(sheet, fromIndex, toIndex) {
    if (!sheet || !sheet.data) return;
    if (fromIndex === toIndex) return;
    if (fromIndex < 0 || toIndex < 0) return;
    for (let r = 0; r < sheet.data.length; r++) {
        const row = sheet.data[r] || [];
        const maxIndex = Math.max(fromIndex, toIndex);
        while (row.length <= maxIndex) {
            row.push('');
        }
        const cell = row.splice(fromIndex, 1)[0];
        row.splice(toIndex, 0, cell);
        sheet.data[r] = row;
    }
}

function toggleRowSelection(rowIndex) {
    const currentSheet = sheetData[currentSheetIndex];
    if (!currentSheet) return;
    if (selectedRows.has(rowIndex)) {
        selectedRows.delete(rowIndex);
    } else {
        selectedRows.add(rowIndex);
    }
    renderTable();
}

function toggleAllRowsSelection() {
    const currentSheet = sheetData[currentSheetIndex];
    if (!currentSheet) return;
    const validRows = getValidRows(currentSheet);
    if (validRows.length === 0) return;
    const headerRowIndex = validRows.includes(0) ? 0 : (validRows.length > 0 ? validRows[0] : 0);
    const dataRows = validRows.filter(rowIndex => rowIndex !== headerRowIndex);
    if (dataRows.length === 0) return;
    const allSelected = dataRows.every(rowIndex => selectedRows.has(rowIndex));
    if (allSelected) {
        dataRows.forEach(rowIndex => selectedRows.delete(rowIndex));
    } else {
        dataRows.forEach(rowIndex => selectedRows.add(rowIndex));
    }
    renderTable();
}

// 渲染表格
function renderTable() {
    const currentSheet = sheetData[currentSheetIndex];
    if (!currentSheet || !currentSheet.data || currentSheet.data.length === 0) {
        tableHead.innerHTML = '<tr><th>无数据</th></tr>';
        tableBody.innerHTML = '<tr><td>此Sheet为空</td></tr>';
        return;
    }
    
    // 获取所有实际存在的列（包括空列），确保新增的列也能显示
    const maxCols = currentSheet.data.length > 0 
        ? Math.max(...currentSheet.data.map(row => (row && row.length) || 0), 0)
        : 0;
    
    // 获取有数据的列
    const dataCols = getValidColumns(currentSheet);
    
    // 合并有数据的列和所有存在的列
    const allCols = [];
    // 先添加有数据的列
    dataCols.forEach(col => {
        if (!allCols.includes(col)) {
            allCols.push(col);
        }
    });
    // 再添加所有存在的列（包括空列），确保新增的列也能显示
    for (let i = 0; i < maxCols; i++) {
        if (!allCols.includes(i)) {
            allCols.push(i);
        }
    }
    // 如果allCols为空，至少显示第一列
    if (allCols.length === 0) {
        allCols.push(0);
    }
    let validCols = allCols;
    
    // 获取所有实际存在的行（包括空行），确保新增的行也能显示
    const rowsWithData = getValidRows(currentSheet);
    const allRows = [];
    // 先添加有数据的行
    rowsWithData.forEach(row => {
        if (!allRows.includes(row)) {
            allRows.push(row);
        }
    });
    // 再添加所有存在的行（包括空行），确保新增的行也能显示
    for (let i = 0; i < currentSheet.data.length; i++) {
        if (!allRows.includes(i)) {
            allRows.push(i);
        }
    }
    // 确保表头行（索引0）在第一位
    allRows.sort((a, b) => a - b);
    let validRows = allRows.length > 0 ? allRows : [0];
    
    // 获取第一行作为表头（优先使用第一行，即使它可能为空）
    const headerRowIndex = validRows.includes(0) ? 0 : (validRows.length > 0 ? validRows[0] : 0);
    const headerRowData = currentSheet.data[headerRowIndex] || [];
    const allDataRowIndices = validRows.filter(rowIndex => rowIndex !== headerRowIndex);
    
    // 计算每列的初始宽度（基于内容）
    const calculatedWidths = calculateColumnWidths(currentSheet, validCols);
    
    // 渲染表头（使用第一行数据）
    tableHead.innerHTML = '';
    const headerRow = document.createElement('tr');
    // 添加行号列（表头行号列为空）
    const rowNumTh = document.createElement('th');
    rowNumTh.className = 'row-number';
    rowNumTh.textContent = '';
    const headerCircle = document.createElement('span');
    headerCircle.className = 'row-select-circle';
    if (allDataRowIndices.length > 0) {
        const allSelected = allDataRowIndices.every(rowIndex => selectedRows.has(rowIndex));
        const anySelected = allDataRowIndices.some(rowIndex => selectedRows.has(rowIndex));
        if (allSelected) {
            headerCircle.classList.add('selected');
        } else if (anySelected) {
            headerCircle.classList.add('partial');
        }
    }
    headerCircle.addEventListener('click', (e) => {
        e.stopPropagation();
        toggleAllRowsSelection();
    });
    rowNumTh.appendChild(headerCircle);
    headerRow.appendChild(rowNumTh);
    
    // 只添加有数据的列
    validCols.forEach((originalColIndex, displayColIndex) => {
        const th = document.createElement('th');
        const headerValue = headerRowData[originalColIndex] !== undefined && headerRowData[originalColIndex] !== null 
            ? String(headerRowData[originalColIndex]) 
            : '';
        th.textContent = headerValue;
        th.dataset.col = originalColIndex; // 保存原始列索引
        th.dataset.displayCol = displayColIndex; // 保存显示列索引
        
        // 设置列宽（优先使用保存的宽度，否则使用计算出的宽度）
        const savedWidth = getColumnWidth(currentSheetIndex, originalColIndex);
        const width = savedWidth || calculatedWidths[displayColIndex] || 120;
        th.style.width = width + 'px';
        th.style.minWidth = width + 'px';
        
        // 允许双击编辑表头
        th.addEventListener('dblclick', () => startEditHeader(th, originalColIndex));
        
        // 添加列宽调整器
        const resizer = document.createElement('div');
        resizer.className = 'column-resizer';
        resizer.addEventListener('mousedown', (e) => startResize(e, th, originalColIndex));
        th.appendChild(resizer);

        th.draggable = true;
        th.addEventListener('dragstart', (e) => {
            dragState.type = 'col';
            dragState.fromIndex = originalColIndex;
            if (e.dataTransfer) {
                e.dataTransfer.effectAllowed = 'move';
            }
        });
        th.addEventListener('dragover', (e) => {
            e.preventDefault();
        });
        th.addEventListener('drop', (e) => {
            e.preventDefault();
            if (dragState.type !== 'col') return;
            const currentSheet = sheetData[currentSheetIndex];
            moveColumnInSheet(currentSheet, dragState.fromIndex, originalColIndex);
            updateWorkbook();
            renderTable();
            applyColumnWidths();
            dragState.type = null;
            dragState.fromIndex = null;
        });
        
        headerRow.appendChild(th);
    });
    tableHead.appendChild(headerRow);
    
    // 渲染表格 body（只显示有数据的行，排除表头行）
    tableBody.innerHTML = '';
    let dataRows = validRows.filter(rowIndex => rowIndex !== headerRowIndex);
    let displayRowIndex = 0;
    
    // 如果没有任何数据行，至少显示一行空行用于编辑
    if (dataRows.length === 0) {
        // 确保至少有一行数据可以编辑
        if (currentSheet.data.length <= 1) {
            // 如果只有表头行，创建一个空数据行
            if (!currentSheet.data[1]) {
                const maxCols = validCols.length > 0 ? Math.max(...validCols) + 1 : 1;
                currentSheet.data[1] = new Array(maxCols).fill('');
            }
            dataRows = [1];
        } else {
            // 如果有数据但都是空的，显示所有数据行（从索引1开始）
            dataRows = [];
            for (let i = 1; i < currentSheet.data.length; i++) {
                dataRows.push(i);
            }
            // 如果还是没有，至少显示第一行数据行
            if (dataRows.length === 0 && currentSheet.data.length > 1) {
                dataRows = [1];
            }
        }
    }
    
    dataRows.forEach(dataRowIndex => {
        const tr = document.createElement('tr');
        const rowData = currentSheet.data[dataRowIndex] || [];
        
        // 添加行号列（从1开始）
        const rowNumTd = document.createElement('td');
        rowNumTd.className = 'row-number';
        rowNumTd.textContent = '';
        const rowCircle = document.createElement('span');
        rowCircle.className = 'row-select-circle';
        if (selectedRows.has(dataRowIndex)) {
            rowCircle.classList.add('selected');
        }
        rowCircle.addEventListener('click', (e) => {
            e.stopPropagation();
            toggleRowSelection(dataRowIndex);
        });
        rowNumTd.appendChild(rowCircle);
        const rowLabel = document.createElement('span');
        rowLabel.className = 'row-number-label';
        rowLabel.textContent = displayRowIndex + 1; // 行号从1开始
        rowNumTd.appendChild(rowLabel);
        rowNumTd.draggable = true;
        rowNumTd.addEventListener('dragstart', (e) => {
            dragState.type = 'row';
            dragState.fromIndex = dataRowIndex;
            if (e.dataTransfer) {
                e.dataTransfer.effectAllowed = 'move';
            }
        });
        tr.addEventListener('dragover', (e) => {
            e.preventDefault();
        });
        tr.addEventListener('drop', (e) => {
            e.preventDefault();
            if (dragState.type !== 'row') return;
            const currentSheet = sheetData[currentSheetIndex];
            moveRowInSheet(currentSheet, dragState.fromIndex, dataRowIndex);
            updateWorkbook();
            renderTable();
            applyColumnWidths();
            dragState.type = null;
            dragState.fromIndex = null;
        });
        if (selectedRows.has(dataRowIndex)) {
            tr.classList.add('row-selected');
        }
        tr.appendChild(rowNumTd);
        
        // 只添加有数据的列
        validCols.forEach((originalColIndex, displayColIndex) => {
            const td = document.createElement('td');
            const cellValue = rowData[originalColIndex] !== undefined && rowData[originalColIndex] !== null 
                ? String(rowData[originalColIndex]) 
                : '';
            td.textContent = cellValue;
            // 保存原始数据行索引和列索引（用于编辑时更新数据）
            td.dataset.row = dataRowIndex;
            td.dataset.col = originalColIndex;
            td.dataset.displayCol = displayColIndex;
            
            // 设置列宽（与表头保持一致）
            const savedWidth = getColumnWidth(currentSheetIndex, originalColIndex);
            const calculatedWidth = calculatedWidths[displayColIndex] || 120;
            const width = savedWidth || calculatedWidth;
            td.style.width = width + 'px';
            td.style.minWidth = width + 'px';
            
            // 添加编辑功能
            td.addEventListener('click', () => startEdit(td));
            td.addEventListener('dblclick', () => startEdit(td));
            
            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
        displayRowIndex++;
    });
    
    // 更新单元格信息
    updateCellInfo();
}

// 获取列名（A, B, C, ...）
function getColumnName(index) {
    let result = '';
    let num = index;
    while (num >= 0) {
        result = String.fromCharCode(65 + (num % 26)) + result;
        num = Math.floor(num / 26) - 1;
    }
    return result;
}

// 计算列宽（基于内容）
function calculateColumnWidths(sheet, cols) {
    const widths = [];
    const padding = 30; // 单元格内边距
    
    for (let col = 0; col < cols; col++) {
        let maxWidth = 80; // 最小宽度
        
        // 检查表头宽度
        if (sheet.data[0] && sheet.data[0][col] !== undefined) {
            const headerText = String(sheet.data[0][col] || '');
            const headerWidth = measureText(headerText) + padding;
            maxWidth = Math.max(maxWidth, headerWidth);
        }
        
        // 检查该列所有数据行的宽度（只检查前100行以提高性能）
        const checkRows = Math.min(sheet.data.length, 100);
        for (let row = 1; row < checkRows; row++) {
            if (sheet.data[row] && sheet.data[row][col] !== undefined) {
                const cellText = String(sheet.data[row][col] || '');
                const cellWidth = measureText(cellText) + padding;
                maxWidth = Math.max(maxWidth, cellWidth);
            }
        }
        
        // 限制最大宽度
        widths[col] = Math.min(maxWidth, 500);
    }
    
    return widths;
}

// 测量文本宽度
let measureElement = null;
function measureText(text) {
    if (!measureElement) {
        measureElement = document.createElement('span');
        measureElement.style.visibility = 'hidden';
        measureElement.style.position = 'absolute';
        measureElement.style.whiteSpace = 'nowrap';
        measureElement.style.fontSize = '14px';
        measureElement.style.fontFamily = '-apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif';
        document.body.appendChild(measureElement);
    }
    measureElement.textContent = text;
    return measureElement.offsetWidth;
}

// 获取列宽
function getColumnWidth(sheetIndex, colIndex) {
    if (!columnWidths[sheetIndex]) return null;
    return columnWidths[sheetIndex][colIndex] || null;
}

// 设置列宽
function setColumnWidth(sheetIndex, colIndex, width) {
    if (!columnWidths[sheetIndex]) {
        columnWidths[sheetIndex] = {};
    }
    columnWidths[sheetIndex][colIndex] = width;
}

// 开始调整列宽
function startResize(e, th, colIndex) {
    e.preventDefault();
    e.stopPropagation();
    
    const startX = e.pageX;
    const startWidth = th.offsetWidth;
    const sheetIndex = currentSheetIndex;
    
    // 创建遮罩层
    const overlay = document.createElement('div');
    overlay.className = 'resize-overlay';
    document.body.appendChild(overlay);
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
    
    const doResize = (e) => {
        const diff = e.pageX - startX;
        const newWidth = Math.max(50, startWidth + diff); // 最小宽度50px
        
        // 更新表头宽度
        th.style.width = newWidth + 'px';
        th.style.minWidth = newWidth + 'px';
        
        // 更新该列所有单元格的宽度
        const allCells = document.querySelectorAll(`th[data-col="${colIndex}"], td[data-col="${colIndex}"]`);
        allCells.forEach(cell => {
            if (!cell.classList.contains('row-number')) {
                cell.style.width = newWidth + 'px';
                cell.style.minWidth = newWidth + 'px';
            }
        });
        
        // 保存列宽
        setColumnWidth(sheetIndex, colIndex, newWidth);
    };
    
    const stopResize = () => {
        document.removeEventListener('mousemove', doResize);
        document.removeEventListener('mouseup', stopResize);
        document.body.removeChild(overlay);
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
    };
    
    document.addEventListener('mousemove', doResize);
    document.addEventListener('mouseup', stopResize);
}

// 开始编辑表头
function startEditHeader(th, col) {
    if (th.classList.contains('editing')) return;
    
    const currentValue = th.textContent;
    
    // 创建输入框
    const input = document.createElement('input');
    input.type = 'text';
    input.value = currentValue;
    input.className = 'cell-input';
    
    // 替换表头内容
    th.classList.add('editing');
    th.innerHTML = '';
    th.appendChild(input);
    input.focus();
    input.select();
    
    // 保存编辑
    const saveEdit = () => {
        const newValue = input.value;
        th.classList.remove('editing');
        th.textContent = newValue;
        
        // 更新数据（第一行，索引0）
        const currentSheet = sheetData[currentSheetIndex];
        if (!currentSheet.data[0]) {
            currentSheet.data[0] = [];
        }
        currentSheet.data[0][col] = newValue;
        
        // 更新workbook
        updateWorkbook();
        
        // 如果数据从空变为有值，重新渲染表格以显示新列
        if (newValue.trim() !== '') {
            renderTable();
        }
        
        updateCellInfo();
        updateSheetInfo();
    };
    
    // 取消编辑
    const cancelEdit = () => {
        th.classList.remove('editing');
        th.textContent = currentValue;
    };
    
    // 事件监听
    input.addEventListener('blur', saveEdit);
    input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            saveEdit();
        } else if (e.key === 'Escape') {
            e.preventDefault();
            cancelEdit();
        }
    });
}

// 开始编辑单元格
function startEdit(td) {
    if (td.classList.contains('editing')) return;
    
    const row = parseInt(td.dataset.row);
    const col = parseInt(td.dataset.col);
    const currentValue = td.textContent;
    
    // 创建输入框
    const input = document.createElement('input');
    input.type = 'text';
    input.value = currentValue;
    input.className = 'cell-input';
    
    // 替换单元格内容
    td.classList.add('editing');
    td.innerHTML = '';
    td.appendChild(input);
    input.focus();
    input.select();
    
    // 保存编辑
    const saveEdit = () => {
        const newValue = input.value;
        td.classList.remove('editing');
        td.textContent = newValue;
        
        // 更新数据
        const currentSheet = sheetData[currentSheetIndex];
        if (!currentSheet.data[row]) {
            currentSheet.data[row] = [];
        }
        currentSheet.data[row][col] = newValue;
        
        // 更新workbook
        updateWorkbook();
        
        // 如果数据从空变为有值，重新渲染表格以显示新行/列
        if (newValue.trim() !== '') {
            renderTable();
        }
        
        updateCellInfo();
        updateSheetInfo();
    };
    
    // 取消编辑
    const cancelEdit = () => {
        td.classList.remove('editing');
        td.textContent = currentValue;
    };
    
    // 事件监听
    input.addEventListener('blur', saveEdit);
    input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            saveEdit();
            // 移动到下一行
            const nextRow = td.parentElement.nextElementSibling;
            if (nextRow) {
                const nextCell = nextRow.children[col];
                if (nextCell) {
                    setTimeout(() => startEdit(nextCell), 10);
                }
            }
        } else if (e.key === 'Escape') {
            e.preventDefault();
            cancelEdit();
        } else if (e.key === 'Tab') {
            e.preventDefault();
            saveEdit();
            // 移动到下一列
            const nextCol = td.nextElementSibling;
            if (nextCol) {
                setTimeout(() => startEdit(nextCol), 10);
            } else {
                // 移动到下一行第一列
                const nextRow = td.parentElement.nextElementSibling;
                if (nextRow) {
                    const firstCell = nextRow.firstElementChild;
                    if (firstCell) {
                        setTimeout(() => startEdit(firstCell), 10);
                    }
                }
            }
        }
    });
}

// 更新Workbook
function updateWorkbook() {
    if (!workbook) return;
    
    sheetData.forEach((sheet, index) => {
        const worksheet = XLSX.utils.aoa_to_sheet(sheet.data);
        workbook.Sheets[sheet.name] = worksheet;
    });
}

// 下载Excel
function downloadExcel() {
    if (!workbook) {
        alert('没有可下载的数据');
        return;
    }
    
    // 更新workbook
    updateWorkbook();
    
    // 生成Excel文件
    const wbout = XLSX.write(workbook, { 
        bookType: 'xlsx', 
        type: 'array' 
    });
    
    // 创建Blob并下载
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'edited_excel.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// 重置应用
function resetApp() {
    if (confirm('确定要加载新文件吗？当前编辑将丢失。')) {
        workbook = null;
        currentSheetIndex = 0;
        sheetData = [];
        uploadSection.style.display = 'block';
        editorSection.style.display = 'none';
        fileInput.value = '';
    }
}

// 更新单元格信息
function updateCellInfo() {
    const selectedCell = document.querySelector('.editing');
    if (selectedCell) {
        const row = parseInt(selectedCell.dataset.row);
        const col = parseInt(selectedCell.dataset.col);
        const colName = getColumnName(col);
        // 如果是第一行（表头），显示为"表头"
        if (row === 0) {
            cellInfo.textContent = `编辑表头: ${colName}1`;
        } else {
            cellInfo.textContent = `编辑: ${colName}${row + 1}`;
        }
    } else {
        cellInfo.textContent = '就绪';
    }
}

// 更新Sheet信息
function updateSheetInfo() {
    const currentSheet = sheetData[currentSheetIndex];
    if (currentSheet) {
        const validCols = getValidColumns(currentSheet);
        const validRows = getValidRows(currentSheet);
        sheetInfo.textContent = `${currentSheet.name} | ${validRows.length} 行 × ${validCols.length} 列`;
    }
}

// 新增行
function addNewRow() {
    const currentSheet = sheetData[currentSheetIndex];
    if (!currentSheet) {
        alert('请先上传Excel文件');
        return;
    }
    
    // 获取当前最大列数（包括有数据的列）
    const validCols = getValidColumns(currentSheet);
    const maxCols = validCols.length > 0 
        ? Math.max(...validCols) + 1  // 使用最大有效列索引+1
        : (currentSheet.data.length > 0 
            ? Math.max(...currentSheet.data.map(row => (row && row.length) || 0))
            : 1);
    
    // 创建新行，确保列数与现有数据一致
    const newRow = new Array(maxCols).fill('');
    
    // 如果当前没有数据，先创建表头行
    if (currentSheet.data.length === 0) {
        currentSheet.data.push(new Array(maxCols).fill(''));
    }
    
    // 确保新行的列数与表头行一致
    if (currentSheet.data[0] && currentSheet.data[0].length > newRow.length) {
        newRow.length = currentSheet.data[0].length;
        newRow.fill('');
    }
    
    // 添加新行（在最后）
    currentSheet.data.push(newRow);
    
    // 更新workbook
    updateWorkbook();
    
    // 重新渲染表格（新增的空行会显示，因为renderTable会处理空行显示）
    renderTable();
    updateSheetInfo();
}

// 新增列
function addNewColumn() {
    const currentSheet = sheetData[currentSheetIndex];
    if (!currentSheet) {
        alert('请先上传Excel文件');
        return;
    }
    
    // 获取当前最大列数
    const maxCols = currentSheet.data.length > 0 
        ? Math.max(...currentSheet.data.map(row => (row && row.length) || 0))
        : 0;
    const newColIndex = maxCols;
    
    // 如果当前没有数据，先创建一行作为表头
    if (currentSheet.data.length === 0) {
        currentSheet.data.push(['']);
    }
    
    // 为每一行添加新列
    for (let i = 0; i < currentSheet.data.length; i++) {
        if (!currentSheet.data[i]) {
            currentSheet.data[i] = [];
        }
        // 确保数组长度足够，添加新列
        while (currentSheet.data[i].length <= newColIndex) {
            currentSheet.data[i].push('');
        }
    }
    
    // 更新workbook
    updateWorkbook();
    
    // 重新渲染表格（新增的空列会显示，因为renderTable会处理空列显示）
    renderTable();
    updateSheetInfo();
}

// 监听表格点击以更新信息
tableBody.addEventListener('click', (e) => {
    if (e.target.tagName === 'TD' && !e.target.classList.contains('editing')) {
        const row = parseInt(e.target.dataset.row);
        const col = parseInt(e.target.dataset.col);
        const colName = getColumnName(col);
        if (row === 0) {
            cellInfo.textContent = `选中表头: ${colName}1`;
        } else {
            cellInfo.textContent = `选中: ${colName}${row + 1}`;
        }
    }
});

// 监听表头点击以更新信息
tableHead.addEventListener('click', (e) => {
    if (e.target.tagName === 'TH') {
        const col = parseInt(e.target.dataset.col);
        const colName = getColumnName(col);
        cellInfo.textContent = `选中表头: ${colName}1`;
    }
});

