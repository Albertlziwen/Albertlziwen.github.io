let currentData = [];
const pageSize = 50; // 每页显示50行
let currentPage = 1;
let fileReader = null;

function handleFileUpload() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    const progressContainer = document.getElementById('progressContainer');
    const progressBar = document.getElementById('progressBar');
    const progressText = document.getElementById('progressText');
    
    if (file) {
        // 显示进度条
        progressContainer.style.display = 'block';
        progressBar.style.width = '0%';
        progressText.textContent = '0%';
        
        // 如果存在之前的 FileReader，中止它
        if (fileReader) {
            fileReader.abort();
        }
        
        fileReader = new FileReader();
        
        // 监听读取进度
        fileReader.onprogress = function(e) {
            if (e.lengthComputable) {
                const progress = Math.round((e.loaded / e.total) * 100);
                progressBar.style.width = progress + '%';
                progressText.textContent = progress + '%';
            }
        };
        
        // 读取完成后的处理
        fileReader.onload = function(e) {
            progressBar.style.width = '100%';
            progressText.textContent = '100%';
            
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                currentData = XLSX.utils.sheet_to_json(firstSheet);
                
                // 生成列控制
                generateColumnControls(currentData[0]);
                // 显示第一页数据
                displayData(1);
                
                // 3秒后隐藏进度条
                setTimeout(() => {
                    progressContainer.style.display = 'none';
                }, 3000);
                
                // 启用导出按钮
                document.querySelector('.export-btn').disabled = false;
                
            } catch (error) {
                alert('文件处理出错：' + error.message);
                progressContainer.style.display = 'none';
            }
        };
        
        // 错误处理
        fileReader.onerror = function() {
            alert('文件读取错误！');
            progressContainer.style.display = 'none';
        };
        
        // 开始读取文件
        fileReader.readAsArrayBuffer(file);
    }
}

function generateColumnControls(firstRow) {
    const controlsDiv = document.getElementById('columnControls');
    controlsDiv.innerHTML = '';
    
    // 确保至少选中一列时导出按钮可用
    const updateExportButton = () => {
        const checkedColumns = document.querySelectorAll('#columnControls input:checked');
        document.querySelector('.export-btn').disabled = checkedColumns.length === 0;
    };

    Object.keys(firstRow).forEach(column => {
        const label = document.createElement('label');
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.checked = true;
        checkbox.value = column;
        checkbox.addEventListener('change', () => {
            refreshTable();
            updateExportButton();
        });
        
        label.appendChild(checkbox);
        label.appendChild(document.createTextNode(column));
        controlsDiv.appendChild(label);
    });
}

function getVisibleColumns() {
    const checkboxes = document.querySelectorAll('#columnControls input:checked');
    return Array.from(checkboxes).map(cb => cb.value);
}

function displayData(page) {
    currentPage = page;
    const visibleColumns = getVisibleColumns();
    const start = (page - 1) * pageSize;
    const end = start + pageSize;
    const pageData = currentData.slice(start, end);
    
    const table = document.createElement('table');
    
    // 创建表头
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    visibleColumns.forEach(column => {
        const th = document.createElement('th');
        th.textContent = column;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);
    
    // 创建表体
    const tbody = document.createElement('tbody');
    pageData.forEach(row => {
        const tr = document.createElement('tr');
        visibleColumns.forEach(column => {
            const td = document.createElement('td');
            td.textContent = row[column] || '';
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    
    // 更新表格容器
    const container = document.getElementById('tableContainer');
    container.innerHTML = '';
    container.appendChild(table);
    
    // 更新分页控件
    updatePagination();
}

function updatePagination() {
    const totalPages = Math.ceil(currentData.length / pageSize);
    const paginationDiv = document.getElementById('pagination');
    paginationDiv.innerHTML = '';
    
    // 添加上一页按钮
    const prevButton = document.createElement('button');
    prevButton.textContent = '上一页';
    prevButton.disabled = currentPage === 1;
    prevButton.onclick = () => displayData(currentPage - 1);
    paginationDiv.appendChild(prevButton);
    
    // 添加页码
    for (let i = 1; i <= totalPages; i++) {
        const pageButton = document.createElement('button');
        pageButton.textContent = i;
        pageButton.className = i === currentPage ? 'active' : '';
        pageButton.onclick = () => displayData(i);
        paginationDiv.appendChild(pageButton);
    }
    
    // 添加下一页按钮
    const nextButton = document.createElement('button');
    nextButton.textContent = '下一页';
    nextButton.disabled = currentPage === totalPages;
    nextButton.onclick = () => displayData(currentPage + 1);
    paginationDiv.appendChild(nextButton);
}

function refreshTable() {
    displayData(currentPage);
}

// 添加导出功能
function exportToExcel() {
    // 如果没有数据，直接返回
    if (!currentData || currentData.length === 0) {
        alert('没有可导出的数据！');
        return;
    }

    try {
        // 获取当前可见的列
        const visibleColumns = getVisibleColumns();
        
        // 创建要导出的数据数组
        const exportData = currentData.map(row => {
            const newRow = {};
            visibleColumns.forEach(column => {
                newRow[column] = row[column] || '';
            });
            return newRow;
        });

        // 创建工作簿
        const wb = XLSX.utils.book_new();
        // 创建工作表
        const ws = XLSX.utils.json_to_sheet(exportData);
        
        // 设置列宽
        const colWidths = {};
        visibleColumns.forEach(col => {
            colWidths[col] = {wch: Math.max(col.length * 2, 10)};
        });
        ws['!cols'] = visibleColumns.map(col => colWidths[col]);

        // 将工作表添加到工作簿
        XLSX.utils.book_append_sheet(wb, ws, "数据导出");

        // 生成文件名
        const date = new Date();
        const fileName = `数据导出_${date.getFullYear()}${(date.getMonth()+1).toString().padStart(2,'0')}${date.getDate().toString().padStart(2,'0')}_${date.getHours().toString().padStart(2,'0')}${date.getMinutes().toString().padStart(2,'0')}.xlsx`;

        // 导出文件
        XLSX.writeFile(wb, fileName);

    } catch (error) {
        console.error('导出错误：', error);
        alert('导出失败：' + error.message);
    }
} 