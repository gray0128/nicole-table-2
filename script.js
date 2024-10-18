// script.js

document.getElementById('file1').addEventListener('change', function() {
    handleFileUpload(this, 'column1');
});

document.getElementById('file2').addEventListener('change', function() {
    handleFileUpload(this, 'column2');
});

function handleFileUpload(input, columnSelectId) {
    const file = input.files[0];
    if (!file) return;

    readExcel(file).then(data => {
        const columnSelect = document.getElementById(columnSelectId);
        columnSelect.innerHTML = data[0].map(col => `<option value="${col}">${col}</option>`).join('');
        checkReadyToCompare();
    });
}

function checkReadyToCompare() {
    const column1 = document.getElementById('column1').value;
    const column2 = document.getElementById('column2').value;
    const file1 = document.getElementById('file1').files[0];
    const file2 = document.getElementById('file2').files[0];

    if (column1 && column2 && file1 && file2) {
        document.getElementById('compareBtn').classList.remove('hidden');
    }
}

document.getElementById('compareBtn').addEventListener('click', function() {
    const file1 = document.getElementById('file1').files[0];
    const file2 = document.getElementById('file2').files[0];
    const column1 = document.getElementById('column1').value;
    const column2 = document.getElementById('column2').value;
    const updateTarget = document.getElementById('updateTarget').value;

    this.disabled = true;
    this.textContent = '处理中...';
    
    Promise.all([readExcel(file1), readExcel(file2)]).then(([data1, data2]) => {
        const columnData1 = getColumnData(data1, column1);
        const columnData2 = getColumnData(data2, column2);
        
        const results = updateTarget === 'table1' ? compareData(columnData1, columnData2) : compareData(columnData2, columnData1);
        
        const updatedData = updateTarget === 'table1' ? updateResults(data1, results) : updateResults(data2, results);
        
        // 保存更新后的数据以便下载时使用
        document.getElementById('downloadBtn').dataset.updatedData = JSON.stringify(updatedData);
        document.getElementById('downloadBtn').dataset.filename = updateTarget === 'table1' ? '表1更新.xlsx' : '表2更新.xlsx';
        
        this.disabled = false;
        this.textContent = '对比';
        document.getElementById('downloadBtn').classList.remove('hidden');
    });
});

document.getElementById('downloadBtn').addEventListener('click', function() {
    const updatedData = JSON.parse(this.dataset.updatedData);
    const filename = this.dataset.filename;
    downloadExcel(updatedData, filename);
});

function readExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const result = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
            resolve(result);
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

function getColumnData(data, column) {
    const columnIndex = data[0].indexOf(column);
    return data.slice(1).map(row => row[columnIndex]);
}

function compareData(data1, data2) {
    return data1.map(item => data2.includes(item) ? '是' : '否');
}

function updateResults(data, results) {
    const updatedData = data.map((row, index) => {
        if (index === 0) {
            // 确保在标题行添加"对比结果"
            row.push('对比结果');
        } else {
            // 确保在每一行的末尾添加结果
            while (row.length < data[0].length - 1) {
                row.push('');
            }
            row.push(results[index - 1]);
        }
        return row;
    });
    return updatedData;
}

function downloadExcel(data, filename) {
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    XLSX.writeFile(workbook, filename);
}
