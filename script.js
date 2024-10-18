const file1Input = document.getElementById('file1');
const file2Input = document.getElementById('file2');
const column1Select = document.getElementById('column1');
const column2Select = document.getElementById('column2');
const updateTableSelect = document.getElementById('update-table');
const compareBtn = document.getElementById('compare-btn');
const downloadBtn = document.getElementById('download-btn');

// ...

// ...

// 文件读取函数
function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = e.target.result;
            const workbook = XLSX.read(data, { type: 'binary' });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const dataJson = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            const columns = dataJson[0];
            const rows = dataJson.slice(1);
            resolve({ columns, rows });
        };
        reader.readAsBinaryString(file);
    });
}

// ...

// 对比函数
function compareFiles(file1, file2, column1, column2, updateTable) {
    return new Promise((resolve, reject) => {
        readExcelFile(file1).then((data1) => {
            readExcelFile(file2).then((data2) => {
                const result = [];
                if (updateTable === '1') {
                    data1.rows.forEach((row) => {
                        const value = row[column1];
                        const exists = data2.rows.find((row) => row[column2] === value);
                        result.push([value, exists ? '存在' : '不存在']);
                    });
                } else {
                    data2.rows.forEach((row) => {
                        const value = row[column2];
                        const exists = data1.rows.find((row) => row[column1] === value);
                        result.push([value, exists ? '存在' : '不存在']);
                    });
                }
                resolve(result);
            });
        });
    });
}

// 下载函数
function downloadFile(data, filename) {
    const workbook = XLSX.utils.book_new();
    const sheetName = '结果';
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(data), sheetName);
    XLSX.writeFile(workbook, filename);
}

// 页面逻辑
file1Input.addEventListener('change', (e) => {
    const file = e.target.files[0];
    readExcelFile(file).then((data) => {
        const columns = data.columns;
        column1Select.innerHTML = '';
        columns.forEach((column) => {
            const option = document.createElement('option');
            option.text = column;
            option.value = column;
            column1Select.appendChild(option);
        });
        column1Select.disabled = false;
    });
});

file2Input.addEventListener('change', (e) => {
    const file = e.target.files[0];
    readExcelFile(file).then((data) => {
        const columns = data.columns;
        column2Select.innerHTML = '';
        columns.forEach((column) => {
            const option = document.createElement('option');
            option.text = column;
            option.value = column;
            column2Select.appendChild(option);
        });
        column2Select.disabled = false;
    });
});

compareBtn.addEventListener('click', () => {
    const file1 = file1Input.files[0];
    const file2 = file2Input.files[0];
    const column1 = column1Select.value;
    const column2 = column2Select.value;
    const updateTable = updateTableSelect.value;
    compareFiles(file1, file2, column1, column2, updateTable).then((result) => {
        downloadBtn.disabled = false;
        downloadBtn.addEventListener('click', () => {
            downloadFile(result, '结果.xlsx');
        });
    });
});

// 初始化页面
compareBtn.disabled = true;
downloadBtn.disabled = true;