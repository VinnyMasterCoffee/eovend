document.addEventListener('DOMContentLoaded', function() {
    const today = new Date();
    document.getElementById('inventoryDate').value = today.toISOString().split('T')[0];
    loadTemplate();
    document.getElementById('downloadBtn').addEventListener('click', downloadInventoryWithExcelJS);
});

const NON_EDITABLE_ITEMS = [
    "Кофе Poetti Espresso Bravo 1кг зерно",
    "Палочки размешиватели GlobalCups 105 мм",
    "Стакан бумажный Формация WAKE ME CUP D80 300мл 50шт/уп",
    "Кофе Live Coffee (Санта Ричи) зерно",
    "Смесь сухая ARISTOCRAT Клубника 1кг"
];

const SIMPLIFIED_NAMES = {
    "Вода Нила Спрингс 19л": "Вода",
    "Горячий шоколад ARISTOCRAT ШВЕЙЦАРСКИЙ гранулы 500г": "Шоколад",
    "Капучино ARISTOCRAT Mokka Toffee 1000г": "Toffee",
    "Капучино TORINO Irish Cream 1кг": "Irish",
    "Кофе Жардин Пьяцца Арабика 1кг зерно": "Кофе Jardin",
    "Крышка пластиковая 80мм без клапана Global Cups": "Крышки",
    "Сладкий сахар в пакетах 1кг": "Сахар",
    "Стакан бумажный GlobalCups D70 150мл 100шт/уп": "Стаканы",
    "Сухое молоко гранул. \"AlpenMilch Плюс\" 1000г": "Молоко",
    "Сухое молоко МАЛИНА 1000г": "Малина",
    "Смесь сухая ARISTOCRAT Цитрус 1кг": "Цитрус"
};

// Функция для нормализации строк (удаление лишних пробелов)
function normalizeString(str) {
    return str.trim().replace(/\s+/g, ' ');
}

async function loadTemplate() {
    try {
        const response = await fetch('/eovend/templates/invent.xlsx');
        const arrayBuffer = await response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, {type: 'array'});
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, {header: 1});
        
        let headerRow = -1;
        for (let i = 0; i < jsonData.length; i++) {
            if (jsonData[i][0] === "№ п/п") {
                headerRow = i;
                break;
            }
        }
        
        if (headerRow === -1) throw new Error("Не удалось найти заголовки");
        
        const tableBody = document.getElementById('inventoryItems');
        tableBody.innerHTML = '';
        
        for (let i = headerRow + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row[0] || isNaN(row[0])) continue;

            // Пропускаем нередактируемые товары
            const originalName = normalizeString(row[2]);
            if (NON_EDITABLE_ITEMS.some(item => normalizeString(item) === originalName)) continue;

            // Получаем упрощённое название
            const simplifiedName = SIMPLIFIED_NAMES[originalName] || originalName;

            const tr = document.createElement('tr');
            
            // Название товара (упрощённое)
            const nameTd = document.createElement('td');
            nameTd.textContent = simplifiedName;
            tr.appendChild(nameTd);

            // Поле ввода
            const tdInput = document.createElement('td');
            const input = document.createElement('input');
            input.type = 'number';
            input.min = '0';
            input.dataset.rowIndex = i - headerRow - 1;
            input.dataset.originalName = originalName; // Сохраняем оригинальное название

            tdInput.appendChild(input);
            tr.appendChild(tdInput);
            
            tableBody.appendChild(tr);
        }
        
    } catch (error) {
        console.error('Ошибка загрузки шаблона:', error);
        alert('Ошибка загрузки шаблона. Пожалуйста, попробуйте позже.');
    }
}

async function downloadInventoryWithExcelJS() {
    const dateInput = document.getElementById('inventoryDate');
    const routeNumber = document.getElementById('routeNumber').value;
    const carNumber = document.getElementById('carNumber').value;
    
    if (!dateInput.value || !routeNumber || !carNumber) {
        alert('Пожалуйста, заполните все обязательные поля');
        return;
    }
    
    try {
        const response = await fetch('/eovend/templates/invent.xlsx');
        const arrayBuffer = await response.arrayBuffer();
        
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(arrayBuffer);
        
        for (let sheetIndex = 0; sheetIndex < workbook.worksheets.length; sheetIndex++) {
            const worksheet = workbook.worksheets[sheetIndex];
            
            const dateCell = findDateCell(worksheet);
            if (dateCell) {
                const date = new Date(dateInput.value);
                const monthNames = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 
                                  'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
                dateCell.value = `"___${date.getDate()}___" ${monthNames[date.getMonth()]} ${date.getFullYear()} г.`;
            }
            
            // Получаем все строки из Excel, чтобы найти соответствия
            const excelRows = [];
            let rowIndex = 6; // Начальная строка с данными
            while (true) {
                const nameCell = worksheet.getCell(`C${rowIndex}`);
                if (!nameCell.value) break;
                
                excelRows.push({
                    rowNumber: rowIndex,
                    name: nameCell.value
                });
                rowIndex++;
            }
            
            const inputs = document.querySelectorAll('#inventoryItems input');
            inputs.forEach((input) => {
                const originalName = input.dataset.originalName;
                const excelRow = excelRows.find(row => row.name === originalName);
                
                if (excelRow) {
                    const row = excelRow.rowNumber;
                    
                    const factCell = worksheet.getCell(`E${row}`);
                    factCell.value = input.value ? parseInt(input.value) : null;
                    
                    const checkCell = worksheet.getCell(`G${row}`);
                    checkCell.value = { formula: `EXACT(F${row},E${row})`, result: false };
                    
                    ['A', 'B', 'C', 'D', 'E', 'F', 'G'].forEach(col => {
                        const cell = worksheet.getCell(`${col}${row}`);
                        cell.border = {
                            top: {style: 'thin'},
                            left: {style: 'thin'},
                            bottom: {style: 'thin'},
                            right: {style: 'thin'}
                        };
                    });
                }
            });
        }
        
        const date = new Date(dateInput.value);
        const monthNames = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
                          'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'];
        const fileName = `Инвентаризация ${monthNames[date.getMonth()]} ${date.getFullYear()} кофе К${routeNumber} ${carNumber}.xlsx`;
        
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
        saveAs(blob, fileName);
        
    } catch (error) {
        console.error('Ошибка при создании файла:', error);
        alert('Произошла ошибка при создании файла. Пожалуйста, попробуйте позже.');
    }
}

function findDateCell(worksheet) {
    for (let row = 1; row <= 10; row++) {
        for (let col = 1; col <= 7; col++) {
            const cell = worksheet.getCell(row, col);
            if (cell.text && cell.text.includes('мая 2025 г.')) {
                return cell;
            }
        }
    }
    return null;
}
