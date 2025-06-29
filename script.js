document.addEventListener('DOMContentLoaded', function() {
    // Устанавливаем сегодняшнюю дату по умолчанию
    const today = new Date();
    document.getElementById('inventoryDate').value = today.toISOString().split('T')[0];
    
    // Загружаем шаблон и заполняем таблицу
    loadTemplate();
    
    // Обработчик кнопки скачивания
    document.getElementById('downloadBtn').addEventListener('click', downloadInventoryWithExcelJS);
});

// Список позиций, которые нельзя редактировать
const NON_EDITABLE_ITEMS = [
    "Кофе Poetti Espresso Bravo 1кг зерно",
    "Палочки размешиватели GlobalCups 105 мм",
    "Стакан бумажный Формация WAKE ME CUP D80 300мл 50шт/уп",
    "Кофе Live Coffee (Санта Ричи) зерно",
    "Смесь сухая ARISTOCRAT Клубника 1кг"
];

async function loadTemplate() {
    try {
        const response = await fetch('/eovend/templates/invent.xlsx');
        const arrayBuffer = await response.arrayBuffer();
        
        // Используем SheetJS только для чтения данных (не для записи)
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, {type: 'array'});
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, {header: 1});
        
        // Находим строку с заголовками
        let headerRow = -1;
        for (let i = 0; i < jsonData.length; i++) {
            if (jsonData[i][0] === "№ п/п") {
                headerRow = i;
                break;
            }
        }
        
        if (headerRow === -1) throw new Error("Не удалось найти заголовки");
        
        // Заполняем таблицу на странице
        const tableBody = document.getElementById('inventoryItems');
        tableBody.innerHTML = '';
        
        for (let i = headerRow + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row[0] || isNaN(row[0])) continue;
            
            // Проверяем, является ли текущая позиция нередактируемой
            const isEditable = !NON_EDITABLE_ITEMS.includes(row[2]);
            
            const tr = document.createElement('tr');
            
            ['0', '1', '2', '3'].forEach(col => {
                const td = document.createElement('td');
                td.textContent = row[col] || '';
                tr.appendChild(td);
            });
            
            const tdInput = document.createElement('td');
            const input = document.createElement('input');
            input.type = 'number';
            input.min = '0';
            input.dataset.rowIndex = i - headerRow - 1;
            
            // Если позиция не редактируемая, делаем поле readonly
            if (!isEditable) {
                input.readOnly = true;
                input.placeholder = 'Не редактируется';
                input.style.backgroundColor = '#f0f0f0';
                input.style.cursor = 'not-allowed';
            }
            
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
        // Загружаем шаблон с помощью ExcelJS
        const response = await fetch('/eovend/templates/invent.xlsx');
        const arrayBuffer = await response.arrayBuffer();
        
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(arrayBuffer);
        
        // Обновляем оба листа
        for (let sheetIndex = 0; sheetIndex < workbook.worksheets.length; sheetIndex++) {
            const worksheet = workbook.worksheets[sheetIndex];
            
            // Обновляем дату в шаблоне
            const dateCell = findDateCell(worksheet);
            if (dateCell) {
                const date = new Date(dateInput.value);
                const monthNames = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 
                                  'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
                dateCell.value = `"___${date.getDate()}___" ${monthNames[date.getMonth()]} ${date.getFullYear()} г.`;
            }
            
            // Заполняем данные
            const inputs = document.querySelectorAll('#inventoryItems input:not([readonly])');
            inputs.forEach((input, index) => {
                const row = 6 + index; // Начальная строка данных
                
                // Фактические остатки (колонка E)
                const factCell = worksheet.getCell(`E${row}`);
                factCell.value = input.value ? parseInt(input.value) : null;
                
                // Формула сверки (колонка G)
                const checkCell = worksheet.getCell(`G${row}`);
                checkCell.value = { formula: `EXACT(F${row},E${row})`, result: false };
                
                // Применяем стили к ячейкам
                ['A', 'B', 'C', 'D', 'E', 'F', 'G'].forEach(col => {
                    const cell = worksheet.getCell(`${col}${row}`);
                    cell.border = {
                        top: {style: 'thin'},
                        left: {style: 'thin'},
                        bottom: {style: 'thin'},
                        right: {style: 'thin'}
                    };
                });
            });
        }
        
        // Генерируем имя файла
        const date = new Date(dateInput.value);
        const monthNames = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
                          'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'];
        const fileName = `Инвентаризация ${monthNames[date.getMonth()]} ${date.getFullYear()} кофе К${routeNumber} ${carNumber}.xlsx`;
        
        // Скачиваем файл
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
        saveAs(blob, fileName);
        
    } catch (error) {
        console.error('Ошибка при создании файла:', error);
        alert('Произошла ошибка при создании файла. Пожалуйста, попробуйте позже.');
    }
}

// Вспомогательная функция для поиска ячейки с датой
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
