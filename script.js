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

// Словарь для упрощенных названий
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
            
            const isEditable = !NON_EDITABLE_ITEMS.includes(row[2]);
            const simplifiedName = SIMPLIFIED_NAMES[row[2]] || row[2];
            
            const tr = document.createElement('tr');
            
            // № п/п (скрытая колонка)
            const td1 = document.createElement('td');
            td1.className = 'hidden-column';
            td1.textContent = row[0];
            tr.appendChild(td1);
            
            // Код (скрытая колонка)
            const td2 = document.createElement('td');
            td2.className = 'hidden-column';
            td2.textContent = row[1];
            tr.appendChild(td2);
            
            // Название (упрощенное)
            const td3 = document.createElement('td');
            td3.textContent = simplifiedName;
            tr.appendChild(td3);
            
            // Ед. изм. (скрытая колонка)
            const td4 = document.createElement('td');
            td4.className = 'hidden-column';
            td4.textContent = row[3];
            tr.appendChild(td4);
            
            // Количество
            const td5 = document.createElement('td');
            const input = document.createElement('input');
            input.type = 'number';
            input.min = '0';
            input.dataset.rowIndex = i - headerRow - 1;
            
            if (!isEditable) {
                input.readOnly = true;
                input.placeholder = 'Не редактируется';
                input.className = 'non-editable';
            }
            
            td5.appendChild(input);
            tr.appendChild(td5);
            
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
            const inputs = document.querySelectorAll('#inventoryItems input:not(.non-editable)');
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
