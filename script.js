const EXCEL_TEMPLATE_PATH = 'https://eovendcoffee.github.io/invent/templates/invent.xlsx';
const DEBUG_MODE = true; // Поставьте true, если нужно включить логирование

const MULTIPLIERS = {
    "Стакан бумажный Формация WAKE ME CUP D80 300мл 50шт/уп": 40,
    "Крышка пластиковая 80мм без клапана Global Cups": 100
};

const ROUTE_TO_CAR_MAPPING = {
    "1": "У840УЕ33",
    "2": "У179УК33",
    "3": "У493УК33",
    "4": "Т372УХ33",
    "5": "К645ХС33",
    "6": "М906ХВ33",
    "7": "М281ХО33",
    "8": "М332ХТ33"
};

const NON_EDITABLE_ITEMS = [
    "Кофе Poetti Espresso Bravo 1кг зерно",
    "Палочки размешиватели GlobalCups 105 мм",
    "Стакан бумажный GlobalCups D70 150мл 100шт/уп",
    "Кофе Live Coffee (Санта Ричи) зерно",
    "Смесь сухая ARISTOCRAT Клубника 1кг"
];

const SIMPLIFIED_NAMES = {
    "Вода Нила Спрингс 19л": "Вода",
    "Горячий шоколад ARISTOCRAT ШВЕЙЦАРСКИЙ гранулы 500г": "Шоколад",
    "Капучино ARISTOCRAT Mokka Toffee 1000г": "Toffee",
    "Капучино TORINO Irish Cream 1кг": "Irish",
    "Кофе Жардин Пьяцца Арабика 1кг зерно": "Кофе Jardin",
    "Крышка пластиковая 80мм без клапана Global Cups": "Крышки (упаковка = 100 шт)",
    "Сладкий сахар в пакетах 1кг": "Сахар",
    "Стакан бумажный Формация WAKE ME CUP D80 300мл 50шт/уп": "Стаканы (упаковка = 40 шт)",
    "Сухое молоко гранул. \"AlpenMilch Плюс\" 1000г": "Молоко",
    "Сухое молоко МАЛИНА 1000г": "Малина",
    "Смесь сухая ARISTOCRAT Цитрус 1кг": "Цитрус"
};

document.addEventListener('DOMContentLoaded', function() {
    const today = getCurrentDateInMSK();
    document.getElementById('inventoryDate').value = today.toISOString().split('T')[0];
    loadTemplate();
    
    const routeSelect = document.getElementById('routeNumber');
    const carNumberInput = document.getElementById('carNumber');
    
    // Один обработчик на изменение маршрута
    routeSelect.addEventListener('change', function() {
        if (this.value) {
            this.classList.remove('error');
            document.getElementById('routeNumberError').style.display = 'none';
        }
        carNumberInput.value = ROUTE_TO_CAR_MAPPING[this.value] || '';
    });
    
    document.getElementById('downloadBtn').addEventListener('click', downloadInventoryWithExcelJS);
});


/*document.getElementById('routeNumber').addEventListener('change', function() {
    if (this.value) {
        this.classList.remove('error');
        document.getElementById('routeNumberError').style.display = 'none';
    }
    // Автозаполнение номера машины
    document.getElementById('carNumber').value = ROUTE_TO_CAR_MAPPING[this.value] || '';
});*/

function getCurrentDateInMSK() {
    const now = new Date();
    // Москва UTC+3, поэтому добавляем 3 часа к UTC
    const mskOffset = 3 * 60 * 60 * 1000; // 3 часа в миллисекундах
    const mskTime = new Date(now.getTime() + mskOffset);
    return new Date(mskTime.setUTCHours(0, 0, 0, 0)); // Устанавливаем начало дня
}

// Функция для нормализации строк (удаление лишних пробелов)
function normalizeString(str) {
    return str.trim().replace(/\s+/g, ' ');
}

async function loadTemplate() {
    if (DEBUG_MODE) {
        console.log('Загрузка шаблона...', EXCEL_TEMPLATE_PATH);
    }
    try {
        const response = await fetch(EXCEL_TEMPLATE_PATH);
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

            const originalName = normalizeString(row[2]);
            
            // Пропускаем только действительно нередактируемые товары
            if (NON_EDITABLE_ITEMS.some(item => normalizeString(item) === originalName)) continue;

            // Получаем упрощённое название
            const simplifiedName = SIMPLIFIED_NAMES[originalName] || originalName;
            const tr = document.createElement('tr');
            
            const nameTd = document.createElement('td');
            nameTd.textContent = simplifiedName;
            tr.appendChild(nameTd);

            const tdInput = document.createElement('td');
            const input = document.createElement('input');
            input.type = 'number';
            input.min = '0';
            input.dataset.rowIndex = i - headerRow - 1;
            input.dataset.originalName = originalName;

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
    if (DEBUG_MODE) {
        console.log('Начало скачивания...');
    }
    const dateInput = document.getElementById('inventoryDate');
    const routeSelect = document.getElementById('routeNumber');
    const carNumberInput = document.getElementById('carNumber');
    
    // Сбрасываем предыдущие ошибки
    document.querySelectorAll('.error').forEach(el => el.classList.remove('error'));
    document.querySelectorAll('.error-message').forEach(el => el.style.display = 'none');
    
    let isValid = true;
    
    // Проверка даты
    if (!dateInput.value) {
        dateInput.classList.add('error');
        isValid = false;
    }
    
    // Проверка маршрута
    if (!routeSelect.value) {
        routeSelect.classList.add('error');
        const errorElement = document.getElementById('routeNumberError');
        errorElement.style.display = 'block';
        errorElement.style.opacity = '1';
        isValid = false;
    }
    
    // Проверка номера машины
    if (!carNumberInput.value) {
        carNumberInput.classList.add('error');
        isValid = false;
    }
    
    if (!isValid) {
        const firstError = document.querySelector('.error');
        if (firstError) {
            firstError.scrollIntoView({ behavior: 'smooth', block: 'center' });
            firstError.focus();
        }
        return;
    }
    
    try {
        const response = await fetch(EXCEL_TEMPLATE_PATH);
        const arrayBuffer = await response.arrayBuffer();
        
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(arrayBuffer);
        
        for (let sheetIndex = 0; sheetIndex < workbook.worksheets.length; sheetIndex++) {
            const worksheet = workbook.worksheets[sheetIndex];
            
            const dateCell = findDateCell(worksheet, dateInput.value);
            if (dateCell) {
                const date = new Date(dateInput.value);
                const monthNames = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 
                                  'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
                dateCell.value = `"___${date.getDate()}___" ${monthNames[date.getMonth()]} ${date.getFullYear()} г.`;
            }
            
            // Получаем все строки из Excel
            const excelRows = [];
            let rowIndex = 6;
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
                const excelRow = excelRows.find(row => normalizeString(row.name) === originalName);
                
                if (excelRow) {
                    const row = excelRow.rowNumber;
                    let value = input.value ? parseInt(input.value) : null;
                    
                    // Применяем множитель для специальных товаров
                    if (MULTIPLIERS[originalName] && value !== null) {
                        value *= MULTIPLIERS[originalName];
                    }
                    
                    const factCell = worksheet.getCell(`E${row}`);
                    factCell.value = value;
                    
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
        
        const routeNumber = routeSelect.value;
        const carNumber = carNumberInput.value;
        
        const fileName = `Инвентаризация ${monthNames[date.getMonth()]} ${date.getFullYear()} кофе К${routeNumber} ${carNumber}.xlsx`;
        
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
        saveAs(blob, fileName);
        
    } catch (error) {
        console.error('Ошибка при создании файла:', error);
        alert('Произошла ошибка при создании файла. Пожалуйста, попробуйте позже.');
    }
}

function formatDate(date) {
    // Убедимся, что дата корректно обрабатывается как МСК
    const mskOffset = 3 * 60 * 60 * 1000;
    const mskDate = new Date(date.getTime() + mskOffset);
    
    const day = String(mskDate.getUTCDate()).padStart(2, '0');
    const month = String(mskDate.getUTCMonth() + 1).padStart(2, '0');
    const year = mskDate.getUTCFullYear();
    return `${day}.${month}.${year}`;
}

function findDateCell(worksheet, dateInput) {
    const date = new Date(dateInput);
    // Добавляем смещение для МСК
    const mskOffset = 3 * 60 * 60 * 1000;
    const mskDate = new Date(date.getTime() + mskOffset);
    
    const monthNames = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 
                      'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];

    for (let row = 1; row <= 10; row++) {
        for (let col = 1; col <= 7; col++) {
            const cell = worksheet.getCell(row, col);
            if (!cell.text) continue;

            if (cell.text.includes('мая 2025 г.')) {
                cell.value = `"___${mskDate.getUTCDate()}___" ${monthNames[mskDate.getUTCMonth()]} ${mskDate.getUTCFullYear()} г.`;
            }

            if (cell.text.trim().startsWith('1с,') && cell.text.includes('.')) {
                cell.value = `1с, ${formatDate(mskDate)}`;
            }
        }
    }
}
// Плавное смещение при фокусе на поле ввода
document.querySelectorAll('#inventoryTable input').forEach(input => {
    input.addEventListener('focus', function() {
        // Вычисляем позицию элемента относительно верха страницы
        const elementPosition = this.getBoundingClientRect().top;
        // Вычисляем текущую позицию прокрутки
        const offsetPosition = elementPosition + window.pageYOffset - 100; // 100px отступ сверху
        
        // Плавная прокрутка к элементу
        window.scrollTo({
            top: offsetPosition,
            behavior: 'smooth'
        });
    });
});

