const fileInput = document.getElementById('file-input');
const barcodeInput = document.getElementById('barcode-input');
const generateButton = document.getElementById('generate-button');
const statusDiv = document.getElementById('status');

// Вешаем обработчик на клик по кнопке
generateButton.addEventListener('click', async () => {
    const file = fileInput.files[0];
    const startBarcode = barcodeInput.value.trim();

    if (!file) {
        updateStatus('Пожалуйста, выберите Excel файл.', 'error');
        return;
    }
    if (!startBarcode) {
        updateStatus('Пожалуйста, введите начальный штрихкод.', 'error');
        return;
    }

    updateStatus('Обработка файла...', 'info');
    generateButton.disabled = true;

    try {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                
                // Получаем первую страницу
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Конвертируем в JSON
                const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});
                
                // Фильтруем пустые строки
                const rows = jsonData.filter(row => row.length > 0 && row[0] !== undefined);
                
                // Обработка данных
                let currentBarcodeInt = parseInt(startBarcode);
                const barcodeLength = startBarcode.length;
                
                // Подготовка данных для "Номенклатура"
                const nomenklaturaData = [];
                rows.forEach(row => {
                    const barcode = currentBarcodeInt.toString().padStart(barcodeLength, '0');
                    nomenklaturaData.push({
                        "Штрихкод": barcode,
                        "Наименование": row[0],
                        "ЭтоУслуга": "Нет",
                        "ЕдиницаИзмерения": "шт",
                        "Ставка НДС": "Без НДС",
                        "Цена": row[2]
                    });
                    currentBarcodeInt++;
                });
                
                // Подготовка данных для "Поступление_товаров"
                const postuplenieData = rows.map(row => ({
                    "Номенклатура": row[0],
                    "Количество": row[1],
                    "Единица измерения": "шт",
                    "Цена": row[2],
                    "ЭтоУслуга": "Нет"
                }));
                
                // Генерация CSV
                const nomenklaturaCsv = convertToCsv(nomenklaturaData);
                const postuplenieCsv = convertToCsv(postuplenieData);
                
                // Скачивание файлов
                downloadZip(nomenklaturaCsv, postuplenieCsv);
                updateStatus('Архив 1c_files.zip успешно сгенерирован и скачан!', 'success');
                generateButton.disabled = false;
            } catch (err) {
                updateStatus(`Ошибка обработки Excel: ${err.message}`, 'error');
                generateButton.disabled = false;
            }
        };
        reader.readAsArrayBuffer(file);
    } catch (err) {
        updateStatus(`Произошла ошибка: ${err.message}`, 'error');
        generateButton.disabled = false;
    }
});

// Вспомогательная функция для обновления статуса
function updateStatus(message, type) {
    statusDiv.textContent = message;
    statusDiv.className = type;
}

// Вспомогательная функция для скачивания архива
function downloadZip(nomenklaturaCsv, postuplenieCsv) {
    return new Promise((resolve) => {
        const zip = new JSZip();
        
        // Добавляем файлы в архив с BOM для кириллицы
        zip.file("Номенклатура.csv", "\uFEFF" + nomenklaturaCsv);
        zip.file("Поступление_товаров.csv", "\uFEFF" + postuplenieCsv);
        
        // Генерируем архив и скачиваем
        zip.generateAsync({type:"blob"})
        .then(function(content) {
            const now = new Date();
            const dateStr = now.toISOString().slice(0, 19).replace(/:/g, '-').replace('T', '_');
            const link = document.createElement("a");
            link.href = URL.createObjectURL(content);
            link.download = `1c_files_${dateStr}.zip`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            resolve(); // Разрешаем Promise после скачивания
        });
    });
}

// Функция для предпросмотра Excel файла
function previewExcel(data) {
    try {
        const workbook = XLSX.read(data, {type: 'array'});
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});
        
        // Фильтруем пустые строки
        const rows = jsonData.filter(row => row.length > 0 && row[0] !== undefined);
        
        // Создаем HTML таблицу с фиксированными заголовками
        let tableHtml = `
            <h3>Предпросмотр Excel файла</h3>
            <table class="excel-preview">
                <thead>
                    <tr>
                        <th>Наименование</th>
                        <th>Количество</th>
                        <th>Цена</th>
                    </tr>
                </thead>
                <tbody>
        `;
        
        // Добавляем только первые три колонки из каждой строки
        rows.forEach(row => {
            tableHtml += '<tr>';
            // Берем первые три ячейки или пустые строки если их нет
            for (let i = 0; i < 3; i++) {
                const cell = row[i] !== undefined ? row[i] : '';
                tableHtml += `<td>${cell}</td>`;
            }
            tableHtml += '</tr>';
        });
        
        tableHtml += '</tbody></table>';
        
        // Обновляем элемент предпросмотра в модальном окне
        const previewDiv = document.getElementById('excel-preview');
        previewDiv.innerHTML = tableHtml;
        
        // Показываем модальное окно
        const modal = document.getElementById('preview-modal');
        modal.style.display = 'block';
    } catch (err) {
        updateStatus(`Ошибка предпросмотра Excel: ${err.message}`, 'error');
    }
}

// Конвертация массива объектов в CSV
function convertToCsv(items) {
    if (items.length === 0) return '';
    
    const headers = Object.keys(items[0]);
    const headerRow = headers.join(';');
    const dataRows = items.map(item => 
        headers.map(header => 
            `"${item[header]}"`
        ).join(';')
    );
    
    return [headerRow, ...dataRows].join('\n');
}

// Инициализация: активируем кнопку при загрузке страницы
window.addEventListener('DOMContentLoaded', () => {
    statusDiv.textContent = 'Готово к работе. Выберите файл и введите штрихкод.';
    generateButton.disabled = false;
    generateButton.textContent = 'Сгенерировать файлы';
    
    // Обработчик для кнопки предпросмотра
    const previewButton = document.getElementById('preview-button');
    previewButton.addEventListener('click', () => {
        const file = fileInput.files[0];
        if (!file) {
            updateStatus('Пожалуйста, выберите Excel файл.', 'error');
            return;
        }
        
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            previewExcel(data);
        };
        reader.readAsArrayBuffer(file);
    });
    
    // Закрытие модального окна
    const modal = document.getElementById('preview-modal');
    const closeBtn = document.querySelector('.close');
    closeBtn.addEventListener('click', () => {
        modal.style.display = 'none';
    });
    
    // Закрытие при клике вне окна
    window.addEventListener('click', (e) => {
        if (e.target === modal) {
            modal.style.display = 'none';
        }
    });
});