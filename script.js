const fileInput = document.getElementById('file-input');
const barcodeInput = document.getElementById('barcode-input');
const nomenklaturaButton = document.getElementById('nomenklatura-button');
const postuplenieButton = document.getElementById('postuplenie-button');
const statusDiv = document.getElementById('status');

// Обработчик для кнопки "Номенклатура"
nomenklaturaButton.addEventListener('click', async () => {
    generateFile('nomenklatura');
});

// Обработчик для кнопки "Поступление"
postuplenieButton.addEventListener('click', async () => {
    generateFile('postuplenie');
});

async function generateFile(fileType) {
    const file = fileInput.files[0];
    const startBarcode = barcodeInput.value.trim();

    if (!file) {
        updateStatus('Пожалуйста, выберите Excel файл.', 'error');
        return;
    }
    if (!startBarcode && fileType === 'nomenklatura') {
        updateStatus('Пожалуйста, введите начальный штрихкод.', 'error');
        return;
    }

    updateStatus('Обработка файла...', 'info');
    nomenklaturaButton.disabled = true;
    postuplenieButton.disabled = true;

    try {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});
                
                const rows = jsonData.filter(row => row.length > 0 && row[0] !== undefined);
                
                if (fileType === 'nomenklatura') {
                    let currentBarcodeInt = parseInt(startBarcode);
                    const barcodeLength = startBarcode.length;
                    
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

                    const nomenklaturaWb = XLSX.utils.book_new();
                    const nomenklaturaWs = XLSX.utils.json_to_sheet(nomenklaturaData);
                    XLSX.utils.book_append_sheet(nomenklaturaWb, nomenklaturaWs, "Номенклатура");
                    XLSX.writeFile(nomenklaturaWb, "Номенклатура.xlsx");
                    updateStatus('Файл Номенклатура.xlsx успешно сгенерирован и скачан!', 'success');
                } else if (fileType === 'postuplenie') {
                    const postuplenieData = rows.map(row => ({
                        "Номенклатура": row[0],
                        "Количество": row[1],
                        "Единица измерения": "шт",
                        "Цена": row[2],
                        "ЭтоУслуга": "Нет"
                    }));

                    const postuplenieWb = XLSX.utils.book_new();
                    const postuplenieWs = XLSX.utils.json_to_sheet(postuplenieData);
                    XLSX.utils.book_append_sheet(postuplenieWb, postuplenieWs, "Поступление_товаров");
                    XLSX.writeFile(postuplenieWb, "Поступление_товаров.xlsx");
                    updateStatus('Файл Поступление_товаров.xlsx успешно сгенерирован и скачан!', 'success');
                }

                nomenklaturaButton.disabled = false;
                postuplenieButton.disabled = false;
            } catch (err) {
                updateStatus(`Ошибка обработки Excel: ${err.message}`, 'error');
                nomenklaturaButton.disabled = false;
                postuplenieButton.disabled = false;
            }
        };
        reader.readAsArrayBuffer(file);
    } catch (err) {
        updateStatus(`Произошла ошибка: ${err.message}`, 'error');
        nomenklaturaButton.disabled = false;
        postuplenieButton.disabled = false;
    }
}

// Вспомогательная функция для обновления статуса
function updateStatus(message, type) {
    statusDiv.textContent = message;
    statusDiv.className = type;
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



// Инициализация: активируем кнопку при загрузке страницы
window.addEventListener('DOMContentLoaded', () => {
    statusDiv.textContent = 'Готово к работе. Выберите файл и введите штрихкод.';
    nomenklaturaButton.disabled = false;
    postuplenieButton.disabled = false;
    
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