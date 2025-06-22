const fileInput = document.getElementById('file-input');
const barcodeInput = document.getElementById('barcode-input');
const generateButton = document.getElementById('generate-button');
const statusDiv = document.getElementById('status');

// Основная асинхронная функция для инициализации
async function main() {
    // Загружаем Pyodide и необходимые пакеты
    let pyodide = await loadPyodide();
    statusDiv.textContent = 'Загрузка библиотеки pandas...';
    await pyodide.loadPackage("pandas");
    
    // После загрузки кнопка становится активной
    statusDiv.textContent = 'Готово к работе. Выберите файл и введите штрихкод.';
    generateButton.disabled = false;
    generateButton.textContent = 'Сгенерировать файлы';

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
            const fileBuffer = await file.arrayBuffer();
            
            // Передаем данные файла и штрихкода в среду Python
            pyodide.globals.set("file_content", fileBuffer);
            pyodide.globals.set("start_barcode_str", startBarcode);

            // Тот самый Python-код из вашего приложения, адаптированный для Pyodide
            let pythonCode = `
import pandas as pd
import io
from datetime import datetime

def process_files():
    try:
        # Преобразуем буфер в DataFrame
        df_input = pd.read_excel(io.BytesIO(file_content.to_py()), header=None)
        
        if df_input.shape[1] < 3:
            return {"error": "Входной Excel файл должен содержать как минимум 3 столбца."}

        df_input.columns = ['col_0', 'col_1', 'col_2'] + [f'col_{i}' for i in range(3, df_input.shape[1])]

        current_barcode_int = int(start_barcode_str)
        barcode_length = len(start_barcode_str)

        # --- Подготовка данных для "Номенклатура" ---
        nomenklatura_data = []
        for index, row in df_input.iterrows():
            next_barcode_numeric_value = current_barcode_int + index
            barcode = str(next_barcode_numeric_value).zfill(barcode_length)
            nomenklatura_data.append({
                "Штрихкод": barcode,
                "Наименование": row['col_0'],
                "ЭтоУслуга": "Нет",
                "ЕдиницаИзмерения": "шт",
                "Ставка НДС": "Без НДС",
                "Цена": row['col_2']
            })
        df_nomenklatura = pd.DataFrame(nomenklatura_data)

        # --- Подготовка данных для "Поступление_товаров" ---
        postuplenie_data = []
        for index, row in df_input.iterrows():
            postuplenie_data.append({
                "Номенклатура": row['col_0'],
                "Количество": row['col_1'],
                "Единица измерения": "шт",
                "Цена": row['col_2'],
                "ЭтоУслуга": "Нет"
            })
        df_postuplenie = pd.DataFrame(postuplenie_data)

        # Преобразование типов
        for col_name in ['Количество', 'Цена']:
            if col_name in df_postuplenie.columns:
                df_postuplenie[col_name] = pd.to_numeric(df_postuplenie[col_name], errors='coerce').fillna(0).astype(int)

        # Преобразуем DataFrame в CSV-строки
        nomenklatura_csv = df_nomenklatura.to_csv(index=False, encoding='utf-8-sig', sep=';')
        postuplenie_csv = df_postuplenie.to_csv(index=False, encoding='utf-8-sig', sep=';')
        
        return {
            "nomenklatura_csv": nomenklatura_csv, 
            "postuplenie_csv": postuplenie_csv
        }

    except Exception as e:
        return {"error": f"Произошла ошибка в Python: {str(e)}"}

# Запускаем функцию и возвращаем результат
process_files()
            `;
            
            // Запускаем Python-код
            const results = await pyodide.runPythonAsync(pythonCode);
            const resultsJs = results.toJs();

            if (resultsJs.has('error')) {
                updateStatus(resultsJs.get('error'), 'error');
            } else {
                // Если все успешно, создаем ссылки для скачивания файлов
                triggerDownload(resultsJs.get('nomenklatura_csv'), 'Номенклатура.csv');
                triggerDownload(resultsJs.get('postuplenie_csv'), 'Поступление_товаров.csv');
                updateStatus('Файлы успешно сгенерированы и скачаны!', 'success');
            }

        } catch (err) {
            updateStatus(`Произошла ошибка: ${err.message}`, 'error');
        } finally {
            generateButton.disabled = false;
        }
    });
}

// Вспомогательная функция для обновления статуса
function updateStatus(message, type) {
    statusDiv.textContent = message;
    statusDiv.className = type; // 'success' или 'error'
}

// Вспомогательная функция для скачивания файла
function triggerDownload(content, filename) {
    const blob = new Blob([content], { type: 'text/csv;charset=utf-8-sig;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// Запускаем все
main();