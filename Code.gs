const API_KEY = "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";
const APP_URL = "https://script.google.com/macros/s/AKfycbznznP9JqkxFtDZxeWy13Z_-EzBKM9WeXmtSGXlf9q6aUZK1q3yHoS8x2Aj4krVcgns/exec";
const DEFAULT_MODEL = "google/gemini-2.0-pro-exp-02-05:free";
const CACHE_SHEET_NAME = "__AI_CACHE__";
const LOG_SHEET_NAME = "__AI_LOGS__";

function onOpen() {
    const ui = SpreadsheetApp.getUi();

    const translateMenu = ui.createMenu('Перевести на...')
        .addItem('Русский', 'translateToRussian')
        .addItem('Английский', 'translateToEnglish')
        .addItem('Китайский (упр.)', 'translateToChinese')
        .addItem('Испанский', 'translateToSpanish')
        .addItem('Французский', 'translateToFrench')
        .addItem('Другой язык...', 'showTranslateDialog');

    //  Обновленное подменю "Объединить ячейки"
    const combineMenu = ui.createMenu('Объединить ячейки')
        .addItem('С пробелами (в одну ячейку)', 'combineCellsWithSpace')
        .addItem('С переносом строки (в одну ячейку)', 'combineCellsWithNewline')
        .addItem('Построчно (с пробелами)', 'combineCellsByRows'); // Новая функция

    ui.createMenu('AI функции')
        .addItem('Загрузить XLSX', 'showXlsxUploader')
        .addItem('Извлечь данные', 'showExtractDataDialog')
        .addItem('Создать таблицу', 'showCreateTableSidebar')
        .addSubMenu(combineMenu) // Обновленное подменю
        .addSubMenu(translateMenu)
        .addToUi();
}

//  --- ФУНКЦИИ ДЛЯ ПЕРЕВОДА ---

function showTranslateDialog() { // Одна функция showTranslateDialog!
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
        'Перевести ячейки',  //  Заголовок
        'Введите название языка (например, русский, английский, китайский):', //  Текст
        ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() === ui.Button.OK) {
        const language = response.getResponseText().trim();
        if (!language) {
            ui.alert('Ошибка', 'Вы не ввели название языка.', ui.ButtonSet.OK);
            return;
        }
        //  Вызываем translateRange с АКТИВНЫМ ДИАПАЗОНОМ
        translateRange(language, SpreadsheetApp.getActiveRange().getA1Notation()); //  Передаем адрес
    }
}

function translateRange(language, rangeStr) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    let range = sheet.getRange(rangeStr);
    const values = range.getValues();
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();

    const model = "deepseek/deepseek-r1-distill-llama-70b:free";
    const temperature = 0.1;
    const maxRetries = 3; // Максимальное количество повторных запросов

    let prompt = `Ты - бот-переводчик. Переведи каждый из следующих фрагментов текста на ${language}. `;
    prompt += `Верни ТОЛЬКО переведенные фрагменты, разделенные символом '¦' (вертикальная черта с двумя концами), без пояснений, без примеров, без дополнительных фраз. `;
    prompt += `Сохраняй порядок фрагментов. Если фрагмент пустой - верни пустую строку для него.`;

     // Формируем изначальный промпт со всеми фрагментами
    let initialPrompt = prompt; // Сохраняем базовую часть промпта
    for (let i = 0; i < numRows; i++) {
        for (let j = 0; j < numCols; j++) {
            initialPrompt += `\nФрагмент ${i * numCols + j + 1}: ${values[i][j]}`;
        }
    }
    logMessage("translateRange prompt: " + prompt);

    let translatedTexts = []; // Здесь будем хранить переведенные фрагменты
    let attempt = 0;

    while (attempt < maxRetries && translatedTexts.length === 0) { // Цикл повторных запросов
        attempt++;
        logMessage(`Попытка перевода №${attempt}`);
        try {
            const jsonResponse = openRouterRequest(initialPrompt, model, temperature); // Используем полный промпт

            if (jsonResponse && jsonResponse.choices && jsonResponse.choices.length > 0 && jsonResponse.choices[0].message) {
                let answer = jsonResponse.choices[0].message.content;
                answer = answer.trim();
                answer = answer.replace(/^¦+/, '');
                answer = answer.replace(/¦+$/, '');
                logMessage("translateRange answer: " + answer);

                // Разбиваем ответ на части (по '¦')
                translatedTexts = answer.split('¦').map(t => t.trim()).filter(t => t !== "");

                // --- Проверка на совпадение с оригиналом ---
                let allMatch = true;
                let translatedIndex = 0;

                for (let i = 0; i < numRows; i++) {
                    for (let j = 0; j < numCols; j++) {
                         if (values[i][j] !== "") { // Проверяем исходные
                            if (translatedIndex < translatedTexts.length) { //Проверяем наличие
                                if (values[i][j].trim().toLowerCase() !== translatedTexts[translatedIndex].toLowerCase()) {
                                     allMatch = false; // Если хоть один фрагмент не совпадает, продолжаем
                                }
                                translatedIndex++;
                            } else { // Если кончился массив с переводами
                                allMatch = false;
                            }
                         }

                    }
                }


                if (allMatch) {
                    logMessage(`Перевод совпадает с оригиналом (попытка ${attempt}). Повторяем запрос...`);
                    translatedTexts = []; // Сбрасываем результат, чтобы повторить запрос
                } else {
                    // Если перевод не совпадает с оригиналом, выходим из цикла
                    break;
                }
            } else {
                const errorMessage = "Ошибка: Неожиданный ответ от OpenRouter при переводе.";
                logMessage(errorMessage, true);
                throw new Error(errorMessage); //  Выходим, если совсем плохой ответ
            }
        } catch (error) {
            logMessage(`Ошибка перевода (попытка ${attempt}): ${error.toString()}`, true);
             if (attempt === maxRetries) {
                throw error; // Если все попытки исчерпаны, пробрасываем ошибку
            }
        }
    }

    if (translatedTexts.length === 0) {
        throw new Error("Не удалось получить перевод после нескольких попыток.");
    }

    // Вставляем переводы обратно в ячейки
       let textIndex = 0;
        for (let i = 0; i < numRows; i++) {
            for (let j = 0; j < numCols; j++) {
                if (values[i][j] !== "") { // Вставляем, только если *исходная* ячейка не пустая
                    if (textIndex < translatedTexts.length) {
                        range.getCell(i + 1, j + 1).setValue(translatedTexts[textIndex]);
                        textIndex++;
                    }
                }
            }
        }

    return "Перевод выполнен";
}

// Функции для вызова из меню (с разными языками)
function translateToRussian() {
    translateRange('русский', SpreadsheetApp.getActiveRange().getA1Notation()); // Передаем адрес
}

function translateToEnglish() {
    translateRange('английский', SpreadsheetApp.getActiveRange().getA1Notation());
}

function translateToChinese() {
    translateRange('китайский', SpreadsheetApp.getActiveRange().getA1Notation());
}

function translateToSpanish() {
    translateRange('испанский', SpreadsheetApp.getActiveRange().getA1Notation());
}

function translateToFrench() {
     translateRange('французский', SpreadsheetApp.getActiveRange().getA1Notation());
}


// --- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---

function getSelectedRangeA1Notation() { //Для получения данных с клиента
  return SpreadsheetApp.getActiveRange().getA1Notation();
}

// ---  ОСТАЛЬНЫЕ ФУНКЦИИ (без изменений) ---

function combineCellsByRows() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    const selectedRange = sheet.getActiveRange();

    if (!selectedRange) {
        SpreadsheetApp.getUi().alert('Ошибка', 'Выделите диапазон ячеек для объединения.', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    }

    const values = selectedRange.getValues();
    const numRows = selectedRange.getNumRows();
    const numCols = selectedRange.getNumColumns();

    //  Итерируем по *строкам*
    for (let i = 0; i < numRows; i++) {
        let combinedString = "";
        //  Объединяем значения *внутри текущей строки*
        for (let j = 0; j < numCols; j++) {
            const cellValue = values[i][j];
            if (cellValue !== "") {
                combinedString += cellValue + " ";
            }
        }
        combinedString = combinedString.trim(); // Убираем лишний пробел в конце

        //  Вставляем объединенную строку в *первую* ячейку текущей строки
        selectedRange.getCell(i + 1, 1).setValue(combinedString);

        //  Очищаем остальные ячейки в текущей строке
        for (let j = 1; j < numCols; j++) { //  Начинаем с j = 1 (второй столбец)
            selectedRange.getCell(i + 1, j + 1).clearContent();
        }
    }
    SpreadsheetApp.getUi().alert("Ячейки объединены",`Результат в строках  ${selectedRange.getA1Notation()}`, SpreadsheetApp.getUi().ButtonSet.OK);
}

function combineCellsWithNewline() {
  // ... (код функции combineCellsWithNewline - без изменений) ...
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    const selectedRange = sheet.getActiveRange();

    if (!selectedRange) {
        SpreadsheetApp.getUi().alert('Ошибка', 'Выделите диапазон ячеек для объединения.', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    }

    const values = selectedRange.getValues();
    let combinedString = "";

    for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values[i].length; j++) {
            const cellValue = values[i][j];
            if (cellValue !== "") {
                combinedString += cellValue + "\n"; // Добавляем перевод строки!
            }
        }
    }

    combinedString = combinedString.trim(); // Убираем лишний перевод строки в конце

    const firstCell = selectedRange.getCell(1, 1);
    firstCell.setValue(combinedString);
    firstCell.setWrap(true); // !!! Добавляем перенос текста в ячейке

    const numRows = selectedRange.getNumRows();
    const numCols = selectedRange.getNumColumns();
    if (numRows > 1 || numCols > 1) {
        for (let i = 0; i < numRows; i++) {
            for (let j = 0; j < numCols; j++) {
                if (i !== 0 || j !== 0) {
                    selectedRange.getCell(i + 1, j + 1).clearContent();
                }
            }
        }
    }
     SpreadsheetApp.getUi().alert('Ячейки объединены',`Результат в ячейке ${firstCell.getA1Notation()}`, SpreadsheetApp.getUi().ButtonSet.OK);

}

function combineCellsWithSpace() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    const selectedRange = sheet.getActiveRange();

    if (!selectedRange) {
        SpreadsheetApp.getUi().alert('Ошибка', 'Выделите диапазон ячеек для объединения.', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    }

    const values = selectedRange.getValues();
    let combinedString = "";

    for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values[i].length; j++) {
            const cellValue = values[i][j];
            if (cellValue !== "") {
                combinedString += cellValue + " ";
            }
        }
    }

    combinedString = combinedString.trim();

    const firstCell = selectedRange.getCell(1, 1);
    firstCell.setValue(combinedString);

    const numRows = selectedRange.getNumRows();
    const numCols = selectedRange.getNumColumns();
    if (numRows > 1 || numCols > 1) {
        for (let i = 0; i < numRows; i++) {
            for (let j = 0; j < numCols; j++) {
                if (i !== 0 || j !== 0) {
                    selectedRange.getCell(i + 1, j + 1).clearContent();
                }
            }
        }
    }
     SpreadsheetApp.getUi().alert('Ячейки объединены',`Результат в ячейке ${firstCell.getA1Notation()}`, SpreadsheetApp.getUi().ButtonSet.OK);

}

function showXlsxUploader() {
     const html = HtmlService.createHtmlOutputFromFile('XlsxUploader')
        .setWidth(300)
        .setHeight(400);
    SpreadsheetApp.getUi().showSidebar(html);
}

function showExtractDataDialog() {
    const html = HtmlService.createHtmlOutputFromFile('ExtractDataDialog')
        .setWidth(400)
        .setHeight(300);
    SpreadsheetApp.getUi().showSidebar(html);
}

function showCreateTableSidebar() {
   const html = HtmlService.createHtmlOutputFromFile('CreateTableSidebar')
        .setWidth(400)
        .setHeight(300);
    SpreadsheetApp.getUi().showSidebar(html);
}

function getOrCreateCacheSheet() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let cacheSheet = spreadsheet.getSheetByName("__COMBINE_CACHE__");
    if (!cacheSheet) {
        cacheSheet = spreadsheet.insertSheet("__COMBINE_CACHE__");
        cacheSheet.hideSheet();
    }
    return cacheSheet;
}

function createTable(query) {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    const activeCell = sheet.getActiveCell();
    if (!activeCell) {
        throw new Error("Выберите ячейку, с которой начать создание таблицы.");
    }
    const brokenBar = '¦';

    const prompt = `
    Создай таблицу на основе следующего запроса:
    ${query}

    Верни ТОЛЬКО таблицу в формате CSV.
    Используй '${brokenBar}' (вертикальная черта с двумя концами) как разделитель столбцов.
    Используй ';' (точку с запятой) как разделитель строк.
    НЕ включай в ответ никаких пояснений, текста и прочего – ТОЛЬКО CSV данные.
    `;

    logMessage("createTable prompt: " + prompt);
    const model = DEFAULT_MODEL;
    const temperature = 0.4;

    try {
        const jsonResponse = openRouterRequest(prompt, model, temperature);

        if (jsonResponse && jsonResponse.choices && jsonResponse.choices.length > 0 && jsonResponse.choices[0].message) {
            let answer = jsonResponse.choices[0].message.content;
            answer = answer.trim();
            answer = answer.replace(/;\s*$/, "");
            answer = answer.replace(/^```csv\s*/i, '');
            answer = answer.replace(/```\s*$/i, '');
            answer = answer.trim();
            logMessage("createTable answer: " + answer);

            let rows = answer.split('\n').map(row => row.trim()).filter(row => row !== ""); // Изменили на \n
            if (rows.length === 0) {
                throw new Error("AI не вернул таблицу.");
            }

            let parsedData = rows.map(row => Utilities.parseCsv(row, brokenBar)[0].map(cell => cell.trim()));


            // --- Унифицируем количество столбцов ---
            const numCols = parsedData[0].length; // Количество столбцов в *первой* строке (заголовки)

            for (let i = 0; i < parsedData.length; i++) {
                if (parsedData[i].length > numCols) {
                    // *Удаляем* лишние столбцы с конца
                    parsedData[i] = parsedData[i].slice(0, numCols);
                } else if (parsedData[i].length < numCols){
                    // *Добавляем* пустые строки
                    while (parsedData[i].length < numCols) {
                        parsedData[i].push("");
                    }
                }
            }


            const numRows = parsedData.length;

            //Вставляем
            const range = activeCell.offset(0, 0, numRows, numCols);
            range.setValues(parsedData);

            range.setBorder(true, true, true, true, true, true);
            range.setVerticalAlignment('middle');
            if (numRows > 0) {
                sheet.getRange(activeCell.getRow(), activeCell.getColumn(), 1, numCols).setFontWeight('bold');
            }

            for (let i = 1; i <= numCols; i++) {
                sheet.autoResizeColumn(activeCell.getColumn() + i - 1);
            }

            return "Таблица создана.";

        } else {
            const errorMessage = "Ошибка: Неожиданный ответ от OpenRouter";
            logMessage(errorMessage, true);
            throw new Error(errorMessage);
        }

    } catch (error) {
        logMessage("Ошибка в createTable: " + error.toString(), true);
        throw error;
    }
}

function extractData(sourceRangeStr, insertRangeStr, headersStr) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    const brokenBar = '¦';

    let sourceRange, insertRange;
    try {
        sourceRange = sheet.getRange(sourceRangeStr);
        insertRange = sheet.getRange(insertRangeStr);
    } catch (e) {
        throw new Error("Неверный формат диапазона: " + e.message);
    }

    const numRows = sourceRange.getNumRows();
    const numCols = sourceRange.getNumColumns();
    const insertNumRows = insertRange.getNumRows();
    const insertNumCols = insertRange.getNumColumns();

    const sourceValues = sourceRange.getValues();
    const headers = headersStr.split(',').map(h => h.trim());

    let prompt = `Извлеки из следующего текста указанные данные. Верни значения, разделенные символом '${brokenBar}' (вертикальная черта с двумя концами). Не добавляй никаких заголовков, пояснений, лишних символов и переводов строк.
\nДанные:\n`;

    for (let i = 0; i < numRows; i++) {
        for (let j = 0; j < numCols; j++) {
            prompt += `Ячейка ${i + 1},${j + 1}: ${sourceValues[i][j]}\n`;
        }
    }

    prompt += `\nИзвлеки ТОЛЬКО следующие данные (в указанном порядке):\n`;
    prompt += headers.join(brokenBar) + '\n';
    prompt +=`Если каких-то данных нет, всё равно выводи разделитель, чтобы сохранить порядок.`;
    prompt += `\nЕсли в тексте встречается несколько значений для одного заголовка, верни их все, разделив ПЕРЕВОДОМ СТРОКИ.`;

    logMessage("extractData prompt: " + prompt);
    const model = DEFAULT_MODEL;
    const temperature = 0.1;

    try {
        const jsonResponse = openRouterRequest(prompt, model, temperature);

        if (jsonResponse && jsonResponse.choices && jsonResponse.choices.length > 0 && jsonResponse.choices[0].message) {
            let answer = jsonResponse.choices[0].message.content;
            answer = answer.trim();
            logMessage("extractData answer: " + answer);

            let parsedData = answer.split(brokenBar).map(cell => cell.trim());
            logMessage("parsedData: " + parsedData)

            let dataIndex = 0;
            for (let i = 0; i < insertNumRows; i++) {
                for (let j = 0; j < insertNumCols; j++) {
                    if (dataIndex < parsedData.length) {
                        const cellValue = parsedData[dataIndex].replace(/\\n/g, '\n');
                        insertRange.getCell(i + 1, j + 1).setValue(cellValue);
                        dataIndex++;
                    } else {
                        insertRange.getCell(i + 1, j + 1).setValue("");
                    }
                }
            }

            return "Данные извлечены.";

        } else {
            const errorMessage = "Ошибка: Неожиданный ответ от OpenRouter";
            logMessage(errorMessage, true);
            throw new Error(errorMessage);
        }

    } catch (error) {
        logMessage("Ошибка в extractData: " + error.toString(), true);
        throw error;
    }
}

function openRouterRequest(prompt, model, temperature, customUrl = null, retries = 3) {
    const url = customUrl ? customUrl : 'https://openrouter.ai/api/v1/chat/completions';
    const headers = {
        'Authorization': `Bearer ${API_KEY}`,
        'Content-Type': 'application/json',
        'HTTP-Referer': APP_URL,
    };
    const data = {
        'model': model,
        'messages': [{ 'role': 'user', 'content': prompt }],
        'temperature': temperature,
    };
    const options = {
        'method': 'post',
        'headers': headers,
        'payload': JSON.stringify(data),
        'muteHttpExceptions': true
    };

    let lastError = null;

    for (let i = 0; i < retries; i++) {
        try {
            const response = UrlFetchApp.fetch(url, options);
            const responseCode = response.getResponseCode();
            const responseText = response.getContentText();

            logMessage(`Попытка ${i + 1}: Response code: ${responseCode}`);
            logMessage(`Попытка ${i + 1}: Response text: ${responseText}`);

            if (responseCode === 200) {
                const jsonResponse = JSON.parse(responseText);
                if (jsonResponse && jsonResponse.choices && jsonResponse.choices.length > 0 && jsonResponse.choices[0].message) {
                  return jsonResponse;
                } else {
                   lastError = new Error("Пустой ответ от OpenRouter");
                   logMessage(lastError.message, true);
                }

            } else {
                let errorMessage = `Ошибка OpenRouter: ${responseCode} - ${responseText}`;
                if (responseCode === 400) {
                  errorMessage += " (Bad Request - возможно, неверный формат запроса)";
                } else if (responseCode === 401) {
                    errorMessage += " (Unauthorized - проверьте API ключ)";
                } else if (responseCode === 429) {
                    errorMessage += " (Too Many Requests - превышен лимит запросов)";
                     Utilities.sleep(5000);

                } else if (responseCode === 500) {
                    errorMessage += " (Internal Server Error - ошибка на стороне OpenRouter)";
                }

                lastError = new Error(errorMessage);
                logMessage(errorMessage, true);
                 if (responseCode !== 429) {
                    Utilities.sleep(1000);
                 }
            }
        } catch (error) {
            lastError = error;
            logMessage(`Попытка ${i + 1}: Ошибка: ${error}`, true);
             Utilities.sleep(1000);

        }
    }
    throw lastError;
}

function processXlsxFile(fileData) {
    try {
        const blob = Utilities.newBlob(Utilities.base64Decode(fileData.data), fileData.type, fileData.name);
        const folder = DriveApp.getRootFolder();
        const tempXlsxFile = folder.createFile(blob);
        const xlsxFileId = tempXlsxFile.getId();
        logMessage("Временный XLSX файл создан: " + xlsxFileId);

        const tempSheet = Drive.Files.copy({
            mimeType: MimeType.GOOGLE_SHEETS,
            parents: [{ id: folder.getId() }],
            title: 'Temp Sheet ' + Date.now()
        }, xlsxFileId);
        const tempSheetId = tempSheet.getId();
        logMessage("Временная Google Таблица создана: " + tempSheetId);

        return { tempSheetId, xlsxFileId };

    } catch (error) {
        logMessage("Ошибка при обработке XLSX: " + error + "\nStack: " + error.stack, true);
        return { error: "Ошибка при обработке файла: " + error };
    }
}

function analyzeAndInsertData(sheetId, xlsxFileId) {
    if (!sheetId) {
        logMessage("Ошибка: sheetId не передан в analyzeAndInsertData.", true);
        return "Ошибка: ID таблицы не передан.";
    }
    logMessage("Начало analyzeAndInsertData. sheetId: " + sheetId);
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = spreadsheet.getActiveSheet();
    try {
        const tempSpreadsheet = SpreadsheetApp.openById(sheetId);
        const sourceSheet = tempSpreadsheet.getSheets()[0];
        const sheetData = sourceSheet.getDataRange().getValues();
        logMessage("sheetData (исходные): " + JSON.stringify(sheetData));

        if (!sheetData || sheetData.length === 0) {
            const errorMessage = "Ошибка: Не удалось прочитать данные из таблицы, или она пустая.";
            logMessage(errorMessage, true);
            deleteTempFiles(xlsxFileId, sheetId);
            logMessage("Временные файлы удалены (ошибка - нет данных).");
            return errorMessage;
        }
        const xlsxHeaders = sheetData[0] ? sheetData[0].map(String) : [];
        if (xlsxHeaders.length === 0) {
            const errorMessage = "Ошибка: Не удалось прочитать заголовки из XLSX.";
            logMessage(errorMessage, true);
            deleteTempFiles(xlsxFileId, sheetId);
            return errorMessage;
        }
        const dataRows = sheetData.slice(1);
        const dataString = dataRows.map(row => row.join('|')).join('\n');

        const prompt = `
        Имеется таблица с данными о товарах. Вот заголовки столбцов:
        ${xlsxHeaders.join(',')}
        Вот данные (в формате CSV, значения разделены ВЕРТИКАЛЬНОЙ ЧЕРТОЙ (|)):
        ${dataString}
        Твоя задача: извлечь из этих данных ТОЛЬКО следующую информацию и вернуть ее в СТРОГОМ формате CSV, БЕЗ ЗАГОЛОВКОВ и БЕЗ ЛИШНИХ СИМВОЛОВ, используя ВЕРТИКАЛЬНУЮ ЧЕРТУ (|) как разделитель *значений* и ТОЧКУ С ЗАПЯТОЙ (;) как разделитель *строк*.

        Требуемые данные (СТРОГО в этом порядке):
        1. Номер заказа: извлеки ТОЛЬКО цифры и символ подчеркивания (если есть). Удали все остальное (буквы, слова "Сделка №" и т.д.).
        2. Название товара: извлеки ПОЛНОЕ название товара и ПЕРЕВЕДИ ЕГО НА АНГЛИЙСКИЙ ЯЗЫК. Не обрезай его!
        3. Бренд: извлеки название бренда.
        4. Количество: извлеки количество товара.

        Правила обработки:
        * ЕСЛИ АРТИКУЛА В НАЗВАНИИ ТОВАРА НЕТ, НО ОН ЕСТЬ в отдельном столбце "Код товара", то ДОБАВЬ ЕГО К НАЗВАНИЮ ТОВАРА ЧЕРЕЗ ПРОБЕЛ В КОНЕЦ (уже после перевода на английский).
        * Если каких-то данных нет (например, не указан бренд), возвращай для этого значения ПУСТУЮ СТРОКУ ("").
        * Никаких других данных, кроме перечисленных выше 4-х, возвращать НЕ НУЖНО.
        * Строго соблюдай формат CSV: значения разделены ВЕРТИКАЛЬНОЙ ЧЕРТОЙ (|), каждая запись разделена ТОЧКОЙ С ЗАПЯТОЙ (;).
        * НЕ добавляй заголовки столбцов в результирующий CSV.
        * НЕ добавляй никаких лишних символов (кавычек, обрамляющих CSV-строку, и т.п.). ТОЛЬКО данные.
        * ВАЖНО: НЕ используй символ перевода строки (\n) внутри значений CSV.  Используй ТОЧКУ С ЗАПЯТОЙ (;) для разделения строк.

        Пример ПРАВИЛЬНОГО формата ответа (с переводом названия на английский):
        12345|Product Name 1|Brand 1|10;67890|Product Name 2||5;|Product Name 3|Brand 3|;
        Верни ТОЛЬКО CSV-данные, без пояснений.
    `;
        logMessage("prompt: " + prompt);
        const model = DEFAULT_MODEL;
        const temperature = 0.1;
        const startTime = new Date().getTime();

        const jsonResponse = openRouterRequest(prompt, model, temperature);

        if (jsonResponse && jsonResponse.choices && jsonResponse.choices.length > 0 && jsonResponse.choices[0].message) {
            let answer = jsonResponse.choices[0].message.content;
            answer = answer.trim();
            answer = answer.replace(/^```csv\s*/i, '');
            answer = answer.replace(/```\s*$/i, '');
            answer = answer.trim();


            logMessage("answer (после очистки): " + answer);
            let aiData = Utilities.parseCsv(answer, '|');

            if (!aiData || aiData.length === 0) {
                const errorMessage = "Ошибка: AI не вернул данные в формате CSV.";
                logMessage(errorMessage, true);
                deleteTempFiles(xlsxFileId, sheetId);
                return errorMessage;
            }
            aiData = aiData.filter(row => Array.isArray(row) && row.some(cell => cell.trim() !== ""));

            let finalAiData = [];
            for (const row of aiData) {
                let temp = row.join('|').split(';');
                temp.forEach(r => {
                    if (r) {
                        let parsed = Utilities.parseCsv(r, '|')[0];
                        if (parsed && parsed.length > 0) {
                            finalAiData.push(parsed);
                        }
                    }
                });

            }

            logMessage("aiData (после фильтрации): " + JSON.stringify(finalAiData));
            const processedData = [];
            let firstRowOfOrder = 0;
            let orderNumberForAI = "";

            for (let i = 0; i < finalAiData.length; i++) {
                const row = finalAiData[i];
                if (!Array.isArray(row) || row.length < 4) {
                    logMessage("Предупреждение: строка aiData имеет неверный формат. Пропускаем.", false);
                    continue;
                }
                let orderNumber = row[0] || "";
                let productName = row[1] || "";
                let brand = row[2] || "";
                let quantity = row[3] || "";
                if (!orderNumber) {
                    logMessage("Пропускаем строку - пустой номер заказа.", false);
                    continue;
                }
                const originalRow = sheetData[i + 1];
                if (!originalRow) {
                    logMessage("Не найдена строка: " + i, false);
                    continue;
                }
                const orderNumberIndex = xlsxHeaders.findIndex(header => header.toLowerCase().includes("сделка"));
                const productNameIndex = xlsxHeaders.findIndex(header => header.toLowerCase().includes("товар"));
                const partNumberIndex = xlsxHeaders.findIndex(header => header.toLowerCase().includes("код товара"));
                const brandIndex = xlsxHeaders.findIndex(header => header.toLowerCase().includes("бренд"));
                const quantityIndex = xlsxHeaders.findIndex(header => header.toLowerCase().includes("кол-во"));
                let originalProductName = originalRow[productNameIndex] ? String(originalRow[productNameIndex]) : '';
                let originalPartNumber = originalRow[partNumberIndex] ? String(originalRow[partNumberIndex]) : '';
                const normalizedPartNumber = normalizePartNumber(originalPartNumber);
                const normalizedProductName = normalizePartNumber(originalProductName);
                if (originalPartNumber && originalPartNumber !== '-' && originalPartNumber.toLowerCase() !== 'null' && originalPartNumber.trim() !== '') {
                    if (!normalizedProductName.includes(normalizedPartNumber)) {
                        productName = productName + " " + originalPartNumber;
                    }
                }
                if (firstRowOfOrder === 0) {
                    firstRowOfOrder = i + 1;
                    orderNumberForAI = orderNumber;
                }
                processedData.push([orderNumber, productName, brand, quantity]);
            }
            logMessage("processedData (после ручной обработки): " + JSON.stringify(processedData));
            const lastRow = targetSheet.getLastRow();
            const lastColumn = targetSheet.getLastColumn();
            const headerRowValues = targetSheet.getRange(2, 1, 1, lastColumn).getValues();
            let headerRow = [];
            if (headerRowValues.length > 0) {
                headerRow = headerRowValues[0].filter(cell => cell !== "");
            }
            logMessage("headerRow: " + JSON.stringify(headerRow));
            if (headerRow.length === 0) {
                const errorMessage = "Ошибка: Вторая строка целевого листа пустая. Невозможно определить заголовки столбцов.";
                logMessage(errorMessage, true);
                deleteTempFiles(xlsxFileId, sheetId);
                return errorMessage;
            }
            const headersMap = {};
            for (let i = 0; i < headerRow.length; i++) {
                headersMap[String(headerRow[i]).toLowerCase()] = i + 1;
            }
            logMessage("headersMap: " + JSON.stringify(headersMap));
            const insertCols = [];
            const requiredHeaders = ["номер задачи", "запрос", "бренд", "кол / qty"];
            for (const requiredHeader of requiredHeaders) {
                let found = false;
                for (const header in headersMap) {
                    if (header.toLowerCase().includes(requiredHeader.toLowerCase())) {
                        insertCols.push(headersMap[header]);
                        found = true;
                        break;
                    }
                }
                if (!found) {
                    insertCols.push(null);
                }
            }
            logMessage("insertCols: " + JSON.stringify(insertCols));
            if (processedData.length > 0) {
                const startRowInsert = lastRow + 1;
                let currentOrderNumber = null;
                let orderStartRow = null;
                for (let i = 0; i < processedData.length; i++) {
                    const row = processedData[i];
                    const orderNumber = row[0];
                    const productName = row[1];
                    const brand = row[2];
                    const quantity = row[3];
                    if (orderNumber !== currentOrderNumber) {
                        currentOrderNumber = orderNumber;
                        orderStartRow = startRowInsert + i;
                    }
                    if (insertCols[0] && orderNumber) {
                        targetSheet.getRange(orderStartRow, insertCols[0]).setValue(orderNumber);
                    }
                    if (insertCols[1] && productName) {
                        targetSheet.getRange(startRowInsert + i, insertCols[1]).setValue(productName);
                    }
                    if (insertCols[2] && brand) {
                        targetSheet.getRange(startRowInsert + i, insertCols[2]).setValue(brand);
                    }
                    if (insertCols[3] && quantity) {
                        targetSheet.getRange(startRowInsert + i, insertCols[3]).setValue(quantity);
                    }
                }
                if (insertCols[0] && processedData.length > 1) {
                    targetSheet.getRange(orderStartRow, insertCols[0], processedData.length, 1).merge();
                }
            }
            logMessage("Данные успешно извлечены и записаны.");

            const endTime = new Date().getTime();
            const elapsedTime = endTime - startTime;
            const delay = Math.max(0, 3000 - elapsedTime);

            //Utilities.sleep(delay);  // Закомментировано

            return "Данные успешно извлечены и записаны.";
        } else {
            const errorMessage = "Ошибка: Неожиданный ответ от OpenRouter";
            logMessage(errorMessage, true);
            deleteTempFiles(xlsxFileId, sheetId);
            return errorMessage;
        }
    } catch (error) {
        logMessage("Ошибка в analyzeAndInsertData: " + error.toString() + " " + error.stack, true);
        deleteTempFiles(xlsxFileId, sheetId);
        return "Ошибка: " + error.toString();
    } finally {
        deleteTempFiles(xlsxFileId, sheetId);
    }

}

function getHeadersMap(sheet) {
    const headersRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    const map = {};
    headersRow.forEach((header, index) => {
        const key = header.toLowerCase();
        if (key.includes('номер задачи')) map['номер задачи'] = sheet.getRange(2, index + 1);
        if (key.includes('запрос')) map['запрос'] = sheet.getRange(2, index + 1);
        if (key.includes('бренд')) map['бренд'] = sheet.getRange(2, index + 1);
        if (key.includes('кол / qty')) map['кол / qty'] = sheet.getRange(2, index + 1);
    });
    return map;
}

function deleteTempFiles(xlsxFileId, sheetId) {
    try { DriveApp.getFileById(xlsxFileId).setTrashed(true); } catch (e) {}
    try { DriveApp.getFileById(sheetId).setTrashed(true); } catch (e) {}
}

function getLogSheet() {
    let logSheet = SpreadsheetApp.getActive().getSheetByName(LOG_SHEET_NAME);
    if (!logSheet) {
        logSheet = SpreadsheetApp.getActive().insertSheet(LOG_SHEET_NAME);
        logSheet.hideSheet();
        logSheet.appendRow(["Время", "Сообщение", "Ошибка"]);
    }
    return logSheet;
}

function logMessage(message, isError = false) {
    getLogSheet().appendRow([new Date(), message, isError]);
}

function normalizePartNumber(partNumber) {
    if (!partNumber) return "";
    return String(partNumber).replace(/[^a-zA-Z0-9]/g, "").toUpperCase();
}