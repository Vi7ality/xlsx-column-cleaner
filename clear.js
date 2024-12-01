const fs = require("fs");
const xlsx = require("xlsx");

// Функция для удаления определённого HTML-кода из строки
function cleanHtmlBlock(text) {
  if (typeof text === "string") {
    // Регулярное выражение для удаления нужного HTML-кода
    const pattern = /<div class="text-center">.*?<\/div>/gs;
    return text.replace(pattern, "");
  }
  return text;
}

// Загрузка Excel-файла
const filePath = "./part2.xlsx"; // Укажите путь к вашему файлу
const workbook = xlsx.readFile(filePath);

// Выбор первого листа
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Конвертируем лист в JSON
const data = xlsx.utils.sheet_to_json(worksheet);

// Применяем очистку к колонке "product-points"
const cleanedData = data.map((row) => {
  if (row["product-points"]) {
    row["product-points"] = cleanHtmlBlock(row["product-points"]);
  }
  return row;
});

// Создаём новый Excel-файл с очищенными данными
const newWorksheet = xlsx.utils.json_to_sheet(cleanedData);
const newWorkbook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);

// Сохраняем файл
const outputPath = "./cleaned_part2.xlsx";
xlsx.writeFile(newWorkbook, outputPath);

console.log(`Файл успешно сохранён: ${outputPath}`);
