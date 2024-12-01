const fs = require("fs");
const xlsx = require("xlsx");

const pattern = /<div class="text-center">.*?<\/div>/gs;
const outputPath = "./output/cleaned_part1.xlsx";
const filePath = "./input/part1.xlsx";

const workbook = xlsx.readFile(filePath);

function cleanHtmlBlock(text) {
  if (typeof text === "string") {
    // Удаляем целый блок по заданному паттерну
    let cleanedText = text.replace(pattern, "");
    // Удаляем лишние пробелы между HTML-тегами
    cleanedText = cleanedText.replace(/>\s+</g, "><").trim();
    return cleanedText;
  }
  return text;
}

const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

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
xlsx.writeFile(newWorkbook, outputPath);

console.log(`Файл успешно сохранён: ${outputPath}`);
