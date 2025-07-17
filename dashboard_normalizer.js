// Modifies the date format of the column "Fecha de Entrega" to work properly in the dashboard

import XLSX from "xlsx";

const file = "./apoyos.xlsx";

const workbook = XLSX.readFile(file);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet);

for (const row of data) {
    const yy = row["Fecha de Entrega"].slice(0, 4)
    const mm = row["Fecha de Entrega"].slice(5, 7);
    const dd = row["Fecha de Entrega"].slice(8, 10);
    let date = `${mm}/${dd}/${yy}`;

    console.log(date);
    row["Fecha de Entrega"] = date;
};

const newSheet = XLSX.utils.json_to_sheet(data);

workbook.Sheets[workbook.SheetNames[0]] = newSheet;
XLSX.writeFile(workbook, "apoyos.xlsx");
