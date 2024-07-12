import * as XLSX from 'xlsx';

interface RowData {
    [key: string]: any;
}

function splitColumnData(data: RowData[], column: string): RowData[] {
    const newData: RowData[] = [];

    data.forEach(row => {
        const rowObj: RowData = { ...row };
        const originalColumns = Object.keys(rowObj);
        const insertIndex = originalColumns.indexOf(column);
        if (insertIndex === -1) return;
        const columnData = rowObj[column] as string;
        const entries = columnData.split('\n').map(entry => entry.trim()).filter(entry => entry);

        entries.forEach(entry => {
            const newRow = { ...rowObj };

            Object.keys(outputPaths).forEach(colName => {
                delete newRow[colName];
            });
            // Split the entry into key-value pairs
            const keyValuePairs = entry.split(/,\s*(?![^()]*\))(?=\s*[A-Za-z ]+:\s*)/).map(pair => pair.trim());
            const newColumns: RowData = {};
            // if (keyValuePairs.length !== 8) {
            //     debugger;
            // }
            keyValuePairs.forEach(pair => {
                let [key, value] = pair.split(':').map(str => str.trim());

                let alias = '';
                if (column === 'Educatore Nido in Famiglia') alias = 'Educatore';
                else if (column === 'EDUCATORE NIDO IN FAMIGLIA ATTIVO NON TITOLARE (Partner)') alias = 'Educatore Attivo';
                else if (column === 'EDUCATORI ISCRITTI CONVENZIONATI MA NON ATTIVI') alias = 'Educatore Iscritto non Attivo';
                if (key.toLowerCase() === "ninfa") {
                    if (value === 'Socia' || value === 'Enif 298' || value === 'Sì' || value === "Si")
                        value = "SI";
                    else value = "NO";
                }
                newColumns[`${key} (${alias})`] = value;
            });

            const updatedRow = Object.entries(newRow);
            updatedRow.splice(insertIndex, 0, ...Object.entries(newColumns));
            const rowToPush = Object.fromEntries(updatedRow);
            newData.push(rowToPush);
        });

    });

    return newData;
}

function transformExcel(inputPath: string, outputPaths: { [key: string]: string }) {
    const workbook = XLSX.readFile(inputPath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    const jsonData: RowData[] = XLSX.utils.sheet_to_json(worksheet);

    const columnsToSplit = Object.keys(outputPaths);

    columnsToSplit.forEach(column => {
        const splitData = splitColumnData([...jsonData], column);
        const newWorksheet = XLSX.utils.json_to_sheet(splitData);
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1');
        XLSX.writeFile(newWorkbook, outputPaths[column]);
    });
}

// Define input and output paths
const inputPath = './docs/EDUCATORI_CONVENZIONATI2024-05-19_03_00_41.xlsx';
const outputPaths = {
    'Educatore Nido in Famiglia': './docs/expanded_educatori_nido_in_famiglia.xlsx',
    'EDUCATORE NIDO IN FAMIGLIA ATTIVO NON TITOLARE (Partner)': './docs/expanded_educatore_nido_in_famiglia_attivo_non_titolare.xlsx',
    'EDUCATORI ISCRITTI CONVENZIONATI MA NON ATTIVI': './docs/expanded_educatori_iscritti_convenzionati_non_attivi.xlsx'
};

// Execute the transformation
transformExcel(inputPath, outputPaths);
console.log('Transformation complete. Files saved to:');
console.log(outputPaths);