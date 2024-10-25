import fs from 'fs';
import path from 'path';
import inquirer from 'inquirer';
import AdmZip from 'adm-zip';
import ExcelJS from 'exceljs';

// Kérdések az inquirer-hez
const questions = [
    {
        type: 'input',
        name: 'zipFilePath',
        message: 'Please enter the path of the ZIP file:',
    },
];

// ZIP fájl feldolgozása
async function processZipFile(zipFilePath) {
    const zip = new AdmZip(zipFilePath);
    const zipEntries = zip.getEntries();

    const data = [];
    
    // Kinyerjük az adatokat a ZIP-ből (mappanevek)
    for (const entry of zipEntries) {
        if (!entry.isDirectory) continue; // Csak a mappákat nézzük

        const folderName = entry.entryName; // Mappa neve
        const parts = folderName.split('_'); // Elválasztás _ karakterrel
        
        // Ellenőrizzük, hogy a mappa neve helyes formátumban van-e
        if (parts.length >= 3) {
            const streamID = parts[parts.length - 1]; // Az utolsó rész az ID
            const streamType = parts[parts.length - 2]; // A második utolsó rész a típus
            const streamName = parts.slice(0, parts.length - 2).join('_'); // Az összes előző rész a név

            data.push([streamID, streamName, streamType]); // Adatok hozzáadása
        }
    }

    return data;
}

// Excel fájl létrehozása
async function createExcelFile(data) {
    const workbook = new ExcelJS.Workbook(); // Használjuk az ExcelJS default exportját
    const worksheet = workbook.addWorksheet('Data');

    // Fejléc hozzáadása
    worksheet.addRow(['Stream ID', 'Stream name', 'Stream type']);

    // Adatok hozzáadása
    data.forEach(row => {
        worksheet.addRow(row);
    });

    // Excel fájl mentése
    const outputFilePath = path.join(process.cwd(), 'output.xlsx');
    await workbook.xlsx.writeFile(outputFilePath);
    console.log(`Excel file created at: ${outputFilePath}`);
}

// Fájlútvonal megadása
async function getFilePath() {
    const answers = await inquirer.prompt(questions);
    const data = await processZipFile(answers.zipFilePath);
    
    if (data.length > 0) {
        console.log('Data from ZIP:', data);
        await createExcelFile(data); // Excel fájl létrehozása
    } else {
        console.log('No data found in the ZIP.');
    }
}

// Fő funkció
async function main() {
    await getFilePath();
}

main().catch(console.error);
