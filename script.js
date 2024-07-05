// Event listeners and initial setup
document.getElementById('fileInput').addEventListener('change', handleFileSelect);
document.getElementById('processButton').addEventListener('click', processSelectedFile);

let selectedFile;

function handleFileSelect(event) {
    selectedFile = event.target.files[0];
    const fileName = selectedFile ? selectedFile.name : '';
    const processButton = document.getElementById('processButton');

    if (!fileName.endsWith('.xls')) {
        alert('Please select an Excel file with .xls');
        selectedFile = null;
        processButton.disabled = true;
    } else {
        processButton.disabled = false;
    }
}

function processSelectedFile() {
    if (selectedFile) {
        processExcelFile(selectedFile).catch(console.error);
    }
}

// Helper functions
function classifyAge(age) {
    const ageGroupsLimit = {
        '1': [0, 4],
        '2': [5, 9],
        '3': [10, 14],
        '4': [15, 19],
        '5': [20, 24],
        '6': [25, 29],
        '7': [30, 34],
        '8': [35, 39],
        '9': [40, 44],
        '10': [45, 49],
        '11': [50, 54],
        '12': [55, 59],
        '13': [60, 64],
        '14': [65, 69],
        '15': [70, 74],
        '16': [75, 79],
        '17': [80, Infinity]
    };

    for (const [group, [start, end]] of Object.entries(ageGroupsLimit)) {
        if (age >= start && age <= end) {
            return group;
        }
    }
    return null;
}

function classifyTime(time) {
    if (time < 12) return 1;
    if (time < 24) return 2;
    return 3;
}

function classifyPrevision(prevision) {
    return prevision === 'FONASA' || prevision === 'PARTICULAR' ? 1 : 0;
}

// Specialty mapping
const specialtyMapping = {
    'OFTALMOLOGIA': 'OFTALMOLOGÍA',
    'TRAUMATOLOGIA GENERAL': 'TRAUMATOLOGÍA Y ORTOPEDIA',
    'CIRUGÍA DIGESTIVA Y COLOPROCTO': 'CIRUGÍA GENERAL',
    'CIRUGIA PLASTICA': 'CIRUGÍA DE CABEZA, CUELLO Y MAXILOFACIAL',
    'CIRUGIA GENERAL': 'CIRUGÍA GENERAL',
    'BRONCOPULMONAR ADULTO': 'MEDICINA INTERNA',  // Adjusted mapping as needed
    'NEUROLOGIA': 'NEUROLOGÍA ADULTOS',
    'NEUROCIRUGIA': 'NEUROCIRUGÍA',
    'MAXILOFACIAL': 'CIRUGÍA DE CABEZA, CUELLO Y MAXILOFACIAL',
    'TRAUMATOLOGIA RODILLA': 'TRAUMATOLOGÍA Y ORTOPEDIA',
    'OTORRINOLARINGOLOGIA': 'OTORRINOLARINGOLOGÍA',
    'UROLOGIA': 'UROLOGÍA',
    'GINECOLOGIA': 'OBSTETRICIA Y GINECOLOGÍA',
    'MEDICINA INTERNA': 'MEDICINA INTERNA',
    'CIRUGIA VASCULAR': 'CIRUGÍA VASCULAR PERIFÉRICA',
    'CIRUGIA DENTAL': 'CIRUGÍA DE CABEZA, CUELLO Y MAXILOFACIAL',
    'ANESTESIOLOGO': 'ANESTESIOLOGÍA',
    'TRAUMATOLOGIA PIE Y TOBILLO': 'TRAUMATOLOGÍA Y ORTOPEDIA',
    'TRAUMATOLOGIA INFANTIL': 'CIRUGÍA PEDIÁTRICA',
    'URGENCIOLOGO': 'URGENCIÓLOGO'
};

// Main processing function
async function processExcelFile(inputFile) {
    // Read the input data
    const inputData = await readExcelFile(inputFile);
    
    // Extract first and last dates
    const { firstDate, lastDate } = extractDates(inputData);
    
    // Fetch the template file (Urgencia.xlsx)
    const templateResponse = await fetch('./en_blanco/Urgencia.xlsx');
    const templateArrayBuffer = await templateResponse.arrayBuffer();
    
    // Load the template workbook using ExcelJS
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateArrayBuffer);
    
    // Get the first sheet
    const sheet = workbook.getWorksheet(1);
  
    // Process the input data
    const cleanedData = inputData.filter(row => row['Diagnóstico'] !== 'NO ESPERA ATENCIÓN' && row['Diagnóstico'] !== 'MAL INGRESADO - FOLIO NULO');
  
    const processedData = cleanedData.map(row => ({
        ...row,
        Age_Group: classifyAge(parseFloat(row['Edad_Años'])),
        Time_ER: (new Date(row['Egreso']) - new Date(row['Ingreso'])) / (1000 * 60 * 60),
        Time_ER_Group: classifyTime((new Date(row['Egreso']) - new Date(row['Ingreso'])) / (1000 * 60 * 60)),
        FONASA: classifyPrevision(row['NOMPREVI'])
    }));

    // Helper functions for generating summary data
    function generateAgeSummary(data) {
        const ageSummary = {};
        data.forEach(row => {
            const group = row.Age_Group;
            const gender = row.sexo;
            if (!ageSummary[group]) {
                ageSummary[group] = { Hombres: 0, Mujeres: 0 };
            }
            ageSummary[group][gender === 'M' ? 'Hombres' : 'Mujeres'] += 1;
        });
        return ageSummary;
    }

    function generateAgeTriageSummary(data) {
        const ageTriageSummary = [];
        data.forEach(row => {
            ageTriageSummary.push({
                Age_Group: row.Age_Group,
                sexo: row.sexo,
                Categorización: row.Categorización,
                count: 1
            });
        });
        return ageTriageSummary.reduce((acc, curr) => {
            const key = `${curr.Age_Group}_${curr.sexo}_${curr.Categorización}`;
            acc[key] = (acc[key] || 0) + curr.count;
            return acc;
        }, {});
    }

    function generateInterSummary(data) {
        const interSummary = {};
        data.forEach(row => {
            const specialty = specialtyMapping[row['Especialidad Inter']] || row['Especialidad Inter'];
            interSummary[specialty] = (interSummary[specialty] || 0) + 1;
        });
        return interSummary;
    }

    function generateHospSummary(data) {
        const hospSummary = {};
        data.forEach(row => {
            const key = `${row.Age_Group}_${row.sexo}_${row.Time_ER_Group}_${row.FONASA}`;
            if (!hospSummary[key]) {
                hospSummary[key] = { count: 0, fonasa: 0 };
            }
            hospSummary[key].count += 1;
            hospSummary[key].fonasa += row.FONASA;
        });
        return hospSummary;
    }
    
    function generateRechazoSummary(data) {
        const rechazoSummary = {};
        data.forEach(row => {
            const key = `${row.Age_Group}_${row.sexo}_${row.Time_ER_Group}_${row.FONASA}`;
            if (!rechazoSummary[key]) {
                rechazoSummary[key] = { count: 0, fonasa: 0 };
            }
            rechazoSummary[key].count += 1;
            rechazoSummary[key].fonasa += row.FONASA;
        });
        return rechazoSummary;
    }
  
    // Generate summary data
    const dfAge = generateAgeSummary(processedData);
    const dfAgeTriage = generateAgeTriageSummary(processedData);
    const dfInter = generateInterSummary(processedData);
    const dfHosp = generateHospSummary(processedData);
    const dfRechazo = generateRechazoSummary(processedData);
  
    // Write sections
    writeSectionA(sheet, dfAge);
    writeSectionB(sheet, dfAgeTriage);
    writeSectionC(sheet, dfInter);
    writeSectionD(sheet, dfHosp, dfRechazo);
    // ... Add more sections as needed
  
    // Generate filename and download the file
    const filename = `REMASEP_${formatDate(firstDate)}_${formatDate(lastDate)}.xlsx`;
    await workbook.xlsx.writeBuffer().then(buffer => {
        saveAs(new Blob([buffer]), filename);
    });
}

// File reading function using XLSX
function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            resolve(XLSX.utils.sheet_to_json(worksheet));
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Function to extract first and last dates
function extractDates(data) {
    const dates = data.map(row => new Date(row['Ingreso'])).sort((a, b) => a - b);
    return {
        firstDate: dates[0],
        lastDate: dates[dates.length - 1]
    };
}

// Function to format date as YYYYMMDD
function formatDate(date) {
    return date.toISOString().slice(0, 10).replace(/-/g, '');
}

// Helper functions for generating summary data
// ... [Keep these functions as they were in the original code]

// Helper function for writing data to Excel while preserving styles
function writeCell(sheet, cellAddress, value) {
    const cell = sheet.getCell(cellAddress);
    cell.value = value;
}

function writeSectionA(sheet, dfAge) {
    const startRow = 12;
    const startColumn = 5; // Column 'E'

    Object.keys(dfAge).forEach((group, index) => {
        const maleColumn = startColumn + 2 * index;
        const femaleColumn = startColumn + 2 * index + 1;

        const maleCell = `${String.fromCharCode(65 + maleColumn)}${startRow}`;
        const femaleCell = `${String.fromCharCode(65 + femaleColumn)}${startRow}`;

        writeCell(sheet, maleCell, dfAge[group]['Hombres'] || 0);
        writeCell(sheet, femaleCell, dfAge[group]['Mujeres'] || 0);
    });
}

function writeSectionB(sheet, dfAgeTriage) {
    const startRow = 21;
    const baseColumn = 5; // Column 'E'
    const categorizacionToRow = {
        '1': 21,
        '2': 22,
        '3': 23,
        '4': 24,
        '5': 25,
        '0': 26
    };

    Object.keys(dfAgeTriage).forEach((key) => {
        const [ageGroup, gender, triage] = key.split('_');
        const count = dfAgeTriage[key];
        const row = categorizacionToRow[triage];
        const column = baseColumn + (parseInt(ageGroup) - 1) * 2 + (gender === 'F' ? 1 : 0);

        const cell = `${String.fromCharCode(65 + column)}${row}`;
        writeCell(sheet, cell, count);
    });
}

function writeSectionC(sheet, dfInter) {
    const startRow = 31;
    const colD = 4; // Column 'D'

    for (let rowIndex = startRow; rowIndex < 50; rowIndex++) {
        const specialtyCell = sheet.getCell(`A${rowIndex}`);
        const specialty = specialtyCell.value;
        if (specialty && dfInter[specialty]) {
            const cell = `D${rowIndex}`;
            writeCell(sheet, cell, dfInter[specialty]);
        }
    }
}

function writeSectionD(sheet, dfHosp, dfRechazo) {
    const baseRow = 56;
    const baseColumn = 6; // Column 'F'
    const totalColumn = 39; // Column 'AN'
    
    Object.keys(dfHosp).forEach((key) => {
        const [ageGroup, gender, timeGroup, fonasa] = key.split('_');
        const { count, fonasa: fonasaCount } = dfHosp[key];
        const row = baseRow + parseInt(timeGroup) - 1;
        const column = baseColumn + (parseInt(ageGroup) - 1) * 2 + (gender === 'F' ? 1 : 0);
        const cell = `${String.fromCharCode(65 + column)}${row}`;
        writeCell(sheet, cell, count);

        const totalCell = `AN${row}`;
        const currentTotal = sheet.getCell(totalCell).value || 0;
        writeCell(sheet, totalCell, currentTotal + fonasaCount);
    });

    const rechazoBaseRow = 60;
    Object.keys(dfRechazo).forEach((key) => {
        const [ageGroup, gender, timeGroup, fonasa] = key.split('_');
        const { count, fonasa: fonasaCount } = dfRechazo[key];
        const row = rechazoBaseRow;
        const column = baseColumn + (parseInt(ageGroup) - 1) * 2 + (gender === 'F' ? 1 : 0);
        const cell = `${String.fromCharCode(65 + column)}${row}`;
        writeCell(sheet, cell, count);

        const totalCell = `AN${row}`;
        const currentTotal = sheet.getCell(totalCell).value || 0;
        writeCell(sheet, totalCell, currentTotal + fonasaCount);
    });
}