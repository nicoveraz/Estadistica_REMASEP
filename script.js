// Global variable to store the selected file
let selectedFile;

// Event listener for file input change
document.getElementById('fileInput').addEventListener('change', handleFileSelect);

// Function to handle file selection
function handleFileSelect(event) {
    selectedFile = event.target.files[0];
    const fileName = selectedFile ? selectedFile.name : '';
    const processButton = document.getElementById('processButton');

    // Check for .xls or .xlsx extension
    if (!fileName.endsWith('.xls') && !fileName.endsWith('.xlsx')) {
        alert('Por favor seleccione un archivo con la extensión .xls o .xlsx.');
        selectedFile = null;
        processButton.disabled = true;
    } else {
        processButton.disabled = false;
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
    'BRONCOPULMONAR ADULTO': 'MEDICINA INTERNA',
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

// Updated File reading function
async function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});
                
                // Process the data
                const headers = jsonData[0];
                const rows = jsonData.slice(1);
                const processedData = rows.map(row => {
                    const obj = {};
                    headers.forEach((header, index) => {
                        obj[header] = row[index];
                    });
                    return obj;
                });
                
                resolve(processedData);
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Updated main processing function
async function processExcelFile(inputFile) {
    try {
        // Read the input data
        const inputData = await readExcelFile(inputFile);
        
        // Extract first and last dates
        const { firstDate, lastDate } = extractDates(inputData);
        
        // Create a new workbook
        const workbook = new ExcelJS.Workbook();
        
        // Load the template workbook
        const templateResponse = await fetch('./en_blanco/Urgencia.xlsx');
        const templateArrayBuffer = await templateResponse.arrayBuffer();
        await workbook.xlsx.load(templateArrayBuffer);

        // Get the first sheet
        const sheet = workbook.getWorksheet(1);

        if (!sheet) {
            throw new Error('Worksheet not found in the template file.');
        }
      
        // Process the input data
        const cleanedData = inputData.filter(row => row['Diagnóstico'] !== 'NO ESPERA ATENCIÓN' && row['Diagnóstico'] !== 'MAL INGRESADO - FOLIO NULO');
      
        const processedData = cleanedData.map(row => ({
            ...row,
            Age_Group: classifyAge(parseFloat(row['Edad_Años'])),
            Time_ER: (new Date(row['Egreso']) - new Date(row['Ingreso'])) / (1000 * 60 * 60),
            Time_ER_Group: classifyTime((new Date(row['Egreso']) - new Date(row['Ingreso'])) / (1000 * 60 * 60)),
            FONASA: classifyPrevision(row['NOMPREVI'])
        }));
      
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
      
        // Generate filename and download the file
        const filename = `REMASEP_${formatDate(firstDate)}_${formatDate(lastDate)}.xlsx`;
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = filename;
        link.click();
    } catch (error) {
        console.error('Error processing Excel file:', error);
        console.error('Error stack:', error.stack);
        alert('An error occurred while processing the file. Please check the console for more details.');
    }
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
    const ageTriageSummary = {};
    data.forEach(row => {
        const key = `${row.Age_Group}_${row.sexo}_${row.Categorización}`;
        ageTriageSummary[key] = (ageTriageSummary[key] || 0) + 1;
    });
    return ageTriageSummary;
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

// Helper functions for writing data to Excel while preserving styles
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

        const maleCell = `${String.fromCharCode(64 + maleColumn)}${startRow}`;
        const femaleCell = `${String.fromCharCode(64 + femaleColumn)}${startRow}`;

        writeCell(sheet, maleCell, dfAge[group]['Hombres'] || 0);
        writeCell(sheet, femaleCell, dfAge[group]['Mujeres'] || 0);
    });
}

function writeSectionB(sheet, dfAgeTriage) {
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

        const cell = `${String.fromCharCode(64 + column)}${row}`;
        writeCell(sheet, cell, count);
    });
}

function writeSectionC(sheet, dfInter) {
    const startRow = 31;
    const colD = 4; // Column 'D' is the 4th column (1-based index)

    for (let rowIndex = startRow; rowIndex < 50; rowIndex++) {
        const specialtyCell = sheet.getCell(`A${rowIndex}`);
        const specialty = specialtyCell.value;
        if (specialty && dfInter[specialty]) {
            const cell = sheet.getCell(`D${rowIndex}`);
            cell.value = dfInter[specialty];
        }
    }
}

function writeSectionD(sheet, dfHosp, dfRechazo) {
    const baseRow = 56;
    const baseColumn = 6; // Column 'F' is the 6th column (1-based index)
    const totalColumn = 40; // Column 'AN' is the 40th column (1-based index)

    Object.keys(dfHosp).forEach((key) => {
        const [ageGroup, gender, timeGroup, fonasa] = key.split('_');
        const { count, fonasa: fonasaCount } = dfHosp[key];
        const row = baseRow + parseInt(timeGroup) - 1;
        const column = baseColumn + (parseInt(ageGroup) - 1) * 2 + (gender === 'F' ? 1 : 0);
        const cell = `${String.fromCharCode(64 + column)}${row}`;
        writeCell(sheet, cell, count);

        const totalCell = `${String.fromCharCode(64 + totalColumn)}${row}`;
        writeCell(sheet, totalCell, (sheet.getCell(totalCell).value || 0) + fonasaCount);
    });

    const rechazoBaseRow = 60;
    Object.keys(dfRechazo).forEach((key) => {
        const [ageGroup, gender, timeGroup, fonasa] = key.split('_');
        const { count, fonasa: fonasaCount } = dfRechazo[key];
        const row = rechazoBaseRow;
        const column = baseColumn + (parseInt(ageGroup) - 1) * 2 + (gender === 'F' ? 1 : 0);
        const cell = `${String.fromCharCode(64 + column)}${row}`;
        writeCell(sheet, cell, count);

        const totalCell = `${String.fromCharCode(64 + totalColumn)}${row}`;