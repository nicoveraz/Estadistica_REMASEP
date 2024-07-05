document.getElementById('fileInput').addEventListener('change', handleFileSelect);

let selectedFile;

function handleFileSelect(event) {
    selectedFile = event.target.files[0];
    const fileName = selectedFile ? selectedFile.name : '';
    const processButton = document.getElementById('processButton');

    // Check for .xlsx extension
    if (!fileName.endsWith('.xls')) {
        alert('Please select an Excel file with .xls extension.');
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
    
    // Load the template workbook
    const workbook = XLSX.read(new Uint8Array(templateArrayBuffer), {type: 'array'});
    
    // Get the first sheet
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
  
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
    // ... Add more sections as needed
  
    // Generate filename and download the file
    const filename = `REMASEP_${formatDate(firstDate)}_${formatDate(lastDate)}.xlsx`;
    XLSX.writeFile(workbook, filename);
}

// File reading function
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

// Helper functions for writing data to Excel
function writeSectionA(sheet, dfAge) {
    const startRow = 12;
    const startColumn = 5; // Column 'E'

    Object.keys(dfAge).forEach((group, index) => {
        const maleColumn = startColumn + 2 * index;
        const femaleColumn = startColumn + 2 * index + 1;

        const maleCell = XLSX.utils.encode_cell({ r: startRow - 1, c: maleColumn });
        const femaleCell = XLSX.utils.encode_cell({ r: startRow - 1, c: femaleColumn });

        sheet[maleCell] = { t: 'n', v: dfAge[group]['Hombres'] || 0 };
        sheet[femaleCell] = { t: 'n', v: dfAge[group]['Mujeres'] || 0 };
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

        const cell = XLSX.utils.encode_cell({ r: row - 1, c: column });
        sheet[cell] = { t: 'n', v: count };
    });
}

function writeSectionC(sheet, dfInter) {
    const startRow = 31;
    const colD = 3; // Column 'D'

    for (let rowIndex = startRow; rowIndex < 50; rowIndex++) {
        const specialtyCell = XLSX.utils.encode_cell({ r: rowIndex - 1, c: 0 });
        const specialty = sheet[specialtyCell] ? sheet[specialtyCell].v : null;
        if (specialty && dfInter[specialty]) {
            const cell = XLSX.utils.encode_cell({ r: rowIndex - 1, c: colD });
            sheet[cell] = { t: 'n', v: dfInter[specialty] };
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
        const cell = XLSX.utils.encode_cell({ r: row - 1, c: column });
        sheet[cell] = { t: 'n', v: count };

        const totalCell = XLSX.utils.encode_cell({ r: row - 1, c: totalColumn });
        sheet[totalCell] = { t: 'n', v: (sheet[totalCell] ? sheet[totalCell].v : 0) + fonasaCount };
    });

    const rechazoBaseRow = 60;
    Object.keys(dfRechazo).forEach((key) => {
        const [ageGroup, gender, timeGroup, fonasa] = key.split('_');
        const { count, fonasa: fonasaCount } = dfRechazo[key];
        const row = rechazoBaseRow;
        const column = baseColumn + (parseInt(ageGroup) - 1) * 2 + (gender === 'F' ? 1 : 0);
        const cell = XLSX.utils.encode_cell({ r: row - 1, c: column });
        sheet[cell] = { t: 'n', v: count };

        const totalCell = XLSX.utils.encode_cell({ r: row - 1, c: totalColumn });
        sheet[totalCell] = { t: 'n', v: (sheet[totalCell] ? sheet[totalCell].v : 0) + fonasaCount };
    });
}

// Event listener for process button
document.getElementById('processButton').addEventListener('click', () => {
    if (selectedFile) {
        processExcelFile(selectedFile).catch(console.error);
    }
});