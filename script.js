document.getElementById('fileInput').addEventListener('change', handleFileSelect);

let selectedFile;

function handleFileSelect(event) {
    selectedFile = event.target.files[0];
    const fileName = selectedFile.name;

    // Check for .xlsx extension
    if (!fileName.endsWith('.xlsx')) {
        alert('Por favor seleccione un archivo excel con la extensión .xlsx.');
        selectedFile = null;
        processButton.disabled = true;
        return;
    }

    document.getElementById('selectedFileName').textContent = fileName;
    processButton.disabled = false;
    
}

function processFile() {
    if (!selectedFile) {
        alert('Por favor seleccione un archivo Excel');
        return;
    }
    const reader = new FileReader();

    reader.onload = function (event) {
        const data = event.target.result;
        const workbook = XLSX.read

        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet);

        // Process the data here
        console.log(data);
    };

    reader.readAsBinaryString(selectedFile);
}