document.getElementById('fileInput').addEventListener('change', handleFileSelect);
document.getElementById('button').disabled = true;

let selectedFile;

function handleFileSelect(event) {
    selectedFile = event.target.files[0];
    const fileName = selectedFile.name;

    // Check for .xlsx extension
    if (!fileName.endsWith('.xlsx')) {
        alert('Por favor seleccione un archivo Excel con la extensión .xlsx');
        selectedFile = null;
    }

    document.getElementById('button').disabled = false;


}
