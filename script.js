document.getElementById('processButton').addEventListener('click', async () => {
    const fileInput = document.getElementById('fileInput');
    if (fileInput.files.length === 0) {
      alert('Please select an Excel file.');
      return;
    }
  
    const file = fileInput.files[0];
    const reader = new FileReader();
  
    reader.onload = async function (e) {
      try {
        const data = e.target.result;
  
        // Read the uploaded file
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(data);
  
        // Process the workbook (e.g., read data, modify it)
        const worksheet = workbook.getWorksheet(1);
        if (!worksheet) {
          throw new Error('Worksheet not found in the uploaded file.');
        }
  
        // Example: Log the first cell's value
        console.log(worksheet.getCell('A1').value);
  
        // Fetch the template file
        const templateResponse = await fetch('templates/Urgencia.xlsx');
        if (!templateResponse.ok) {
          throw new Error('Failed to fetch the template file.');
        }
  
        const templateArrayBuffer = await templateResponse.arrayBuffer();
  
        const templateWorkbook = new ExcelJS.Workbook();
        await templateWorkbook.xlsx.load(templateArrayBuffer);
  
        // Copy data from the uploaded file to the template while preserving styles
        const templateWorksheet = templateWorkbook.getWorksheet(1);
        if (!templateWorksheet) {
          throw new Error('Worksheet not found in the template file.');
        }
  
        worksheet.eachRow((row, rowNumber) => {
          row.eachCell((cell, colNumber) => {
            const templateCell = templateWorksheet.getCell(rowNumber, colNumber);
            templateCell.value = cell.value;
  
            // Copy styles
            templateCell.style = cell.style;
          });
        });
  
        // Generate a new Excel file with the processed data
        const buffer = await templateWorkbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  
        // Create a link to download the file
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'Urgencia_processed.xlsx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
  
      } catch (error) {
        console.error('Error processing file:', error);
        alert('Error processing file: ' + error.message);
      }
    };
  
    reader.onerror = function (error) {
      console.error('FileReader error:', error);
      alert('FileReader error: ' + error.message);
    };
  
    reader.readAsArrayBuffer(file);
  });