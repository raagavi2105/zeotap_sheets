// Download functionality
function downloadSheet() {
    // Get the current state of the spreadsheet
    const sheetData = {
        matrix: spreadsheetMatrix.matrix,
        currentSheet: currentSheet,
        sheets: sheets
    };

    // Convert to JSON string
    const jsonString = JSON.stringify(sheetData, null, 2);

    // Create a Blob with the JSON data
    const blob = new Blob([jsonString], { type: 'application/json' });

    // Create a temporary download link
    const downloadLink = document.createElement('a');
    downloadLink.href = URL.createObjectURL(blob);
    
    // Set the filename with current date
    const date = new Date();
    const fileName = `spreadsheet_${date.getFullYear()}${(date.getMonth() + 1).toString().padStart(2, '0')}${date.getDate().toString().padStart(2, '0')}.json`;
    downloadLink.download = fileName;

    // Append link to body, click it, and remove it
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);

    // Clean up the URL object
    URL.revokeObjectURL(downloadLink.href);
}

// Add keyboard shortcut for download (Ctrl + S)
document.addEventListener('keydown', function(event) {
    if (event.ctrlKey && event.key === 's') {
        event.preventDefault(); // Prevent default browser save
        downloadSheet();
    }
});