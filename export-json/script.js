document.getElementById('upload').addEventListener('change', handleFile, false);

let excelData = null;

function handleFile(e) {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = function(event) {
        try {
            const json = JSON.parse(event.target.result);
            excelData = json;
            document.getElementById('download').style.display = 'inline-block';
        } catch (err) {
            alert("Error: El archivo no es un JSON válido.");
        }
    };

    reader.readAsText(file);
}

document.getElementById('download').addEventListener('click', () => {
    if (!excelData) return;

    // Crear una hoja de trabajo (worksheet)
    // Si el JSON es una lista de objetos, se convierte automáticamente
    const data = Array.isArray(excelData) ? excelData : [excelData];
    const ws = XLSX.utils.json_to_sheet(data);

    // Crear un libro de trabajo (workbook)
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Datos");

    // Generar archivo y descargar
    XLSX.writeFile(wb, "reporte_convertido.xlsx");
});
