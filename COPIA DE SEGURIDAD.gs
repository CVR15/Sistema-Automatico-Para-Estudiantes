function crearCopiaDeSeguridadEstatica() {
  // 1. Obtener el archivo original y su nombre
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nombreArchivoOriginal = ss.getName();
  var archivoOriginal = DriveApp.getFileById(ss.getId());
  
  // 2. Obtener la carpeta donde vive el original
  var carpetasParent = archivoOriginal.getParents();
  if (!carpetasParent.hasNext()) {
    Logger.log("El archivo no está en una carpeta accesible.");
    return;
  }
  var carpetaPadre = carpetasParent.next();
  
  // 3. Definir nombre dinámico de la subcarpeta
  var nombreSubcarpeta = "Copias de seguridad " + nombreArchivoOriginal;
  var subcarpeta;
  
  // 4. Buscar o crear la subcarpeta
  var carpetasEncontradas = carpetaPadre.getFoldersByName(nombreSubcarpeta);
  if (carpetasEncontradas.hasNext()) {
    subcarpeta = carpetasEncontradas.next();
  } else {
    subcarpeta = carpetaPadre.createFolder(nombreSubcarpeta);
  }
  
  // 5. Crear el nombre de la copia con la fecha
  var fecha = Utilities.formatDate(new Date(), "GMT-6", "yyyy-MM-dd"); // Ajusta tu zona horaria si es necesario
  var nombreCopia = "Respaldo_" + nombreArchivoOriginal + "_" + fecha;
  
  // 6. Crear la copia en la subcarpeta
  var copiaArchivo = archivoOriginal.makeCopy(nombreCopia, subcarpeta);
  var copiaSpreadsheet = SpreadsheetApp.openById(copiaArchivo.getId());
  
  // 7. Convertir fórmulas a valores en todas las hojas
  var hojas = copiaSpreadsheet.getSheets();
  hojas.forEach(function(hoja) {
    var rango = hoja.getDataRange();
    if (rango.getNumRows() > 0 && rango.getNumColumns() > 0) {
      rango.setValues(rango.getValues());
    }
  });
  
  Logger.log("Respaldo completado en la carpeta: " + nombreSubcarpeta);
}
