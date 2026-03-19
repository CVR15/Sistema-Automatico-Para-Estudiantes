/**
 * Crea un menú personalizado para que el administrador dispare la sincronización.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ ADMINISTRADOR')
    .addItem('Sincronizar Mentores Ahora 🔄', 'ejecutarSincronizacionCompletaV2')
    .addItem('Actualizar Forms a Mentores 👥','actualizarRegistroConAlertas')
    .addItem('Actualizar Datos Pestaña Forms 📥', 'actualizaFormsCompleto')
    .addToUi();
}

/**
 * Elimina acentos y diéresis de un texto, convirtiéndolos a sus letras base.
 *
 * @param {"Canción"} inputText El texto o la celda que deseas limpiar.
 * @return {string} El texto original sin acentos.
 * @customfunction
 */
function QuitaAcentos(inputText) {
  if (!inputText) return ""; // Evita errores si la celda está vacía
  
  var oldChars = "áéíóúÁÉÍÓÚäëïöüÄËÏÖÜ";
  var newChars = "aeiouAEIOUaeiouAEIOU";

  for (var i = 0; i < oldChars.length; i++) {
    inputText = inputText.toString().replace(new RegExp(oldChars.charAt(i), 'g'), newChars.charAt(i));
  }

  return inputText;
}

/**
 * Elimina acentos y diéresis de un texto y convierte el resultado a MAYÚSCULAS.
 *
 * @param {"canción"} inputText El texto o la celda que deseas transformar.
 * @return {string} El texto sin acentos y en mayúsculas.
 * @customfunction
 */
function QuitaAcentosMays(inputText) {
  if (!inputText) return "";
  
  // Reutilizamos la lógica de la función anterior y aplicamos toUpperCase()
  var textoSinAcentos = QuitaAcentos(inputText);
  return textoSinAcentos.toUpperCase();
}

