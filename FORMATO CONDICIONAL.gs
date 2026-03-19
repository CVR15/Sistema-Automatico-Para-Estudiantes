/**
 * Script Optimizado para LISTA_TOTAL con Gestión de Columna O e Interfaz Unificada
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const col = range.getColumn();
  const row = range.getRow();
  
  const TARGET_SHEET = "LISTA_TOTAL"; 
  const MAX_ROW = 3000;
  const START_ROW = 2;
  const NUM_COLS = 16; // A hasta P

  // Monitorear cambios en A(1), B(2), D(4), E(5), M(13) y O(15)
  if (sheetName === TARGET_SHEET && row >= START_ROW && row <= MAX_ROW && [1, 2, 4, 5, 13, 15].includes(col)) {
    
    const dataRange = sheet.getRange(START_ROW, 1, MAX_ROW - START_ROW + 1, NUM_COLS);
    const data = dataRange.getValues();
    
    const backgrounds = [];
    const fontWeights = [];

    const colDValues = data.map(r => r[3]);
    const colEValues = data.map(r => r[4]);

    for (let i = 0; i < data.length; i++) {
      let rowStyles = new Array(NUM_COLS).fill(null);
      let rowWeights = new Array(NUM_COLS).fill("normal");

      const valA = data[i][0];
      const valB = data[i][1];
      const valD = data[i][3];
      const valE = data[i][4];
      const valM = data[i][12]; // Columna M
      const valO = data[i][14]; // Columna O (Índice 14)
      const valB_Ant = (i > 0) ? data[i-1][1] : null;

      // --- NIVEL 5: Columna M (Reactivado / Aceptó Inv) ---
      if (valM === "Reactivado" || valM === "Aceptó Inv") {
        rowStyles[12] = "#b6d7a8"; 
      }

      // --- NIVEL 4: Columna A (Sección A hasta L) ---
      let colorA = null;
      if (valA === 1) colorA = "#b4a7d6";
      else if (valA === 2) colorA = "#a4c2f4";
      else if (valA === 3) colorA = "#f6b26b";
      else if (valA === 4) colorA = "#93c47d";

      if (colorA) {
        for (let j = 0; j <= 11; j++) { rowStyles[j] = colorA; }
      }

      // --- NIVEL 3: Fila Amarilla por Columna B (A:P) ---
      if (i > 0 && valB !== "" && valB !== valB_Ant) {
        rowStyles.fill("#ffe599"); 
      }

      // --- NIVEL 2.5: Estatus de Registro (Columna O) ---
      // Aplicamos estilos específicos solo a la celda O
      if (valO === "REVISAR") {
        rowStyles[14] = "#ff9900"; // Naranja brillante
        rowWeights[14] = "bold";
      } else if (valO === true) {
        rowStyles[14] = "#d9ead3"; // Verde muy tenue (Registrado)
      } else if (valO === false) {
        rowStyles[14] = "#efefef"; // Gris claro (Pendiente)
      }

      // --- NIVEL 2: Rechazo en Columna M ---
      if (valM === "Rechazo") {
        for (let j = 3; j <= 12; j++) { rowStyles[j] = "#ff0000"; }
      }

      // --- NIVEL 1: Duplicados D y E (Máxima Prioridad) ---
      if (valD !== "" && colDValues.indexOf(valD) !== colDValues.lastIndexOf(valD)) {
        rowStyles[3] = "#00ffff"; 
        rowWeights[3] = "bold";
      }
      if (valE !== "" && colEValues.indexOf(valE) !== colEValues.lastIndexOf(valE)) {
        rowStyles[4] = "#00ffff"; 
        rowWeights[4] = "bold";
      }

      backgrounds.push(rowStyles);
      fontWeights.push(rowWeights);
    }

    dataRange.setBackgrounds(backgrounds);
    dataRange.setFontWeights(fontWeights);
  }
}
