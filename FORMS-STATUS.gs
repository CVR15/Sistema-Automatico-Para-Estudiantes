/**
 * Sincronización de registros con salida Booleana y Alerta por Correo.
 */
function actualizarRegistroConAlertas() {
  const ID_FORMULARIO_1 = "INSETA_EL_ID_AQUI"; //ACTUALIZACION DE DATOS 2026
  const ID_FORMULARIO_2 = "INSETA_EL_ID_AQUI"; //NUEVOS MIEMBROS 2026
  const EMAIL_NOTIFICACION = "INSETA_TU_CORREO_AQUI"; // Correo de notificación
  
  const ssMaster = SpreadsheetApp.getActiveSpreadsheet();
  const hojaMaster = ssMaster.getSheetByName("LISTA_TOTAL");
  
  if (!hojaMaster) return;

  // 1. CARGAR DATOS DE FORMULARIOS EXTERNOS
  let registrosForms = [];
  [ID_FORMULARIO_1, ID_FORMULARIO_2].forEach(id => {
    try {
      const valores = SpreadsheetApp.openById(id).getSheets()[0].getDataRange().getValues();
      valores.shift();
      valores.forEach(f => registrosForms.push({ nombre: limpiar(f[6]), curp: limpiar(f[10]) }));
    } catch (e) { console.log("Error en ID: " + id); }
  });

  // 2. PROCESAR LISTA MASTER (Columna D=Nombre, E=CURP, O=Estado)
  const datosMaster = hojaMaster.getDataRange().getValues();
  const resultadosO = [];
  const casosParaCorreo = [];

  for (let i = 1; i < datosMaster.length; i++) {
    const nombreM = limpiar(datosMaster[i][3]); // Columna D
    const curpM = limpiar(datosMaster[i][4]);   // Columna E
    const estadoActual = datosMaster[i][14];    // Columna O (índice 14)

    // REGLA DE SALTO: Si ya es TRUE o ya fue marcado como REVISAR/OMITIR, omitimos y pasamos al siguiente.
    if (estadoActual === true || estadoActual === "REVISAR" || estadoActual === "OMITIR") {
      resultadosO.push([estadoActual]);
      continue;
    }

    let encontrado = false;
    let dudaRazonable = false;
    let infoDuda = "";

    for (let reg of registrosForms) {
      const distCurp = calcularDistancia(curpM, reg.curp);
      
      // A. COINCIDENCIA TOTAL
      if (curpM !== "" && curpM === reg.curp) {
        encontrado = true;
        break;
      }
      
      // B. COINCIDENCIA DUDOSA (Error de 1-2 letras en CURP + Nombre similar)
      if (curpM.length > 10 && distCurp > 0 && distCurp <= 2) {
        if (calcularDistancia(nombreM, reg.nombre) < 5) {
          encontrado = true;
          break;
        } else {
          dudaRazonable = true;
          infoDuda = `Fila ${i+1}: Master(${curpM} - ${nombreM}) vs Form(${reg.curp} - ${reg.nombre})`;
        }
      }
    }

    // ASIGNACIÓN DE ESTADOS
    if (encontrado) {
      resultadosO.push([true]);
    } else if (dudaRazonable) {
      resultadosO.push(["REVISAR"]); // Esto frena futuros correos para esta fila
      casosParaCorreo.push(infoDuda);
    } else {
      resultadosO.push([false]); // Sigue buscando en la próxima corrida por si llena el form después
    }
  }

  // 3. ACTUALIZACIÓN MASIVA
  if (resultadosO.length > 0) {
    hojaMaster.getRange(2, 15, resultadosO.length, 1).setValues(resultadosO);
  }

  // 4. ENVÍO DE CORREO (Solo si hay nuevos casos de REVISAR)
  if (casosParaCorreo.length > 0) {
    const cuerpo = "Se han detectado registros que coinciden parcialmente. " +
                   "Se marcaron como 'REVISAR' en la columna O para tu inspección:\n\n" + 
                   casosParaCorreo.join("\n\n");
    MailApp.sendEmail(EMAIL_NOTIFICACION, "⚠️ Nuevas revisiones manuales requeridas", cuerpo);
  }
}

// Funciones auxiliares (Limpiar y Distancia)
function limpiar(t) { return t ? t.toString().trim().toUpperCase().replace(/\s+/g, ' ') : ""; }

function calcularDistancia(s1, s2) {
  const m = s1.length, n = s2.length;
  let dp = Array.from({ length: m + 1 }, (_, i) => [i]);
  for (let j = 1; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] = s1[i-1] === s2[j-1] ? dp[i-1][j-1] : Math.min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1]) + 1;
    }
  }
  return dp[m][n];
}
