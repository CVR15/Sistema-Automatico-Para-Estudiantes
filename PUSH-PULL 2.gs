/**
 * SISTEMA DE GESTIÓN ACADÉMICA - VERSIÓN FINAL REVISADA
 * Incluye: Fechas operativas, Reportes detallados y Fusión de columnas N-Z.
 */

function ejecutarSincronizacionCompletaV2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaMaster = ss.getSheetByName("LISTA_TOTAL");
  const hojaConfig = ss.getSheetByName("CONFIG_MENTORES");

  if (!hojaMaster || !hojaConfig) {
    Logger.log("❌ ERROR: No se encontró LISTA_TOTAL o CONFIG_MENTORES");
    SpreadsheetApp.getUi().alert("❌ ERROR: No se encontró la pestaña 'LISTA_TOTAL' o 'CONFIG_MENTORES'.");
    return;
  }

  const datosConfig = hojaConfig.getDataRange().getValues();
  let directorioMentores = {};
  
  for (let i = 1; i < datosConfig.length; i++) {
    let codigoMentor = datosConfig[i][0].toString().trim();
    let idDoc = datosConfig[i][2] ? datosConfig[i][2].toString().trim() : "";
    let estatus = datosConfig[i][3] ? datosConfig[i][3].toString().toUpperCase() : "";
    if (codigoMentor && estatus !== "INACTIVO" && idDoc.length > 10) {
      directorioMentores[codigoMentor] = idDoc;
    }
  }

  const rangoMaster = hojaMaster.getDataRange();
  const datosMaster = rangoMaster.getValues();
  
  const COL_CODIGO_MENTOR = 1; // B
  const COL_NOMBRE_INI = 3;    // D
  const COL_CURP = 4;          // E
  const COL_SITUACION = 12;    // M (Master)
  const COL_FECHA = 13;        // N (Master)
  const COL_COMENTARIOS = 15;  // P (Master)

  let mentoresSincronizados = 0;
  let errores = [];

  for (let codigo in directorioMentores) {
    try {
      const ssExterno = SpreadsheetApp.openById(directorioMentores[codigo]);
      const hojaExterna = ssExterno.getSheetByName("LISTA");
      if (!hojaExterna) {
        errores.push("Mentor " + codigo + ": No existe pestaña 'LISTA'");
        continue;
      }
      
      const ultFilaActual = hojaExterna.getLastRow();
      let infoCompletaMentor = new Map();
      
      // --- FASE PULL ---
      if (ultFilaActual > 1) {
        const datosCuerpo = hojaExterna.getRange(2, 2, ultFilaActual - 1, 25).getValues();
        datosCuerpo.forEach(fila => {
          let curp = fila[0] ? fila[0].toString().trim() : "";
          if (curp) infoCompletaMentor.set(curp, fila); 
        });
      }

      let filasParaEnviar = [];
      let curpsEnListaNueva = new Set();

      for (let j = 1; j < datosMaster.length; j++) {
        if (datosMaster[j][COL_CODIGO_MENTOR].toString().trim() === codigo) {
          let curpMaster = datosMaster[j][COL_CURP].toString().trim();
          curpsEnListaNueva.add(curpMaster);

          // 1. Crear el recorte base (D a P del Master) - Son 13 columnas
          let filaRecorte = datosMaster[j].slice(COL_NOMBRE_INI, COL_COMENTARIOS + 1);

          // 2. FUSIÓN Y ACTUALIZACIÓN
          if (infoCompletaMentor.has(curpMaster)) {
            let dAnt = infoCompletaMentor.get(curpMaster);
            
            // ... (tus actualizaciones de Master y Recorte actuales) ...
            datosMaster[j][COL_SITUACION] = dAnt[8];  
            datosMaster[j][COL_FECHA] = dAnt[9];      
            datosMaster[j][COL_COMENTARIOS] = dAnt[11]; 
            filaRecorte[9]  = dAnt[8];  
            filaRecorte[10] = dAnt[9];  
            filaRecorte[12] = dAnt[11]; 

            if (filaRecorte[8] === "true") filaRecorte[8] = true;
            if (filaRecorte[8] === "false") filaRecorte[8] = false;

            // Fusionamos las extras que ya existían
            let extras = dAnt.slice(12); 
            filaRecorte = filaRecorte.concat(extras);
          } else {
            // --- NUEVO AJUSTE: Si el alumno es nuevo, rellenamos hasta la columna Z ---
            // Si el recorte tiene 13, y el mentor llega hasta la 26 (columna Z),
            // necesitamos añadir 13 espacios vacíos para que el ancho total sea 26.
            let espaciosVacios = new Array(13).fill(""); 
            filaRecorte = filaRecorte.concat(espaciosVacios);
          }
          
          filasParaEnviar.push(filaRecorte);
        }
      }

      if (filasParaEnviar.length > 0) {
        if (hojaExterna.getLastRow() > 1) {
          hojaExterna.getRange(2, 1, hojaExterna.getLastRow(), hojaExterna.getLastColumn()).clearContent();
        }

        const rangoDestino = hojaExterna.getRange(2, 1, filasParaEnviar.length, filasParaEnviar[0].length);
        rangoDestino.setValues(filasParaEnviar);
        
        // --- AJUSTE DE FORMATOS ---
        // 1. Columnas A y B: Texto plano (para códigos e identidades)
        hojaExterna.getRange(2, 1, filasParaEnviar.length, 2).setNumberFormat("@"); 

        // 2. Columna K: Fecha
        hojaExterna.getRange(2, 11, filasParaEnviar.length, 1).setNumberFormat("dd/mmm/yyyy"); 

        // 3. NUEVO - Columna L: Booleano Real (Fórmulas)
        // Aplicamos formato "General" o vaciamos el formato de texto para que Sheets reconozca el True/False
        const rangoL = hojaExterna.getRange(2, 12, filasParaEnviar.length, 1);
        rangoL.setNumberFormat("General"); 
        rangoL.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build()); //Formato de checkbox        

        // --- VALIDACIONES ---
        for (let k = 0; k < filasParaEnviar.length; k++) {
          const filaA = k + 2;
          const valorI = filasParaEnviar[k][8]; 
          let valorJ = filasParaEnviar[k][9];   
          
          if (valorI === "") continue;
          let opciones = (valorI === "NVO.") ? ["Sin contactar", "Intento", "Contactado", "Aceptó Inv", "Rechazo"] : 
                         (valorI === "CONT.") ? ["Sin contactar", "Reactivado", "Rechazo"] : [];

          if (opciones.length > 0) {
            const regla = SpreadsheetApp.newDataValidation().requireValueInList(opciones, true).build();
            hojaExterna.getRange(filaA, 10).setDataValidation(regla);
            
            if (valorJ === "" || valorJ === null) {
              let hoy = new Date();
              hojaExterna.getRange(filaA, 10).setValue("Sin contactar");
              hojaExterna.getRange(filaA, 11).setValue(hoy);
              // Actualizar matriz master
              for (let m = 1; m < datosMaster.length; m++) {
                if (datosMaster[m][COL_CURP].toString().trim() === filasParaEnviar[k][1]) {
                  datosMaster[m][COL_SITUACION] = "Sin contactar";
                  datosMaster[m][COL_FECHA] = hoy;
                }
              }
            }
          }
        }

        // --- HISTÓRICO ---
        let filasHuerfanas = [];
        infoCompletaMentor.forEach((datosFila, curpKey) => {
          if (!curpsEnListaNueva.has(curpKey)) filasHuerfanas.push(datosFila);
        });

        if (filasHuerfanas.length > 0) {
          const filaInicioH = filasParaEnviar.length + 10;
          hojaExterna.getRange(filaInicioH - 1, 1).setValue("⚠️ HISTÓRICO").setFontWeight("bold").setFontColor("red");
          hojaExterna.getRange(filaInicioH, 2, filasHuerfanas.length, filasHuerfanas[0].length).setValues(filasHuerfanas).setBackground("#f3f3f3");
          Logger.log("📂 Mentor " + codigo + ": Se archivaron " + filasHuerfanas.length + " alumnos.");
        }
      }
   
      mentoresSincronizados++;
      ss.toast("Sincronizando: " + codigo, "Progreso", 2);
      Logger.log("✅ Mentor " + codigo + " sincronizado con éxito.");

    } catch (e) {
      let msgErr = "Mentor " + codigo + ": " + e.message;
      errores.push(msgErr);
      Logger.log("❌ " + msgErr);
    }
  }

  // 4. VOLCAR TODOS LOS CAMBIOS A LA HOJA MASTER
  Logger.log("💾 Iniciando volcado en LISTA_TOTAL...");

  try {
    // Definimos el rango de la columna de códigos (Columna B)
    const rangoCodigoMentor = hojaMaster.getRange(2, COL_CODIGO_MENTOR + 1, datosMaster.length - 1, 1);
    
    // PASO CLAVE: Cambiamos a formato "Automático" antes de escribir para evitar el bloqueo de la Tabla
    rangoCodigoMentor.setNumberFormat("General"); 
    
    // Volcamos todos los datos a la tabla
    rangoMaster.setValues(datosMaster);
    
    // Una vez volcados, volvemos a poner formato de texto SOLO si es necesario
    // rangoCodigoMentor.setNumberFormat("@"); 

    Logger.log("✅ Cambios volcados con éxito en la Tabla Master.");
  } catch (e) {
    Logger.log("⚠️ Error al volcar en Tabla: " + e.message);
    // Si falla por el formato, intentamos volcar sin tocar formatos
    rangoMaster.setValues(datosMaster);
  }

  Logger.log("💾 Cambios volcados en LISTA_TOTAL.");

  // 5. FINALIZACIÓN
  if (errores.length > 0) {
    const reporteFinal = "Terminado con advertencias:\n" + errores.join("\n");
    SpreadsheetApp.getUi().alert(reporteFinal);
    console.log("⚠️ Sincronización finalizada con errores acumulados.\n" + errores.join("\n"));
  } else {
    ss.toast("✅ " + mentoresSincronizados + " mentores sincronizados correctamente.", "Finalizado");
    console.log("🚀 Proceso completado exitosamente: " + mentoresSincronizados + " mentores.");
  }
}
