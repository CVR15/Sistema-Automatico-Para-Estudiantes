/**
 * SISTEMA DE GESTIÓN ACADÉMICA - SINCRONIZACIÓN DINÁMICA
 * * Pestañas requeridas: 
 * 1. LISTA_TOTAL (Base de datos principal)
 * 2. CONFIG_MENTORES (Directorio de enlaces y IDs)
 */

function ejecutarSincronizacionCompleta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaMaster = ss.getSheetByName("LISTA_TOTAL");
  const hojaConfig = ss.getSheetByName("CONFIG_MENTORES");

  if (!hojaMaster || !hojaConfig) {
    SpreadsheetApp.getUi().alert("❌ ERROR: No se encontró la pestaña 'LISTA_TOTAL' o 'CONFIG_MENTORES'.");
    return;
  }

  // 1. CARGAR DIRECTORIO
  const datosConfig = hojaConfig.getDataRange().getValues();
  let directorioMentores = {};
  
  for (let i = 1; i < datosConfig.length; i++) {
    let codigoMentor = datosConfig[i][0].toString().trim();
    let idDoc = datosConfig[i][2] ? datosConfig[i][2].toString().trim() : "";
    let estatus = datosConfig[i][3] ? datosConfig[i][3].toString().toUpperCase() : "";
    
    if (codigoMentor && estatus !== "INACTIVO") {
      if (idDoc !== "" && idDoc !== "LINK INVÁLIDO O VACÍO" && idDoc.length > 10) {
        directorioMentores[codigoMentor] = idDoc;
      }
    }
  }

  // 2. OBTENER DATOS MASTER
  const rangoMaster = hojaMaster.getDataRange();
  const datosMaster = rangoMaster.getValues();
  
  const COL_CODIGO_MENTOR = 1; // B
  const COL_NOMBRE_INI = 3;    // D
  const COL_CURP = 4;          // E
  const COL_SITUACION = 12;    // M
  const COL_FECHA = 13;        // N
  const COL_COMENTARIOS = 15;  // P

  let mentoresSincronizados = 0;
  let errores = [];

  // 3. CICLO POR MENTOR
  for (let codigo in directorioMentores) {
    try {
      const idExterno = directorioMentores[codigo];
      const ssExterno = SpreadsheetApp.openById(idExterno);
      const hojaExterna = ssExterno.getSheetByName("LISTA");
      
      if (!hojaExterna) {
        errores.push("Mentor " + codigo + ": No se encontró la pestaña 'LISTA'.");
        continue;
      }
      
      const datosExternos = hojaExterna.getDataRange().getValues();

      // --- FASE PULL ---
      let feedbackMap = new Map();
      if (datosExternos.length > 1) {
        for (let i = 1; i < datosExternos.length; i++) {
          let curpExt = datosExternos[i][1] ? datosExternos[i][1].toString().trim() : "";
          if (curpExt) {
            feedbackMap.set(curpExt, {
              situacion: datosExternos[i][9],
              fecha: datosExternos[i][10],
              comentarios: datosExternos[i][12]
            });
          }
        }
      }

      // --- FASE PUSH ---
      let filasParaEnviar = [];
      for (let j = 1; j < datosMaster.length; j++) {
        if (datosMaster[j][COL_CODIGO_MENTOR].toString().trim() === codigo) {
          let curpMaster = datosMaster[j][COL_CURP].toString().trim();

          if (feedbackMap.has(curpMaster)) {
            let info = feedbackMap.get(curpMaster);
            datosMaster[j][COL_SITUACION] = info.situacion;
            datosMaster[j][COL_FECHA] = info.fecha;
            datosMaster[j][COL_COMENTARIOS] = info.comentarios;
          }

          let recorte = datosMaster[j].slice(COL_NOMBRE_INI);
          filasParaEnviar.push(recorte);
        }
      }

      let mapaFechasExistentes = new Map();
      // --- ESCRIBIR EN ARCHIVO DEL MENTOR ---
      if (filasParaEnviar.length > 0) {
        if (hojaExterna.getLastRow() > 1) {
          // Preservar fechas antes de limpiar... (lógica ya existente)
          const datosViejos = hojaExterna.getRange(2, 2, hojaExterna.getLastRow() - 1, 10).getValues(); 
          for (let r = 0; r < datosViejos.length; r++) {
            let curpKey = datosViejos[r][0].toString().trim();
            let fechaVal = datosViejos[r][9];
            if (curpKey && fechaVal) mapaFechasExistentes.set(curpKey, fechaVal);
          }
          hojaExterna.getRange(2, 1, hojaExterna.getLastRow() - 1, hojaExterna.getLastColumn()).clearContent();
        }

        // --- APLICAR FORMATO TEXTO PLANO ANTES DE PEGAR ---
        const rangoDestino = hojaExterna.getRange(2, 1, filasParaEnviar.length, filasParaEnviar[0].length);
        rangoDestino.setNumberFormat("@"); // Fuerza a que 1.10 no sea 1.1
        rangoDestino.setValues(filasParaEnviar);

        // --- VALIDACIONES Y SINCRONIZACIÓN INMEDIATA ---
        const ultFilaM = hojaExterna.getLastRow();
        const datosNuevos = hojaExterna.getRange(2, 1, ultFilaM - 1, 11).getValues();
        
        for (let k = 0; k < datosNuevos.length; k++) {
          const filaA = k + 2;
          const curpA = datosNuevos[k][1].toString().trim();
          const valorI = datosNuevos[k][8];
          let valorJ = datosNuevos[k][9];
          const celdaJ = hojaExterna.getRange(filaA, 10);
          const celdaK = hojaExterna.getRange(filaA, 11);

          if (valorI === "") continue;

          let opciones = (valorI === "NVO.") ? ["Sin contactar", "Intento", "Contactado", "Aceptó Inv", "Rechazo"] : 
               (valorI === "CONT.") ? ["Sin contactar", "Reactivado", "Rechazo"] : []; 

          if (opciones.length > 0) {
            const regla = SpreadsheetApp.newDataValidation().requireValueInList(opciones, true).build();
            celdaJ.setDataValidation(regla);

            if (valorJ === "" || valorJ === null) {
              valorJ = "Sin contactar";
              let hoy = new Date();
              celdaJ.setValue(valorJ);
              celdaK.setValue(hoy).setNumberFormat("dd/mmm/yyyy");
              
              for (let m = 1; m < datosMaster.length; m++) {
                if (datosMaster[m][COL_CURP].toString().trim() === curpA) {
                  datosMaster[m][COL_SITUACION] = valorJ;
                  datosMaster[m][COL_FECHA] = hoy;
                }
              }
            } else if (mapaFechasExistentes.has(curpA)) {
              celdaK.setValue(mapaFechasExistentes.get(curpA)).setNumberFormat("dd/mmm/yyyy");
            }
          }
        }
      }
      
      mentoresSincronizados++;
      ss.toast("Sincronizando: " + codigo, "Progreso", 2);

    } catch (e) {
      errores.push("Mentor " + codigo + ": " + e.message);
      Logger.log("Mentor " + codigo + ": " + e.message)
    }
  }

  // 4. VOLCAR TODOS LOS CAMBIOS A LA HOJA MASTER
  hojaMaster.getRange(2, COL_CODIGO_MENTOR + 1, datosMaster.length - 1, 1).setNumberFormat("@");
  rangoMaster.setValues(datosMaster);

  // 5. FINALIZACIÓN
  if (errores.length > 0) {
    SpreadsheetApp.getUi().alert("Terminado con advertencias:\n" + errores.join("\n"));
  } else {
    ss.toast("✅ " + mentoresSincronizados + " mentores sincronizados correctamente.", "Finalizado");
  }
}
