/**
 * Consolida registros en la pestaña FORMS con indicación de procedencia.
 */
function consolidarRegistrosConFuente() {
  const ID_FORM_1 = "INSERTA_ID_AQUI"; //ACTUALIZACION DE DATOS 2026
  const ID_FORM_2 = "INSERTA_ID_AQUI"; //NUEVOS MIEMBROS 2026
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaDestino = ss.getSheetByName("FORMS");
  
  if (!hojaDestino) {
    console.error("No se encontró la pestaña FORMS");
    return;
  }

  // 1. OBTENER DATOS DE FORMULARIOS CON ETIQUETA DE FUENTE
  const data1 = obtenerDiccionarioForm(ID_FORM_1, "ACT. DATOS");
  const data2 = obtenerDiccionarioForm(ID_FORM_2, "NVOS. MB");

  // Unificar: Si alguien está en ambos, prevalece el Form 1 (ACT. DATOS)
  const todosLosRegistros = { ...data2, ...data1 };
  
  // 2. LEER DATOS ACTUALES PARA EVITAR DUPLICADOS
  const rangoActual = hojaDestino.getDataRange();
  const datosActuales = rangoActual.getValues();
  const mapaExistentes = new Set();
  
  // Mapeamos CURPs existentes (Columna B = índice 1)
  for (let i = 1; i < datosActuales.length; i++) {
    const curpExistente = limpiar(datosActuales[i][1]);
    if (curpExistente) mapaExistentes.add(curpExistente);
  }

  const nuevasFilas = [];

  // 3. FILTRAR NUEVOS REGISTROS
  Object.keys(todosLosRegistros).forEach(curp => {
    if (!mapaExistentes.has(curp) && curp !== "") {
      const d = todosLosRegistros[curp];
      nuevasFilas.push([
        d.nombre, d.curp, d.tel, d.correo, d.grado, d.muni, d.cct, d.escuela, d.nivel, d.fuente
      ]);
    }
  });

  // 4. INSERTAR Y MARCAR FUENTE PARA REGISTROS MANUALES
  if (nuevasFilas.length > 0) {
    hojaDestino.getRange(hojaDestino.getLastRow() + 1, 1, nuevasFilas.length, 10).setValues(nuevasFilas);
  }

  // Llenar celdas vacías en la columna J (Fuente) como "MANUAL / PREVIO"
  const ultimaFilaPost = hojaDestino.getLastRow();
  if (ultimaFilaPost > 1) {
    const rangoFuente = hojaDestino.getRange(2, 10, ultimaFilaPost - 1, 1);
    const valoresFuente = rangoFuente.getValues();
    let huboCambioManual = false;

    for (let k = 0; k < valoresFuente.length; k++) {
      if (valoresFuente[k][0] === "") {
        valoresFuente[k][0] = "MANUAL / PREVIO";
        huboCambioManual = true;
      }
    }
    if (huboCambioManual) rangoFuente.setValues(valoresFuente);
  }

  // 5. VALIDACIÓN VISUAL (Grado vs Nivel)
  validarGradoNivel(hojaDestino);
}

/**
 * Mapea columnas y asigna la etiqueta de fuente
 */
function obtenerDiccionarioForm(id, etiquetaFuente) {
  const dict = {};
  try {
    const valores = SpreadsheetApp.openById(id).getSheets()[0].getDataRange().getValues();
    valores.shift();
    
    valores.forEach(f => {
      const curp = limpiar(f[10]); // Columna K
      if (curp) {
        dict[curp] = {
          nombre: f[6], curp: curp, tel: f[20], correo: f[21],
          grado: f[19], muni: f[17], cct: f[13], escuela: f[16], nivel: f[18],
          fuente: etiquetaFuente
        };
      }
    });
  } catch(e) { console.log("Error en " + etiquetaFuente + ": " + id); }
  return dict;
}

/**
 * Valida discrepancias y pinta de amarillo brillante (#FFFF00)
 */
function validarGradoNivel(hoja) {
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const rangoTotal = hoja.getRange(2, 1, ultimaFila - 1, 10);
  const datos = rangoTotal.getValues();
  rangoTotal.setBackground(null); // Limpiar
  
  const colores = [];
  for (let i = 0; i < datos.length; i++) {
    const grado = limpiar(datos[i][4]); // Columna E
    const nivel = limpiar(datos[i][8]); // Columna I
    
    if (grado !== "" && nivel !== "" && grado.indexOf(nivel) === -1) {
      colores.push(Array(10).fill("#FFFF00")); 
    } else {
      colores.push(Array(10).fill(null));
    }
  }
  rangoTotal.setBackgrounds(colores);
}

function validarInscritosDobleCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shForms = ss.getSheetByName("FORMS");
  const shLista = ss.getSheetByName("LISTA_TOTAL");
  
  // --- 1. CARGAR DATOS DE REFERENCIA (LISTA_TOTAL) ---
  // D: Nombre (Col 4), E: CURP (Col 5)
  const lastRowLista = shLista.getLastRow();
  const dataLista = shLista.getRange("D2:E" + lastRowLista).getValues();
  
  // Creamos dos mapas de referencia para búsquedas instantáneas
  const setCurpsRef = new Set();
  const setNombresRef = new Set();
  
  dataLista.forEach(fila => {
    const nombre = fila[0].toString().trim();
    const curp = fila[1].toString().trim().toUpperCase();
    
    if (curp) setCurpsRef.add(curp);
    if (nombre) {
      // Tokenizamos el nombre de referencia (ordenar palabras)
      const nombreToken = nombre.split(/\s+/).sort().join(" ");
      setNombresRef.add(nombreToken);
    }
  });

  // --- 2. CARGAR DATOS DE INSCRITOS (FORMS) ---
  // A: Nombre (Col 1), B: CURP (Col 2)
  const lastRowForms = shForms.getLastRow();
  if (lastRowForms < 2) return; // Si no hay datos, salir
  
  const dataInscritos = shForms.getRange("A2:B" + lastRowForms).getValues();
  
  // --- 3. PROCESAR Y COMPARAR ---
  const resultados = dataInscritos.map(fila => {
    const nombreInscrito = fila[0].toString().trim();
    const curpInscrito = fila[1].toString().trim().toUpperCase();
    
    // Paso 1: Validar por CURP (Prioridad máxima)
    if (curpInscrito && setCurpsRef.has(curpInscrito)) {
      return ["✅ OK (CURP)"];
    }
    
    // Paso 2: Si falla el CURP, intentar por Nombre (Tokenizado)
    if (nombreInscrito) {
      const nombreInscritoToken = nombreInscrito.split(/\s+/).sort().join(" ");
      if (setNombresRef.has(nombreInscritoToken)) {
        return ["⚠️ OK (Nombre, revisar CURP)"];
      }
    }
    
    // Paso 3: Si nada coincide
    return ["❌ No encontrado"];
  });

  // --- 4. ESCRIBIR RESULTADOS ---
  // Escribiremos en la Columna K de la hoja FORMS
  shForms.getRange(2, 11, resultados.length, 1).setValues(resultados);
  
  // Solo mostrar UI si la ejecución es manual
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Validación completada con éxito.");
  } catch (e) {
    // Si falla (porque es un trigger de tiempo), se registra en el log en lugar de lanzar error
    console.log("Ejecución finalizada mediante activador automático.");
  }
}

function actualizaFormsCompleto(){
  consolidarRegistrosConFuente();
  validarInscritosDobleCheck();
}
