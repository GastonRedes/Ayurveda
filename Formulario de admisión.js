function enEnviarFormulario(e) {

  const FORMULARIO = FormApp.getActiveForm(),
        RESPUESTAS = FORMULARIO.getResponses(),
        CANTIDAD_DE_RESPUESTAS = RESPUESTAS.length,
        ULTIMA_RESPUESTA = RESPUESTAS[CANTIDAD_DE_RESPUESTAS - 1],
        RESPUESTAS_DE_ULTIMO_FORMULARIO = ULTIMA_RESPUESTA.getItemResponses(),

        ID_DE_CARPETA_PLANTILLAS = "1-Et3khPb1kkSS7B6tCnUnfRX2qiLgoMc",
        CARPETA_PLANTILLAS = DriveApp.getFolderById(ID_DE_CARPETA_PLANTILLAS),
        
        ID_DE_CARPETA_PACIENTES = "1VSnbiN8sT-qbw9QyvVDW3T42vYCZOBHY",
        CARPETA_PACIENTES = DriveApp.getFolderById(ID_DE_CARPETA_PACIENTES),

        NUMERO_DE_PACIENTE = CANTIDAD_DE_RESPUESTAS.toString(),
        NOMBRE_DE_PACIENTE = RESPUESTAS_DE_ULTIMO_FORMULARIO[1].getResponse(),
        NOMBRE_DE_CARPETA_PACIENTE = NUMERO_DE_PACIENTE + " - " + NOMBRE_DE_PACIENTE,
        CARPETA_PACIENTE = CARPETA_PACIENTES.createFolder(NOMBRE_DE_CARPETA_PACIENTE);
        
  copiarCarpeta(CARPETA_PLANTILLAS, CARPETA_PACIENTE);
  rellenarPlantillaFormulario(RESPUESTAS_DE_ULTIMO_FORMULARIO, CARPETA_PACIENTE);
}

function copiarCarpeta(carpetaOrigen, carpetaDestino) {

  const ARCHIVOS = carpetaOrigen.getFiles(),
        CARPETAS = carpetaOrigen.getFolders();
  var   archivo, carpeta, nombre, carpetaCreada;
  
  while (ARCHIVOS.hasNext()) {

    archivo = ARCHIVOS.next();
    nombre = archivo.getName();
    archivo.makeCopy(nombre, carpetaDestino);
  }

  while (CARPETAS.hasNext()) {

    carpeta = CARPETAS.next();
    nombre = carpeta.getName();
    carpetaCreada = carpetaDestino.createFolder(nombre);
    copiarCarpeta(carpeta, carpetaCreada);
  }
}

function rellenarPlantillaFormulario(respuestasDeFormulario, carpetaPaciente) {

  const ARCHIVOS = carpetaPaciente.getFilesByName("Formulario de admisión"),
        ARCHIVO = ARCHIVOS.next(),
        ID = ARCHIVO.getId(),
        PLANTILLA_FORMULARIO = SpreadsheetApp.openById(ID),
        HOJA_FORMULARIO = PLANTILLA_FORMULARIO.getSheets()[0],
        CANTIDAD_DE_RESPUESTAS = respuestasDeFormulario.length,
        COLUMNA = 3;
  var   fila, celda, i, respuesta;

  fila = 8;
  celda = HOJA_FORMULARIO.getRange(fila, COLUMNA);
  celda.setValue(new Date());
  fila = 9;

  for (i = 0; i < CANTIDAD_DE_RESPUESTAS; i++) {
    
    respuesta = respuestasDeFormulario[i].getResponse();
    celda = HOJA_FORMULARIO.getRange(fila + i, COLUMNA);
    celda.setValue(respuesta.toString());
  }
}



