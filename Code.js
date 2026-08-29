/******************************************************
 * SISTEMA CONTROL CARGADORES ELÉCTRICOS
 * EDIFICIO TERRA 93 PH
 *
 * ARCHIVO: Code.gs
 * VERSIÓN 2026
 * Prueba de despliegue automático via GitHub Actions
 ******************************************************/

//======================================================
// CONFIGURACIÓN GENERAL
//======================================================

const CORREO_ADMIN = "edificioterra93@gmail.com";
const ID_CARPETA_DRIVE = "1zg96u9Kb1FisAi67zydqZkrAeBSuzOUb";

const NOMBRE_HOJA = "Registro_Cargadores";

const ENCABEZADOS_HOJA = [
  "Fecha", "Hora Entrada", "Punto de Carga", "Bloque Horario",
  "Apartamento", "Placa", "Nombre Residente", "Cédula",
  "Lectura Inicial (kWh)", "Lectura Final (kWh)", "Consumo (kWh)",
  "Vigilante", "Hora Salida", "Correo Residente",
  "Foto Inicial (URL)", "Foto Final (URL)",
  "Firma Vigilante (URL)", "Firma Residente (URL)", "Observaciones",
  "Vigilante Salida"
];

//======================================================
// CARGAR SISTEMA WEB
//======================================================

function doGet() {
  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Control de Cargadores Eléctricos - Terra 93 PH")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

//======================================================
// INCLUIR HTML
//======================================================

function include(nombreArchivo) {
  return HtmlService.createHtmlOutputFromFile(nombreArchivo).getContent();
}

//======================================================
// PRUEBA DE CONEXIÓN
//======================================================

function probarConexion() {
  return {
    estado: "OK",
    mensaje: "Conexión exitosa.",
    fecha: new Date()
  };
}

//======================================================
// HOJA DE REGISTRO (se crea sola si no existe)
//======================================================

function obtenerHojaRegistro() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName(NOMBRE_HOJA);

  if (!hoja) {
    hoja = libro.insertSheet(NOMBRE_HOJA);
    hoja.appendRow(ENCABEZADOS_HOJA);
    hoja.setFrozenRows(1);
  }

  return hoja;
}

//======================================================
// SUBIR ARCHIVO (FOTO O FIRMA) A GOOGLE DRIVE
//======================================================

function subirArchivoDrive(base64, apto, placa, prefijoNombre) {
  if (!base64) {
    return "";
  }

  var partes = base64.split(",");
  var cabecera = partes[0];
  var contenido = partes[1];

  var coincidencia = cabecera.match(/data:(.*);base64/);
  var mime = coincidencia ? coincidencia[1] : "image/jpeg";
  var extension = mime.split("/")[1] || "jpg";

  var carpetaPrincipal = DriveApp.getFolderById(ID_CARPETA_DRIVE);
  var nombreCarpeta = "Apto " + apto + " - " + placa;
  var carpetas = carpetaPrincipal.getFoldersByName(nombreCarpeta);

  var carpetaVehiculo = carpetas.hasNext()
    ? carpetas.next()
    : carpetaPrincipal.createFolder(nombreCarpeta);

  var marcaTiempo = Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss"
  );

  var blob = Utilities.newBlob(
    Utilities.base64Decode(contenido),
    mime,
    prefijoNombre + "_" + marcaTiempo + "." + extension
  );

  var archivo = carpetaVehiculo.createFile(blob);

  archivo.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  return archivo.getUrl();
}

//======================================================
// OBTENER ARCHIVO DESDE DRIVE (BLOB)
//======================================================

function obtenerArchivoDrive(url) {
  if (!url) {
    return null;
  }

  var id = url.match(/[-\w]{25,}/);

  if (!id) {
    return null;
  }

  try {
    return DriveApp.getFileById(id[0]).getBlob();
  } catch (error) {
    return null;
  }
}

//======================================================
// OBTENER IMAGEN DE DRIVE COMO BASE64 (para incrustar en el PDF)
//======================================================

function obtenerImagenBase64Drive(url) {
  var blob = obtenerArchivoDrive(url);

  if (!blob) {
    return "";
  }

  var mime = blob.getContentType();
  var base64 = Utilities.base64Encode(blob.getBytes());

  return "data:" + mime + ";base64," + base64;
}

//======================================================
// GUARDAR REGISTRO DE ENTRADA
//======================================================

function guardarEntrada(datos) {
  try {
    validarEntrada(datos);

    datos.apto = String(datos.apto).trim().toUpperCase();
    datos.placa = String(datos.placa).trim().toUpperCase();

    var activo = buscarVehiculoActivo(datos.apto, datos.placa);

    if (activo) {
      throw new Error(
        "Ya existe una carga activa para el apartamento " + datos.apto +
        " y placa " + datos.placa + ". Registre la salida antes de una nueva entrada."
      );
    }

    var hoja = obtenerHojaRegistro();

    var urlFotoInicial = subirArchivoDrive(
      datos.fotoInicial, datos.apto, datos.placa, "Entrada_Foto"
    );

    var urlFirmaVigilante = subirArchivoDrive(
      datos.firmaVigilante, datos.apto, datos.placa, "Entrada_FirmaVigilante"
    );

    var urlFirmaResidente = subirArchivoDrive(
      datos.firmaResidente, datos.apto, datos.placa, "Entrada_FirmaResidente"
    );

    datos.fotoInicial = urlFotoInicial;

    hoja.appendRow([
      datos.fecha,              // A
      datos.horaEntrada,        // B
      datos.puntoCarga,         // C
      datos.bloque,             // D
      datos.apto,               // E
      datos.placa,              // F
      datos.nombre,             // G
      datos.cedula,             // H
      datos.lecturaInicial,     // I
      "",                       // J Lectura final
      "",                       // K Consumo
      datos.vigilante,          // L
      "",                       // M Hora salida
      datos.correoResidente,    // N
      urlFotoInicial,           // O
      "",                       // P Foto final
      urlFirmaVigilante,        // Q
      urlFirmaResidente,        // R
      datos.observaciones || "" // S
    ]);

    enviarCorreoInicio(datos);

    return {
      success: true,
      mensaje: "Entrada registrada correctamente."
    };

  } catch (error) {
    return {
      success: false,
      mensaje: error.message
    };
  }
}

//======================================================
// OBTENER VEHÍCULOS EN CARGA
//======================================================

function obtenerVehiculosEnCarga() {
  var hoja = obtenerHojaRegistro();
  var datos = hoja.getDataRange().getValues();
  var lista = [];

  for (var i = 1; i < datos.length; i++) {
    // Columna M (Hora salida) vacía = sigue cargando
    if (datos[i][12] === "") {
      lista.push({
        fila: i + 1,
        apto: datos[i][4],
        placa: datos[i][5],
        nombre: datos[i][6]
      });
    }
  }

  return lista;
}

//======================================================
// REGISTRAR SALIDA
//======================================================

function registrarSalida(datos) {
  try {
    if (!datos.fila || isNaN(Number(datos.fila))) {
      throw new Error("Seleccione un vehículo válido.");
    }

    if (datos.lecturaFinal === "" || datos.lecturaFinal === undefined || isNaN(Number(datos.lecturaFinal))) {
      throw new Error("Ingrese una lectura final válida.");
    }

    if (!datos.horaSalida) {
      throw new Error("Ingrese la hora de salida.");
    }

    if (!datos.vigilanteSalida) {
      throw new Error("El nombre del vigilante de salida es obligatorio.");
    }

    var hoja = obtenerHojaRegistro();
    var fila = Number(datos.fila);

    var registro = obtenerRegistro(fila);

    if (registro.horaSalida) {
      throw new Error("Este registro ya tiene una salida registrada.");
    }

    var lecturaInicial = Number(registro.lecturaInicial);
    var lecturaFinal = Number(datos.lecturaFinal);

    if (lecturaFinal < lecturaInicial) {
      throw new Error("La lectura final no puede ser menor a la inicial.");
    }

    var consumo = lecturaFinal - lecturaInicial;

    var urlFotoFinal = subirArchivoDrive(
      datos.fotoFinal, registro.apto, registro.placa, "Salida_Foto"
    );

    // J Lectura Final
    hoja.getRange(fila, 10).setValue(lecturaFinal);

    // K Consumo
    hoja.getRange(fila, 11).setValue(consumo);

    // M Hora Salida
    hoja.getRange(fila, 13).setValue(datos.horaSalida);

    // P URL Foto Final
    hoja.getRange(fila, 16).setValue(urlFotoFinal);

    // T Vigilante Salida
    hoja.getRange(fila, 20).setValue(datos.vigilanteSalida);

    registro.lecturaFinal = lecturaFinal;
    registro.consumo = consumo;
    registro.horaSalida = datos.horaSalida;
    registro.fotoFinal = urlFotoFinal;
    registro.vigilanteSalida = datos.vigilanteSalida;

    // Recuperar las firmas originales (tomadas en la entrada) para el PDF de salida
    registro.firmaVigilante = obtenerImagenBase64Drive(registro.firmaVigilanteURL);
    registro.firmaResidente = obtenerImagenBase64Drive(registro.firmaResidenteURL);

    enviarCorreoSalida(registro);

    return {
      success: true,
      consumo: consumo,
      mensaje: "Salida registrada correctamente."
    };

  } catch (error) {
    return {
      success: false,
      mensaje: error.message
    };
  }
}

//======================================================
// BUSCAR VEHÍCULO ACTIVO (misma apto + placa, sin hora de salida)
//======================================================

function buscarVehiculoActivo(apto, placa) {
  var hoja = obtenerHojaRegistro();
  var datos = hoja.getDataRange().getValues();

  for (var i = 1; i < datos.length; i++) {
    var mismoApto = String(datos[i][4]).trim().toUpperCase() === String(apto).trim().toUpperCase();
    var mismaPlaca = String(datos[i][5]).trim().toUpperCase() === String(placa).trim().toUpperCase();
    var sinSalida = datos[i][12] === "" || datos[i][12] === null;

    if (mismoApto && mismaPlaca && sinSalida) {
      return { fila: i + 1, datos: datos[i] };
    }
  }

  return null;
}

//======================================================
// BUSCAR VEHÍCULO (cualquier coincidencia apto + placa)
//======================================================

function buscarVehiculo(apto, placa) {
  var hoja = obtenerHojaRegistro();
  var datos = hoja.getDataRange().getValues();

  for (var i = 1; i < datos.length; i++) {
    if (
      String(datos[i][4]).trim().toUpperCase() === String(apto).trim().toUpperCase() &&
      String(datos[i][5]).trim().toUpperCase() === String(placa).trim().toUpperCase()
    ) {
      return { fila: i + 1, datos: datos[i] };
    }
  }

  return null;
}

//======================================================
// OBTENER DATOS DEL REGISTRO
//======================================================

function obtenerRegistro(fila) {
  var hoja = obtenerHojaRegistro();
  var datos = hoja.getRange(fila, 1, 1, 20).getValues()[0];

  return {
    fecha: datos[0],
    horaEntrada: datos[1],
    puntoCarga: datos[2],
    bloque: datos[3],
    apto: datos[4],
    placa: datos[5],
    nombre: datos[6],
    cedula: datos[7],
    lecturaInicial: datos[8],
    lecturaFinal: datos[9],
    consumo: datos[10],
    vigilante: datos[11],
    horaSalida: datos[12],
    correoResidente: datos[13] || "",
    fotoInicial: datos[14],
    fotoFinal: datos[15],
    firmaVigilanteURL: datos[16],
    firmaResidenteURL: datos[17],
    observaciones: datos[18] || "",
    vigilanteSalida: datos[19] || ""
  };
}

//======================================================
// ENVIAR CORREO INICIO DE CARGA
//======================================================

function enviarCorreoInicio(datos) {
  var pdf = crearPDFRegistro("Inicio de carga", datos);
  var archivos = [pdf];

  var foto = obtenerArchivoDrive(datos.fotoInicial);

  if (foto) {
    archivos.push(foto);
  }

  MailApp.sendEmail({
    to: (datos.correoResidente ? datos.correoResidente + "," : "") + CORREO_ADMIN,

    subject: "Inicio carga vehículo - Apto " + datos.apto + " - " + datos.placa,

    htmlBody:
      "<h2>Control Cargadores Terra 93 PH</h2>" +
      "<p>Se confirma inicio de carga eléctrica.</p><br>" +
      "<b>Apartamento:</b> " + datos.apto + "<br>" +
      "<b>Placa:</b> " + datos.placa + "<br>" +
      "<b>Residente:</b> " + datos.nombre + "<br>" +
      "<b>Fecha:</b> " + datos.fecha + "<br>" +
      "<b>Hora inicio:</b> " + datos.horaEntrada + "<br>" +
      "<b>Lectura inicial:</b> " + datos.lecturaInicial + " kWh",

    attachments: archivos
  });
}

//======================================================
// ENVIAR CORREO FINALIZACIÓN
//======================================================

function enviarCorreoSalida(datos) {
  var pdf = crearPDFRegistro("Finalización de carga", datos);
  var archivos = [pdf];

  var foto = obtenerArchivoDrive(datos.fotoFinal);

  if (foto) {
    archivos.push(foto);
  }

  MailApp.sendEmail({
    to: (datos.correoResidente ? datos.correoResidente + "," : "") + CORREO_ADMIN,

    subject: "Finalización carga vehículo - Apto " + datos.apto + " - " + datos.placa,

    htmlBody:
      "<h2>Control Cargadores Terra 93 PH</h2>" +
      "<p>Se confirma la finalización de la carga eléctrica.</p><br>" +
      "<b>Apartamento:</b> " + datos.apto + "<br>" +
      "<b>Placa:</b> " + datos.placa + "<br>" +
      "<b>Residente:</b> " + datos.nombre + "<br>" +
      "<b>Hora salida:</b> " + datos.horaSalida + "<br>" +
      "<b>Lectura final:</b> " + datos.lecturaFinal + " kWh<br>" +
      "<b>Consumo:</b> " + datos.consumo + " kWh<br>" +
      "<b>Vigilante salida:</b> " + datos.vigilanteSalida,

    attachments: archivos
  });
}

//======================================================
// CREAR PDF CONFIRMACIÓN
//======================================================

function crearPDFRegistro(tipo, datos) {
  var html =
    "<html><body style='font-family:Arial'>" +
    "<h1 style='text-align:center'>Terra 93 PH</h1>" +
    "<h2 style='text-align:center'>" + tipo + "</h2><hr>" +

    "<h3>Datos del vehículo</h3>" +
    "<p><b>Apartamento:</b> " + datos.apto + "</p>" +
    "<p><b>Placa:</b> " + datos.placa + "</p>" +
    "<p><b>Residente:</b> " + datos.nombre + "</p>" +
    "<p><b>Cédula:</b> " + datos.cedula + "</p>" +

    "<h3>Información de carga</h3>" +
    "<p><b>Fecha:</b> " + datos.fecha + "</p>" +
    "<p><b>Hora entrada:</b> " + datos.horaEntrada + "</p>" +
    "<p><b>Punto carga:</b> " + datos.puntoCarga + "</p>" +
    "<p><b>Bloque:</b> " + datos.bloque + "</p>" +
    "<p><b>Lectura inicial:</b> " + datos.lecturaInicial + " kWh</p>";

  if (datos.lecturaFinal) {
    html +=
      "<p><b>Hora salida:</b> " + datos.horaSalida + "</p>" +
      "<p><b>Lectura final:</b> " + datos.lecturaFinal + " kWh</p>" +
      "<p><b>Consumo:</b> " + datos.consumo + " kWh</p>";
  }

  if (datos.observaciones) {
    html += "<h3>Observaciones</h3><p>" + datos.observaciones + "</p>";
  }

  html +=
    "<h3>Responsables</h3>" +
    "<p><b>Vigilante:</b> " + datos.vigilante + "</p>" +
    (datos.vigilanteSalida ? "<p><b>Vigilante salida:</b> " + datos.vigilanteSalida + "</p>" : "") +

    "<h3>Firmas</h3>" +
    "<p>Firma Vigilante:</p>" +
    (datos.firmaVigilante ? "<img src='" + datos.firmaVigilante + "' width='250'>" : "Sin firma") +

    "<p>Firma Residente:</p>" +
    (datos.firmaResidente ? "<img src='" + datos.firmaResidente + "' width='250'>" : "Sin firma") +

    "<br><br><p>Documento generado automáticamente por el sistema Terra 93 PH.</p>" +
    "</body></html>";

  var blob = Utilities.newBlob(html, "text/html", "comprobante.html");
  var pdf = blob.getAs("application/pdf");

  pdf.setName(tipo + " - Apto " + datos.apto + ".pdf");

  return pdf;
}

//======================================================
// VALIDAR CORREOS
//======================================================

function validarCorreo(correo) {
  if (!correo) {
    return false;
  }

  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo);
}

//======================================================
// VALIDAR DATOS DE ENTRADA
//======================================================

function validarEntrada(datos) {
  if (!datos.apto) {
    throw new Error("El apartamento es obligatorio.");
  }

  if (!datos.placa) {
    throw new Error("La placa es obligatoria.");
  }

  if (!datos.nombre) {
    throw new Error("El nombre del residente es obligatorio.");
  }

  if (!datos.cedula) {
    throw new Error("La cédula del residente es obligatoria.");
  }

  if (!datos.vigilante) {
    throw new Error("El nombre del vigilante es obligatorio.");
  }

  if (!datos.correoResidente) {
    throw new Error("El correo del residente es obligatorio.");
  }

  if (!validarCorreo(datos.correoResidente)) {
    throw new Error("El correo del residente no es válido.");
  }

  if (datos.lecturaInicial === "" || datos.lecturaInicial === undefined || isNaN(Number(datos.lecturaInicial))) {
    throw new Error("La lectura inicial debe ser un número válido.");
  }

  return true;
}

//======================================================
// PRUEBA DE ENVÍO DE CORREO
//======================================================

function probarCorreo() {
  MailApp.sendEmail({
    to: CORREO_ADMIN,
    subject: "Prueba sistema cargadores Terra 93",
    htmlBody: "<h2>Sistema funcionando</h2><p>Prueba de correo exitosa.</p>"
  });

  return "Correo enviado correctamente";
}

