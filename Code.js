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

    // La firma del vigilante ya no se sube a Drive: se guarda tal cual
    // (base64) directo en la celda, ahorrando una llamada a Drive por registro.
    var urlFirmaResidente = subirArchivoDrive(
      datos.firmaResidente, datos.apto, datos.placa, "Entrada_FirmaResidente"
    );

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
      datos.firmaVigilante,     // Q (base64 directo, sin Drive)
      urlFirmaResidente,        // R
      datos.observaciones || "" // S
    ]);

    datos.fotoInicial = urlFotoInicial;

    // El PDF y el correo de confirmación se generan en segundo plano
    // (vía trigger), para no hacer esperar al usuario en el formulario.
    encolarTareaEnSegundoPlano("entrada", datos);

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

    if (!datos.fotoFinal) {
      throw new Error("La foto de la lectura final es obligatoria.");
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

    // El PDF y el correo de finalización se generan en segundo plano
    // (vía trigger). Las firmas (vigilante en base64 directo, residente
    // vía Drive) se recuperan dentro de enviarCorreoSalida en ese momento,
    // no aquí, para no hacer esperar al usuario en el formulario.
    encolarTareaEnSegundoPlano("salida", registro);

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
  // La firma del vigilante ya viene en base64 directo desde la celda
  // (columna "Firma Vigilante (URL)"); la del residente sigue en Drive.
  datos.firmaVigilante = datos.firmaVigilanteURL;
  datos.firmaResidente = obtenerImagenBase64Drive(datos.firmaResidenteURL);

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
// TAREAS EN SEGUNDO PLANO (PDF + correo, fuera de la
// respuesta síncrona a google.script.run)
//======================================================

var NOMBRE_PROP_COLA_TAREAS = "colaTareasPendientes";
var NOMBRE_FUNCION_TAREAS = "procesarTareasEnSegundoPlano";
var CACHE_TTL_TAREA_SEGUNDOS = 21600; // 6 horas (máximo permitido por CacheService)
var MAX_INTENTOS_CORREO = 5;

function encolarTareaEnSegundoPlano(tipo, datos) {
  encolarTareaConReintento(tipo, datos, 1);
}

function encolarTareaConReintento(tipo, datos, intento) {
  var id = Utilities.getUuid();

  CacheService.getScriptCache().put(
    "tarea_" + id,
    JSON.stringify({ tipo: tipo, datos: datos, intento: intento }),
    CACHE_TTL_TAREA_SEGUNDOS
  );

  var props = PropertiesService.getScriptProperties();
  var cola = JSON.parse(props.getProperty(NOMBRE_PROP_COLA_TAREAS) || "[]");
  cola.push(id);
  props.setProperty(NOMBRE_PROP_COLA_TAREAS, JSON.stringify(cola));

  // Primer intento: casi inmediato. Reintentos: espera creciente (30s, 1, 2,
  // 4 min...) con tope de 15 min, para darle tiempo a fallas transitorias
  // (cuota de correo, timeout de red) a resolverse solas.
  var retrasoMs = intento <= 1
    ? 1000
    : Math.min(30000 * Math.pow(2, intento - 2), 15 * 60 * 1000);

  ScriptApp.newTrigger(NOMBRE_FUNCION_TAREAS)
    .timeBased()
    .after(retrasoMs)
    .create();
}

function procesarTareasEnSegundoPlano() {
  eliminarTriggersDe(NOMBRE_FUNCION_TAREAS);

  var props = PropertiesService.getScriptProperties();
  var cache = CacheService.getScriptCache();
  var cola = JSON.parse(props.getProperty(NOMBRE_PROP_COLA_TAREAS) || "[]");

  if (cola.length === 0) {
    return;
  }

  props.deleteProperty(NOMBRE_PROP_COLA_TAREAS);

  cola.forEach(function (id) {
    var raw = cache.get("tarea_" + id);

    if (!raw) {
      // No debería pasar (TTL de 6h), pero si el dato ya expiró no hay
      // forma de recuperarlo: no se puede reintentar algo que ya no existe.
      return;
    }

    cache.remove("tarea_" + id);

    var tarea = JSON.parse(raw);

    try {
      if (tarea.tipo === "entrada") {
        enviarCorreoInicio(tarea.datos);
      } else if (tarea.tipo === "salida") {
        enviarCorreoSalida(tarea.datos);
      }
    } catch (error) {
      manejarFalloEnvioCorreo(tarea, error);
    }
  });
}

//======================================================
// REINTENTOS Y ALERTA DE RESPALDO SI EL CORREO SIGUE FALLANDO
//======================================================

function manejarFalloEnvioCorreo(tarea, error) {
  var intentoActual = tarea.intento || 1;

  if (intentoActual < MAX_INTENTOS_CORREO) {
    encolarTareaConReintento(tarea.tipo, tarea.datos, intentoActual + 1);
    return;
  }

  // Se agotaron los reintentos: se manda una alerta simple (sin PDF ni
  // fotos adjuntas, para que no dependa de lo mismo que pudo estar
  // fallando) al administrador, avisando que hay que revisar y reenviar
  // el comprobante manualmente. El registro en la hoja nunca se pierde.
  try {
    MailApp.sendEmail({
      to: CORREO_ADMIN,
      subject: "FALLÓ el envío del comprobante - Apto " + tarea.datos.apto + " - " + tarea.datos.placa,
      htmlBody:
        "<h2>Aviso: no se pudo enviar el correo de confirmación</h2>" +
        "<p>Tipo: " + (tarea.tipo === "entrada" ? "Inicio de carga" : "Finalización de carga") + "</p>" +
        "<p><b>Apartamento:</b> " + tarea.datos.apto + "</p>" +
        "<p><b>Placa:</b> " + tarea.datos.placa + "</p>" +
        "<p>Se intentó " + MAX_INTENTOS_CORREO + " veces sin éxito. El registro SÍ quedó " +
        "guardado correctamente en la hoja de cálculo; solo falló el envío del comprobante " +
        "por correo. Revisar manualmente y reenviarlo si es necesario.</p>" +
        "<p><b>Último error:</b> " + (error && error.message ? error.message : error) + "</p>"
    });
  } catch (errorAlerta) {
    // Si ni siquiera esta alerta (la versión más simple posible del correo)
    // se puede enviar, es una falla más de fondo (ej. cuota diaria de
    // MailApp agotada) que no se puede resolver reintentando desde aquí.
  }
}

function eliminarTriggersDe(nombreFuncion) {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === nombreFuncion) {
      ScriptApp.deleteTrigger(trigger);
    }
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

  if (!datos.fotoInicial) {
    throw new Error("La foto de la lectura inicial es obligatoria.");
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

