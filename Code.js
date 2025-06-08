// Función que permite conectarse a una hoja de google sheet 
function conectarHoja(sheetName) {
  var spreadsheetId = '1-sTiLcNB8V3N3vYsSsqYoKYt7jXYqo0Ymk2_l-b-8QU'; // Reemplaza con el id de la hoja 
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  console.log("Si se suben los datos desde el escritorio");

  if (sheetName) {
    return spreadsheet.getSheetByName(sheetName);
  } else {
    return spreadsheet.getActiveSheet(); // Retorna la pestaña activa si no se especifica un nombre
  }
}

function obtenerDatos(sheetName) {
  var sheet = conectarHoja(sheetName);
  var data = sheet.getDataRange().getValues();
  return data;
}

function obtenerDatosLogistica(sheetName) {
  var sheet = conectarHoja(sheetName);
  var data = sheet.getDataRange().getValues();
  return data;
}

function conectarHojaLogistica(sheetName) {
  var spreadsheetId = '1D5TH7LVa0MFcoHql-mE00msXfU4WTctjcUFk02wxyYI'; // Reemplaza con el id de la hoja 
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  if (sheetName) {
    return spreadsheet.getSheetByName(sheetName);
  } else {
    return spreadsheet.getActiveSheet(); // Retorna la pestaña activa si no se especifica un nombre
  }
}

// Incluir archivo HTML
function include(fileName) {
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

// Función que se activa cada vez que el responsable aprueba o desaprueba una solicitud
function doGet(e) {

  // obteniendo el id de la solicitud
  var filterData = obtenerUltimosRegistros(e.parameter.solicitudId);

  if (filterData.length > 0) {
    //Obtener la primera fila 
    var primeraFila = filterData[0];
    var estado = e.parameter.estado;

    // Lo está aprobando el jefe o gerente de área 
    if (e.parameter.numberAprobadores == 1) {
      if (primeraFila[15] == "Aprobado") {
        return ContentService.createTextOutput(
          "LA SOLICITUD NO PUEDE SER " + e.parameter.estado + " PORQUE YA FUÉ APROBADA."
        );
      } else if (primeraFila[15] == "Desaprobado") {
        return ContentService.createTextOutput(
          "LA SOLICITUD NO PUEDE SER " + e.parameter.estado + " PORQUE YA FUÉ DESAPROBADA."
        );
      }

    } // Lo está aprobando el gerente general
    else if (e.parameter.numberAprobadores == 2) {
      if (primeraFila[18] == "Aprobado") {
        return ContentService.createTextOutput(
          "LA SOLICITUD NO PUEDE SER " + e.parameter.estado + " PORQUE YA FUÉ APROBADA."
        );
      } else if (primeraFila[18] == "Desaprobado") {
        return ContentService.createTextOutput(
          "LA SOLICITUD NO PUEDE SER " + e.parameter.estado + " PORQUE YA FUÉ DESAPROBADA."
        );
      }
    }
  }

  if (
    (e.parameter.solicitudId && e.parameter.estado && e.parameter.aprobadoresEmail)
  ) {

    return handleEstadoRequest(e, filterData[0][21]);

  }

  var template = HtmlService.createTemplateFromFile("index");
  //template.pubUrl = ScriptApp.getService().getUrl(); // Obtener la URL del script dinámicamente
  return template.evaluate();
}

//Verifica que el usuario está autorizado o no a aprobar la solicitud. 
function isUserAuthorized(email) {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios_Autorizados');
  //var data = sheet.getDataRange().getValues();
  var data = obtenerDatos("users");
  console.log("Listado de usuarios isUserAuthorized: " + data);
  return data.some(row => row[2].toLowerCase() === email.toLowerCase());
}


// Función para verificar las credenciales del usuario
function checkUserCredentials(username, password) {

  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("users");
  //var data = sheet.getDataRange().getValues();
  var data = obtenerDatos("users");

  // Transformar el nombre de usuario a minúsculas y eliminar espacios
  username = username.trim().toLowerCase();
  password = password.trim();

  // Hashear la contraseña ingresada por el usuario
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  const hashedPassword = digest.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');


  //Cargar las solicitudes de compra hechas por el usuario que se está logueando
  for (var i = 1; i < data.length; i++) {

    // Verificar si las columnas de username y password contienen datos antes de convertirlos
    var dbUsername = (data[i][8] !== undefined && data[i][8] !== null) ? String(data[i][8]).trim() : "";
    var dbPassword = (data[i][9] !== undefined && data[i][9] !== null) ? String(data[i][9]).trim() : "";

    // Si el username o la password están vacíos, omitir esta fila
    if (dbUsername === "" || dbPassword === "") {
      continue;
    }

    if (dbUsername === username && dbPassword === hashedPassword) {

      if (data[i][6] == 0) {
        return {
          isValid: false,
          message: "El usuario no tiene permiso. Por favor, contacte al area de IT."
        };
      }
      return {
        isValid: true,
        user: data[i]
      };
    }
  }
  return { isValid: false, message: "Credenciales incorrectas" };
}

function getUserEmail() {
  var email = Session.getActiveUser().getEmail();
  var email11 = Session.getEffectiveUser().getEmail();
  var email22 = email = Session.getUser().getEmail();

  Logger.log("Email: " + email);
  Logger.log("Email 11: " + email11);
  Logger.log("Email 22: " + email22);

  if (!email) {
    email = Session.getEffectiveUser().getEmail();
    Logger.log("Email 2: " + email);
  }
  if (!email) {
    email = Session.getUser().getEmail();
    Logger.log("Email 3: " + email);
  }
  return { email: email };
}

function checkUserEmail(email) {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("users");
  //var data = sheet.getDataRange().getValues();

  var data = obtenerDatos("users");

  email = email.trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var dbEmail = String(data[i][2]).trim().toLowerCase(); // Asumiendo que el email está en la columna 3

    if (dbEmail === email) {
      //Verificar que el usuario esté habilitado para realizar solicitudes de compra
      if (data[i][6] == 0) {
        return {
          isValid: false,
          message: "El usuario no tiene permiso. Por favor, contacte al area de IT."
        };
      }
      return {
        isValid: true,
        user: data[i]
      };
    }
  }
  return { isValid: false, message: "Correo electrónico no registrado" };
}

function formatSolicitudesPorEmail(email) {
  var solicitudesPorUsuario = cargarSolicitudesUsuarioPorEmail(email);
  var formatSolicitudes = adaptarSolicitudesUsuario(solicitudesPorUsuario);
  return JSON.stringify(formatSolicitudes);
}

// Función para cargar las solicitudes hechas por el usuario logueado 
function cargarSolicitudesUsuarioPorEmail(email) {

  var data = obtenerDatos("index");
  var solicitudes = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][2] === email) {
      solicitudes.push(data[i]);
    }
  }

  // Ordenar las solicitudes de manera descendente por fecha
  solicitudes.sort(function (a, b) {
    // Asumiendo que la fecha está en la columna 5
    var dateA = new Date(a[4]);
    var dateB = new Date(b[4]);
    return dateB - dateA; // Orden descendente
  });

  return solicitudes;
}

// Función para dar formato a las solicitudes hechas por el usuario logeado, adaptando a la estructura de la tabla 
function adaptarSolicitudesUsuario(solicitudes) {
  var solicitudesAdaptadas = [];

  solicitudes.forEach((solicitud) => {
    var id = solicitud[0];
    var existente = solicitudesAdaptadas.find((item) => item.id === id);

    if (!existente) {
      var nuevaSolicitud = {
        id: id,
        razonCompra: solicitud[3],
        fechaRegistro: solicitud[4],
        justificacion: solicitud[6],
        totalCompra: parseFloat(solicitud[13]) || 0,
        estadoJefe: solicitud[15],
        estadoGerente: solicitud[18] ? solicitud[18] : "x",
        estado: ""
      };
      solicitudesAdaptadas.push(nuevaSolicitud);
    } else {
      existente.totalCompra += parseFloat(solicitud[13]) || 0;
    }
  });

  // Determinar el estado de cada solicitud
  solicitudesAdaptadas.forEach((solicitud) => {
    solicitud.estado = determinarEstadoSolicitud(solicitud.totalCompra, solicitud.estadoJefe, solicitud.estadoGerente);
  });
  return solicitudesAdaptadas;
}


//Función para determinar el estado de la solicitud teniendo en cuenta los campos de estado que se manejan
function determinarEstadoSolicitud(totalCompra, columnaEstadoJefe, columnaEstadoGerente) {

  if (totalCompra <= 500 && columnaEstadoJefe == "Pendiente") {
    return "Pendiente";

  } else if (totalCompra <= 500 && columnaEstadoJefe == "Aprobado") {
    return "Aprobado";

  } else if (totalCompra <= 500 && columnaEstadoJefe == "Desaprobado") {
    return "Desaprobado";

  } else if (totalCompra > 500 && columnaEstadoGerente == "Pendiente") {
    return "Pendiente";

  } else if (totalCompra > 500 && columnaEstadoGerente == "Aprobado") {
    return "Aprobado";

  } else if (totalCompra > 500 && columnaEstadoGerente == "Desaprobado") {
    return "Desaprobado";
  } else {
    return "Desconocido";
  }

}

function uploadFiles(form) {

  var sheetRegistro = conectarHoja("index");
  console.log("Hoja de registro: " + sheetRegistro);

  // Verificar si el usuario está logueado
  var userId = form.loggedUser ? form.loggedUser : null;
  if (!userId) {
    return "ERROR: Usuario no logueado.";
  }

  var solicitantes = cargarDataUsers(userId);
  console.log("Solicitantes en la función cargarDataUsers: " + solicitantes);
  if (!solicitantes) {
    return "ERROR: No se encontró al usuario en la base de datos.";
  }

  var formatSolicitante = transformarData(solicitantes);
  var solicitudId = generarSolicitudId(sheetRegistro);

  console.log("Solicitud ID en la función formatSolicitante: " + formatSolicitante);
  console.log("LLama a la función generarSolicitudId : " + solicitudId);

  var cotizacionUrl = "";
  if (form.file && form.file.length > 0) {
    try {
      cotizacionUrl = guardaCotizacionEnDrive(form, solicitudId);
      console.log("LLama a la función guardaCotizacionEnDrive : " + cotizacionUrl);
    } catch (error) {
      return "ERROR: No se pudo guardar la cotización. " + error.message;
    }
  }

  var retornoDeResitrarProductos;
  try {
    retornoDeResitrarProductos = registrarProductos(form, solicitudId, sheetRegistro, cotizacionUrl, formatSolicitante);
    console.log("LLama a la función registrarProductos : " + retornoDeResitrarProductos.totalCompra);
  } catch (error) {
    return "ERROR: No se pudieron registrar los productos. " + error.message;
  }

  Utilities.sleep(2000);
  console.log("");

  try {
    var enviarMail = enviarEmail(retornoDeResitrarProductos, solicitudId, formatSolicitante, false);
    console.log("LLama a la función enviarEmail : " + enviarMail);
  } catch (error) {
    return "ERROR: No se pudo enviar el email. " + error.message;
  }

  console.log("Tu solicitud de compra se envió correctamente");
  return "Tu solicitud de compra se envió correctamente";
}


//Función para guardar la cotización en el drive
function guardaCotizacionEnDrive(form, solicitudId) {
  const folder = DriveApp.getFolderById('1LmCAtbwu0BAm4DX9tehB1jMkOlJqReUO');

  if (form.file == null || form.file == undefined || form.file == "") {
    return "";

  } else {
    const fileBlob = form.file;

    // Obtén el nombre original del archivo
    const originalFileName = fileBlob.getName();

    // Genera el nuevo nombre de archivo
    const newFileName = `Coti_${solicitudId}_${originalFileName}`;

    // Crea el archivo en la carpeta con el nuevo nombre
    const file = folder.createFile(fileBlob).setName(newFileName);

    const fileUrl = file.getUrl();
    return fileUrl;
  }

}

// Genera un ID único para la solicitud
function generarSolicitudId(sheetRegistro) {
  var lastRow = sheetRegistro.getLastRow();
  if (lastRow > 1) {
    var lastId = sheetRegistro.getRange(lastRow, 1).getValue();
    return lastId + 1;
  }
  return 1;
}

// Registra los productos en la base de datos (google sheet)
function registrarProductos(form, solicitudId, sheetRegistro, cotizacionUrl, formatSolicitante) {
  var fechaRegistro = new Date();
  var totalCompra = 0;
  var productos = obtenerDatosProductos(form);

  var observaciones = form.observaciones || "";

  var requiereCapex = form.requiereCapex; // "Sí" o "No"
  var tipoCapex = form.tipoCapex || ""; // Tipo de inversión seleccionado o vacío


  productos.forEach((producto) => {
    var subtotal = calcularSubtotal(producto.cantidad, producto.precio);
    totalCompra += subtotal;
  });

  // Arreglo para devolver los productos registrados
  var productosRegistrados = [];

  console.log("Solicitante en la función registrarProductos: " + formatSolicitante.solicitante);
  console.log("Solicitante en la función registrarProductos: " + formatSolicitante.solicitante.names);
  console.log("Solicitante en la función registrarProductos: " + formatSolicitante.solicitante.email);
  console.log("Solicitante en la función registrarProductos: " + formatSolicitante.solicitante);

  // Segundo forEach: registrar productos
  productos.forEach((producto) => {
    var subtotal = calcularSubtotal(producto.cantidad, producto.precio);

    // Armar fila con los 32 campos exactos
    var fila = [
      solicitudId,                                          // ID
      formatSolicitante.solicitante.names,                 // NOMBRE Y APELLIDOS
      formatSolicitante.solicitante.email,                 // EMAIL
      form.razonCompra,                                     // RAZÓN DE COMPRA
      fechaRegistro,                                        // FECHA DE REGISTRO
      form.prioridad,                                       // PRIORIDAD
      form.justificacion,                                   // JUSTIFICACIÓN GENERAL
      producto.nombre,                                      // PRODUCTO
      producto.marca,                                       // MARCA
      producto.especificaciones,                            // ESPECIFICACIONES
      producto.centroCosto,                                 // CENTRO DE COSTO
      producto.cantidad,                                    // CANTIDAD
      producto.precio,                                      // PRECIO
      subtotal,                                             // SUBTOTAL
      observaciones                                         // OBSERVACIONES DE LA COMPRA
    ];

    // Columnas 16–20: Aprobaciones
    if (totalCompra <= 500) {
      var jefeArea = determinarDestinatario(totalCompra, formatSolicitante);
      fila.push(
        "Pendiente",                                        // ESTADO PRIMERA APROBACIÓN
        jefeArea.solicitante.email,                         // APROBADO POR
        "",                                                 // FECHA DE APROBACIÓN
        "",                                                 // ESTADO SEGUNDA APROBACIÓN
        ""                                                  // APROBADO POR GERENTE
      );
    } else {
      var primerAprobador = determinarDestinatario(totalCompra, formatSolicitante);
      fila.push(
        "Pendiente",                                        // ESTADO PRIMERA APROBACIÓN
        primerAprobador.solicitante.email,                 // APROBADO POR
        "",                                                 // FECHA DE APROBACIÓN
        "Pendiente",                                        // ESTADO SEGUNDA APROBACIÓN
        ""                                                  // APROBADO POR GERENTE
      );
    }

    fila.push(
      "",                     // FECHA DE APROBACIÓN (Gerente)
      cotizacionUrl,          // LINK DE COTIZACIÓN
      producto.justifyCompraProds || "", // JUSTIFICACIÓN INDIVIDUAL
      form.DescriptCompra || "",         // DESCRIPCIÓN DE LA COMPRA
      "",                     // LINK DE CAPEX
      "",                     // LINK DE CAPEX FIRMADO
      "",                     // Fecha Última Notificación (Gerente Área)
      "",                     // Contador Notificaciones (Gerente Área)
      "",                     // Fecha Última Notificación (Gerente General)
      "",                     // Contador Notificaciones (Gerente General)
      requiereCapex,          // ES ACTIVO
      tipoCapex               // TIPO DE CAPEX
    );

    // Asegurar 32 columnas
    //while (fila.length < 32) fila.push("");

    // Escribir en hoja
    sheetRegistro.appendRow(fila);

    // Guardar producto registrado
    productosRegistrados.push({
      solicitudId: solicitudId,
      solicitante: formatSolicitante.solicitante.names,
      email: formatSolicitante.solicitante.email,
      razonCompra: form.razonCompra,
      fechaRegistro: fechaRegistro,
      prioridad: form.prioridad,
      justificacion: form.justificacion,
      nombreProducto: producto.nombre,
      marca: producto.marca,
      especificaciones: producto.especificaciones,
      centroCosto: producto.centroCosto,
      cantidad: producto.cantidad,
      precio: producto.precio,
      subtotal: subtotal,
      observaciones: observaciones,
      estado1: "Pendiente",
      aprobador1: totalCompra <= 500 ? jefeArea.solicitante.email : primerAprobador.solicitante.email,
      cotizacionUrl: cotizacionUrl,
      justificacionPorProducto: producto.justifyCompraProds || "",
      descripcionCompra: form.DescriptCompra || "",
      requiereCapex: requiereCapex,
      tipoCapex: tipoCapex
    });
  });

  return {
    totalCompra: totalCompra,
    productos: productosRegistrados
  };
}

// Maneja la solicitud de actualización de estado
function handleEstadoRequest(e, cotizacionUrl) {
  var solicitudId = parseInt(e.parameter.solicitudId, 10);
  var nuevoEstado = e.parameter.estado;
  //Id de la persona que aprobó o desaprobó la solicitud 
  var aprobadorId = e.parameter.aprobadoresEmail;
  var numberAprobadores = e.parameter.numberAprobadores;

  if (!solicitudId || !nuevoEstado || !aprobadorId) {
    return ContentService.createTextOutput(
      "Solicitud ID o estado faltante o aprobadorEmail."
    );
  }

  var resultado = actualizarEstado(solicitudId, nuevoEstado, aprobadorId, numberAprobadores, cotizacionUrl);

  return ContentService.createTextOutput(resultado);
}

// Determina la columna de estado a actualizar según el total de la compra y el ID de la solicitud
function obtenerColumnaEstado(totalCompra, solicitudId, numberAprobadores) {
  //var data = cargarDataHoja("index");
  var data = obtenerDatos("index");

  var columnaEstado = 0;
  console.log("Número de aprobadores en la función obtenerColumnaEstado: " + numberAprobadores);
  data.forEach((row, i) => {
    if (row[0] == solicitudId) {
      if (totalCompra <= 500) {
        columnaEstado = 15; // Columna para jefe del área
      } else {

        if (numberAprobadores == 1) {
          columnaEstado = 15; // Columna para gerente de área
        } else if (numberAprobadores == 2) {
          columnaEstado = 18; // Columna para gerente general
        }
      }
    }
  });
  return columnaEstado;
}

// Extrae los datos de los productos del evento POST
function obtenerDatosProductos(form) {

  const jsonProductos = form.productListJson;

  if (!jsonProductos) {
    Logger.log("No se recibió productListJson.");
    return;
  }

  // Parsear el JSON
  const productos = JSON.parse(jsonProductos);

  /*var nombres = limpiarCadena(form["productNames[]"]).split("!");
  var marcas = limpiarCadena(form["productBrands[]"]).split("!");
  var cantidades = limpiarCadena(form["productQuantities[]"]).split("!");
  var precios = limpiarCadena(form["productPrices[]"]).split("!");
  var especificaciones = limpiarCadena(form["productSpecs[]"]).split("!");
  var centroCostos = limpiarCadena(form["productCentroCostos[]"]).split("!");
  var justifyCompraProds = limpiarCadena(form["justifyCompraProds[]"]).split("!");


  return nombres.map((nombre, i) => ({
    nombre: nombre,
    marca: marcas[i] || "",
    cantidad: parseFloat(cantidades[i]) || 0,
    precio: parseFloat(precios[i]) || 0,
    especificaciones: especificaciones[i] || "",
    centroCosto: centroCostos[i] || "",
    justifyCompraProds: justifyCompraProds[i] || ""
  })); */

  // Retornar la lista de productos tal cual o transformada si deseas
  return productos.map(producto => ({
    nombre: producto.name || "",
    marca: producto.brand || "",
    cantidad: parseFloat(producto.quantity) || 0,
    precio: parseFloat(producto.price) || 0,
    especificaciones: producto.specs || "",
    centroCosto: producto.centroCostos || "",
    justifyCompraProds: producto.justificacion || ""
  }));
}
function limpiarCadena(cadena) {
  if (cadena == undefined || cadena == null) {
    return "";
  }
  return cadena.replace(/!\s*$/, "");
}

// Calcula el subtotal de un producto
function calcularSubtotal(cantidad, precio) {
  if (!isNaN(cantidad) && !isNaN(precio)) {
    return cantidad * precio;
  }
  return 0;
}

function actualizarEstado(solicitudId, nuevoEstado, aprobadorId, numberAprobadores, cotizacionUrl) {
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var sheet = ss.getSheetByName("index");
  //var data = sheet.getDataRange().getValues();

  var sheet = conectarHoja("index");
  var data = obtenerDatos("index");

  var totalCompra = costoTotalSolicitud(solicitudId);
  var solicitanteEmail = null;
  var columnaEstado = obtenerColumnaEstado(totalCompra, solicitudId, numberAprobadores);
  var usuarioAprobador = buscarSolicitantePorId(aprobadorId);
  var formatoUsuarioAprobador = transformarData(usuarioAprobador);

  var registrosActualizados = [];

  if (columnaEstado > 0) {
    data.forEach((row, i) => {
      if (row[0] == solicitudId) {
        //verificar si el estado actual se puede modificar
        var estadoActual = row[columnaEstado];

        if (estadoActual === "Pendiente") {

          sheet.getRange(i + 1, columnaEstado + 1).setValue(nuevoEstado);
          sheet.getRange(i + 1, columnaEstado + 2).setValue(formatoUsuarioAprobador.solicitante.email);
          sheet.getRange(i + 1, columnaEstado + 3).setValue(new Date());
          registrosActualizados.push(row);
          solicitanteEmail = row[2];
        } else {
          return ContentService.createTextOutput(
            `La solicitud ya ha sido ${estadoActual} y no puede ser actualizada nuevamente. `
          );
        }
      }
    });
  } else {
    // Esta opción se cumplicará siempre y la columna de estado sea igual a cero lo que implica que la solicitud ha sido aprobada o desaprobada
    return "No se pudo determinar la columna de estado.";
  }


  if (registrosActualizados.length > 0) {
    if (solicitanteEmail) {
      // Enviar correo para notificación al emisor del correo
      enviarCorreoRemitente(solicitanteEmail, solicitudId, nuevoEstado, formatoUsuarioAprobador, columnaEstado, totalCompra);
    }
    // Solo si el nuevo estado es aprobado entonces el flujo continúa. 
    if (nuevoEstado === "Aprobado") {
      enviarCorreoAprobado(registrosActualizados, totalCompra, formatoUsuarioAprobador, columnaEstado, cotizacionUrl);
    }

    //Termina el flujo cada vez que se actualiza la solicitud con un estado de rechazado 
    return `El estado de la solicitud ha sido actualizado a: ${nuevoEstado}`;
  } else {
    return "Solicitud no encontrada. xxxxxxxx";
  }
}


// Función para actualizar el estado de la solicitud en la hoja
function actualizarEstadoSolicitud(solicitudId, totalCompra, nuevoEstado, usuario) {
  //var hoja = cargarDataHoja("index"); // Asegúrate de que esta función cargue correctamente los datos de la hoja
  var hoja = obtenerDatos("index")
  var columnaEstado = obtenerColumnaEstado(totalCompra, solicitudId);

  hoja.forEach((row, i) => {
    if (row[0] == solicitudId) {
      if (columnaEstado > 0) {
        // Verifica si el estado puede ser actualizado
        var estadoActual = row[columnaEstado];
        if (estadoActual === "Pendiente" || estadoActual === "Observado") {
          hoja.getRange(i + 1, columnaEstado + 1).setValue(nuevoEstado);
          console.log(`Estado actualizado a '${nuevoEstado}' en la columna ${columnaEstado} por el usuario ${usuario}`);
        } else {
          console.log(`La solicitud ya ha sido ${estadoActual} y no puede ser actualizada nuevamente por ${usuario}`);
        }
      } else {
        console.log("No se pudo determinar la columna de estado.");
      }
    }
  });
}

//Funcion para cargar la data de la hoja de google sheet
function cargarDataHoja(nombreOja) {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName(nombreOja);
  var data = hoja.getDataRange().getValues();
  return data;
}

// Calcula el costo total de la solicitud
function costoTotalSolicitud(solicitudId) {
  //var data = cargarDataHoja("index");
  var data = obtenerDatos("index");
  var totalCompra = 0;

  data.forEach((row) => {
    if (row[0] == solicitudId) {
      totalCompra += parseFloat(row[13]);
    }
  });

  return totalCompra;
}

// Enviar correo electrónico después de la primera o segunda aprobación 
function enviarCorreoAprobado(registros, totalCompra, formatoUsuarioAprobador, columnaEstado, cotizacionUrl) {

  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Compras");
  //var data = sheet.getDataRange().getValues();

  var data = obtenerDatos("Compras");

  var destinatario = data[1][1];

  if (totalCompra <= 500) {

    //destinatario = "jerrytocto@gmail.com"; //Correo del área de compras
    const namesAprobador = formatoUsuarioAprobador.solicitante.names;
    const cargoAprobador = formatoUsuarioAprobador.solicitante.cargo;

    var nombreCargoAprobador = formatoUsuarioAprobador.solicitante.names + " - " + formatoUsuarioAprobador.solicitante.cargo;
    enviarCorreoCompras(registros, nombreCargoAprobador, destinatario, totalCompra, cotizacionUrl);
    eviarSolicitudAprobadaAlDBLogistica(registros, totalCompra, cotizacionUrl, namesAprobador, cargoAprobador);
  } else {
    //Verificar la columna de estado que se ha modificado.
    if (columnaEstado == 15) {

      //Fué aprobado por el gerente de área, por lo que se debe extraer el correo del gerente general 
      var aprobadorId = formatoUsuarioAprobador.solicitante.jefe;
      var jefeAprobador = cargarDataUsers(aprobadorId);
      var formatoJefeAprobador = transformarData(jefeAprobador);
      console.log("Si lee que la solicitud está entre 500 y 1000. Y se envía al correo de compras: " + columnaEstado)
      destinatario = formatoJefeAprobador.solicitante.email; //Correo del gerente general



      var nombreAreaGA = formatoJefeAprobador.solicitante.nombre + " - " + formatoJefeAprobador.solicitante.cargo;

      enviarCorreoGerenteGeneral(
        registros,
        formatoUsuarioAprobador,
        destinatario,
        totalCompra,
        cotizacionUrl,
        false
      );

      //Correo ha sido revisado y aprobado por el gerente general
    } else if (columnaEstado == 18) {



      var nombreAreaGG = formatoUsuarioAprobador.solicitante.names + " - " + formatoUsuarioAprobador.solicitante.cargo;
      const aprobador = formatoUsuarioAprobador.solicitante.names;
      const cargoAprobador = formatoUsuarioAprobador.solicitante.cargo;

      //var nombreCargoAprobadores = nombreAreaGA + " y " + nombreAreaGG;
      var nombreCargoAprobadores = nombreAreaGG;

      var esActivo = registros[0][30];
      Logger.log("Es activo en la función enviar correo aprobado: " + esActivo);
      var correoSolicitante = registros[0][2];
      var fileSolictante = buscarUserPorEmail(correoSolicitante);
      var formatoSolicitante = transformarData(fileSolictante);

      var idAreaSolicitante = formatoSolicitante.solicitante.area;

      var nombreAreaSolicitante = obtenerNombreArea(idAreaSolicitante) ? obtenerNombreArea(idAreaSolicitante) : "";

      var debeFirmarCapex = (nombreAreaSolicitante === "TI" || nombreAreaSolicitante === "IT") || (esActivo === "SI" && totalCompra > 1000);

      if (debeFirmarCapex) {
        var idSolicitud = registros[0][0];
        //var capexFirmado = generateCapex(totalCompra, registros, nombreCargoAprobadores, true, nombreAreaGG);
        //enviarCorreoGerenteGeneral
        var pdfCapexFirmado = firmarCapex(nombreAreaGG, registros);

        if (pdfCapexFirmado) {
          agregarLinkCapexFirmado(idSolicitud, pdfCapexFirmado.link);
        }

      }

      enviarCorreoCompras(
        registros,
        (nombreAreaGG),
        destinatario,
        totalCompra,
        cotizacionUrl
      );
      eviarSolicitudAprobadaAlDBLogistica(registros, totalCompra, cotizacionUrl, aprobador, cargoAprobador);
    }
  }
}

function obtenerNombreArea(idAreaSolicitante) {
  var datosDeAreas = obtenerDatos("Areas");

  Logger.log("Datos de Áreas:", datosDeAreas);

  for (var i = 0; i < datosDeAreas.length; i++) {
    if (datosDeAreas[i][0] == idAreaSolicitante) { // Comparar con ==
      Logger.log("Área encontrada:", datosDeAreas[i][1]);
      return datosDeAreas[i][1]; // Devolver directamente el nombre del área
    }
  }

  Logger.log("No se encontró el área para el ID:", idAreaSolicitante);
  return ""; // Devolver una cadena vacía si no se encuentra
}

//Enviar data al sheet de logística 
function eviarSolicitudAprobadaAlDBLogistica(registrosAprobados, totalCompra, cotizacionUrl, namesAprobador, cargoAprobador) {
  var sheetLogistica = conectarHojaLogistica("Solicitudes");
  var fecha = new Date();
  //const totalCompra = totalCompra.toFixed(2);
  const nombreAprobador = namesAprobador;
  const cargoApro = cargoAprobador;
  //const cotizacionFile = obtenerCotizacionDeDrive(cotizacionUrl);

  const solicitudId = registrosAprobados[0][0];
  const emisor = registrosAprobados[0][1];
  const emailEmisor = registrosAprobados[0][2];
  const razonDeCompra = registrosAprobados[0][3];
  const fechaSolicitud = registrosAprobados[0][4];
  const prioridad = registrosAprobados[0][5];
  const justificacion = registrosAprobados[0][6];
  const observaciones = registrosAprobados[0][14] ? registrosAprobados[0][14] : "";

  registrosAprobados.forEach((producto) => {
    var subtotal = calcularSubtotal(producto[11], producto[12]);

    var fila = [
      solicitudId,
      emisor,
      emailEmisor,
      razonDeCompra,
      fechaSolicitud,
      prioridad,
      justificacion,
      producto[7],
      producto[8],
      producto[9],
      producto[10],
      producto[11],
      producto[12],
      subtotal,
      observaciones,
      nombreAprobador,
      cargoApro,
      fecha,
      cotizacionUrl = ! "" ? cotizacionUrl : ""
    ];

    sheetLogistica.appendRow(fila);
  });
}

//Función para buscar un users por email 
function buscarUserPorEmail(email) {

  var data = obtenerDatos("users");

  for (var i = 0; i < data.length; i++) {
    if (data[i][2] === email) {
      return data[i];
    }
  }
}

// Función para añadir nuevo link del capex firmado
function agregarLinkCapexFirmado(idSolicitud, nuevoCapexLink) {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("index");
  //var data = sheet.getDataRange().getValues();
  var sheet = conectarHoja("index");
  var data = obtenerDatos("index");

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === idSolicitud) {
      sheet.getRange(i + 1, 26).setValue(nuevoCapexLink);
    }
  }
}

function enviarCorreoGerenteGeneral(
  registrosAprobados,
  formatoSolicitante,
  destinatario,
  totalCompra,
  cotizacionUrl,
  esNotificacion
) {
  var htmlTemplate = HtmlService.createTemplateFromFile("tablaRequisitosEmail");
  var scriptUrl = obtenerUrl();
  var esActivo = registrosAprobados[0][30];
  var linkCotizacion = registrosAprobados[0][21] ? registrosAprobados[0][21] : "";

  // Inicializar linkCapex con un valor por defecto
  htmlTemplate.linkCapex = "";  // <- Añadido inicialización

  // Configuración del template HTML
  htmlTemplate.solicitudId = registrosAprobados[0][0];
  htmlTemplate.emisor = registrosAprobados[0][1];
  htmlTemplate.razonDeCompra = registrosAprobados[0][3];
  htmlTemplate.fechaSolicitud = registrosAprobados[0][4];
  htmlTemplate.justificacion = registrosAprobados[0][6];
  htmlTemplate.centroDeCosto = registrosAprobados[0][10];
  htmlTemplate.observaciones = registrosAprobados[0][14] ? registrosAprobados[0][14] : "";
  htmlTemplate.descriptionCompra = registrosAprobados[0][23];
  htmlTemplate.activo = registrosAprobados[0][30] ? registrosAprobados[0][30] : "NO";
  htmlTemplate.tipoInversion = registrosAprobados[0][31] ? registrosAprobados[0][31] : "";
  htmlTemplate.tablaSolicitud = registrosAprobados;
  htmlTemplate.numberAprobadores = 2;
  htmlTemplate.scriptUrl = scriptUrl;
  htmlTemplate.linkCotizacion = linkCotizacion;

  var nombreAreaGA = formatoSolicitante.solicitante.names + " - " + formatoSolicitante.solicitante.cargo;

  htmlTemplate.aprobadoresEmail = formatoSolicitante.solicitante.jefe;
  htmlTemplate.nombreCargoAprobador = nombreAreaGA;
  htmlTemplate.mostrarCampoAprobador = 1;
  htmlTemplate.paraAprobar = true;
  htmlTemplate.totalCompra = totalCompra.toFixed(2);

  var asunto;

  try {
    if (!esNotificacion) {
      Logger.log("Es activo: " + esActivo);
      if (esActivo === "SI") {
        var correoSolicitante = registrosAprobados[0][2]; // extraigo el correo del solicitante 
        var fileSolictante = buscarUserPorEmail(correoSolicitante);  // Busco el usuario por email
        var formatoSolicitante = transformarData(fileSolictante);    // Transformo la data del usuario

        var idAreaSolicitante = formatoSolicitante.solicitante.area; // Extraigo el id del área del solicitante

        // Extraigo el nombre del área del solicitante 
        var nombreAreaSolicitante = obtenerNombreArea(idAreaSolicitante) ? obtenerNombreArea(idAreaSolicitante) : "";
        Logger.log("Nombre del área solicitante: " + nombreAreaSolicitante);

        if (nombreAreaSolicitante === "IT" || nombreAreaSolicitante === "TI") {
          Logger.log("El área solicitante es IT o TI");
          var capex = generateCapex(totalCompra, registrosAprobados, nombreAreaGA, false);
          Logger.log("Link del capex: " + capex.link);

          actualizarHojaConEnlaceCapex(htmlTemplate.solicitudId, capex.link);
          htmlTemplate.linkCapex = capex.link;
          asunto = "Nueva Solicitud de Compra para Aprobar";
        } else if (totalCompra > 1000) {
          var capex = generateCapex(totalCompra, registrosAprobados, nombreAreaGA, false);
          actualizarHojaConEnlaceCapex(htmlTemplate.solicitudId, capex.link);
          htmlTemplate.linkCapex = capex.link;
          asunto = "Nueva Solicitud de Compra para Aprobar";
        } else {
          Logger.log("El área solicitante no es IT o TI y la solicitud es menor a $1000");
          asunto = "Nueva Solicitud de Compra para Aprobar";
        }
      } else {
        Logger.log("Lo que se está solicitando no es activo");
        asunto = "Nueva Solicitud de Compra para Aprobar";
      }
    } else {
      var capexUrl = registrosAprobados[0][24];
      htmlTemplate.linkCapex = capexUrl || "";
      asunto = `¡PENDIENTE DE APROBACIÓN! SOLICITUD DE COMPRA N°${htmlTemplate.solicitudId}`;
    }

    var html = htmlTemplate.evaluate().getContent();

    GmailApp.sendEmail(
      destinatario,
      asunto,
      "Nueva solicitud de compra.",
      {
        htmlBody: html
      }
    );
  } catch (error) {
    Logger.log("Error al enviar el correo: " + error);
    throw error;
  }
}
//https://drive.google.com/drive/folders/1T_vTy4BVj3ypQbes5jMm4yWYOaWmacF3
//function actualizarHojaConEnlaceCapex(solicitudId, enlaceCapex)
function actualizarHojaConEnlaceCapex(solicitudId, enlaceCapex) {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("index");
  //var data = sheet.getDataRange().getValues();
  console.log("Solicitud Id: Función actualizarHojaConEnlaceCapex:  " + solicitudId);
  console.log("Enlace de capex : Función actualizarHojaConEnlaceCapex:  " + enlaceCapex);


  var spreadsheetId = '1-sTiLcNB8V3N3vYsSsqYoKYt7jXYqo0Ymk2_l-b-8QU';// Id de la hoja sheet
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName("index");
  var data = sheet.getDataRange().getValues();

  //var sheet = conectarHoja("index");
  //var data = obtenerDatos("index");
  var actualizaciones = [];

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(solicitudId)) {
      actualizaciones.push({
        range: sheet.getRange(i + 1, 25),
        value: enlaceCapex
      });
    }
  }

  // Aplicar todas las actualizaciones de una vez
  if (actualizaciones.length > 0) {
    actualizaciones.forEach(function (update) {
      update.range.setValue(update.value);
    });
    Logger.log('Se actualizaron ' + actualizaciones.length + ' registros para la solicitud ' + solicitudId);
  } else {
    Logger.log('No se encontraron registros para la solicitud ' + solicitudId);
  }
}


// Enviar correo de notificación al remitente
function enviarCorreoRemitente(email, solicitudId, estado, formatoSolicitante, columnaEstado, totalCompra) {

  var solicitante = cargarDataUsersPorEmail(email);
  var formatSolic = transformarData(solicitante);
  var nombreSolicitante = formatSolic.solicitante.names;

  var subject = "";
  var bodyContent = "";
  var statusClass = "";
  var statusIcon = "";

  // Determinar clase CSS y icono según el estado
  if (estado === "Aprobado") {
    statusClass = "success";
    statusIcon = "✅";
  } else if (estado === "Rechazado") {
    statusClass = "danger";
    statusIcon = "❌";
  } else {
    statusClass = "warning";
    statusIcon = "⏳";
  }

  //Notificar cuando la solicitud es menor a 500
  if (totalCompra <= 500 && estado != "Aprobado") {
    subject = `Actualización en su solicitud de comrpa #${solicitudId}`;
    bodyContent = `
      <p>Estimado(a) <strong>${nombreSolicitante}</strong>,</p>
      <p>LE INFORMAMOS QUE SU SOLICITUD DE COMPRA CON ID <strong>${solicitudId}</strong> ha sido <span class="badge bg-${statusClass}">${estado}</span> POR ${formatoSolicitante.solicitante.names} - ${formatoSolicitante.solicitante.cargo}.</p>
      <p>AGRADECEMOS SU ATENCIÓN.</p>
    `;

  } else if (totalCompra <= 500 && estado === "Aprobado") {
    subject = `Actualización en su solicitud de compra #${solicitudId}`;
    bodyContent = `
      <p>Estimado(a) <strong>${nombreSolicitante}</strong>,</p>
      <p>Nos complace informarle que su solicitud de compra con ID <strong>${solicitudId}</strong> ha sido <span class="badge bg-${statusClass}">${estado}</span> por ${formatoSolicitante.solicitante.names} - ${formatoSolicitante.solicitante.cargo}.</p>
      <p>Su solicitud ahora ha sido derivada al área de logística(Compras) para materializar lo solicitado.</p>
      <p>Agradecemos de antemano su comprensión y paciencia en este proceso. Si tiene alguna consulta o requiere información adicional, no dude en ponerse en contacto con el área de Logística.</p>
    `;

  } else if (totalCompra > 500 && columnaEstado == 15 && estado === "Aprobado") {

    var aprobadorId = formatoSolicitante.solicitante.jefe;
    var jefeAprobador = cargarDataUsers(aprobadorId);
    var formatoJefeAprobador = transformarData(jefeAprobador);

    subject = `Actualización en su solicitud de compra #${solicitudId}`;
    bodyContent = `
      <p>Estimado(a) <strong>${nombreSolicitante}</strong>,</p>
      <p>Le informamos que su solicitud de compra con ID <strong>${solicitudId}</strong> ha sido <span class="badge bg-${statusClass}">${estado}</span> por ${formatoSolicitante.solicitante.names} - ${formatoSolicitante.solicitante.cargo}.</p>
      <p>Actualmente, se encuentra en evaluación por ${formatoJefeAprobador.solicitante.names} - ${formatoJefeAprobador.solicitante.cargo}.</p>
      <p>No dude en contactar al responsable de la evaluación ante cualquier duda o consulta.</p>
    `;

  } else if (totalCompra > 500 && columnaEstado == 15 && estado != "Aprobado") {
    subject = `Actualización en su solicitud de compra #${solicitudId}`;
    bodyContent = `
      <p>Estimado(a) <strong>${nombreSolicitante}</strong>,</p>
      <p>Le informamos que su solicitud de compra con ID <strong>${solicitudId}</strong> ha sido <span class="badge bg-${statusClass}">${estado}</span> por ${formatoSolicitante.solicitante.names} - ${formatoSolicitante.solicitante.cargo}.</p>
      <p>Agradecemos su comprensión.</p>
    `;

  } else if (totalCompra > 500 && columnaEstado == 18 && estado === "Aprobado") {
    subject = `Novedades en tu solicitud de compra #${solicitudId}`;
    bodyContent = `
      <p>Estimado(a) <strong>${nombreSolicitante}</strong>,</p>
      <p>Nos complace informarle que su solicitud de compra con ID <strong>${solicitudId}</strong> ha sido <span class="badge bg-${statusClass}">${estado}</span> por ${formatoSolicitante.solicitante.names} - ${formatoSolicitante.solicitante.cargo}.</p>
      <p>Actualmente, su solicitud se encuentra siendo evaluada por el área de compras.</p>
      <p>Si desea saber más sobre el estado de su solicitud, no dude en ponerse en contacto con el área de compras.</p>
    `;

  } else if (totalCompra > 500 && columnaEstado == 18 && estado != "Aprobado") {
    subject = `Actualización en su solicitud de compra #${solicitudId}`;
    bodyContent = `
      <p>Estimado(a) <strong>${nombreSolicitante}</strong>,</p>
      <p>Le informamos que su solicitud de compra con ID <strong>${solicitudId}</strong> ha sido <span class="badge bg-${statusClass}">${estado}</span> por ${formatoSolicitante.solicitante.names} - ${formatoSolicitante.solicitante.cargo}.</p>
      <p>No dude en contactar al responsable de la evaluación ante cualquier duda o consulta.</p>
    `;
  }


  // Cargar la plantilla HTML
  var template = HtmlService.createTemplateFromFile('PlantillaNotificacion');

  // Asignar datos a la plantilla
  template.SUBJECT = subject;
  template.SOLICITUD_ID = solicitudId;
  template.ESTADO = estado;
  template.STATUS_CLASS = statusClass;
  template.BODY_CONTENT = bodyContent;
  template.APROBADO_ICON = statusIcon;
  template.CURRENT_YEAR = new Date().getFullYear();

  // Evaluar plantilla y obtener HTML final
  var htmlFinal = template.evaluate().getContent();

  // Enviar correo HTML
  GmailApp.sendEmail(
    email,
    subject,
    "Este correo requiere un lector de correo compatible con HTML para visualizarse correctamente.", // Mensaje alternativo para clientes que no soportan HTML
    { htmlBody: htmlFinal }
  );
}

// Enviar correo al área de compras
function enviarCorreoCompras(
  registrosAprobados,
  aprobadoresEmail,
  destinatario,
  totalCompra,
  cotizacionUrl
) {
  var htmlTemplate = HtmlService.createTemplateFromFile("tablaRequisitosEmail");
  //var scriptUrl = ScriptApp.getService().getUrl(); // Obtén la URL del script
  var scriptUrl = obtenerUrl();
  var linkCotizacion = cotizacionUrl ? cotizacionUrl : "";
  var linkCapex = "";

  htmlTemplate.solicitudId = registrosAprobados[0][0];
  htmlTemplate.emisor = registrosAprobados[0][1];
  htmlTemplate.razonDeCompra = registrosAprobados[0][3];
  htmlTemplate.fechaSolicitud = registrosAprobados[0][4];
  htmlTemplate.justificacion = registrosAprobados[0][6];
  htmlTemplate.centroDeCosto = registrosAprobados[0][10];
  htmlTemplate.observaciones = registrosAprobados[0][14] ? registrosAprobados[0][14] : "";
  htmlTemplate.descriptionCompra = registrosAprobados[0][23];
  htmlTemplate.activo = registrosAprobados[0][30] ? registrosAprobados[0][30] : "NO";
  htmlTemplate.tipoInversion = registrosAprobados[0][31] ? registrosAprobados[0][31] : "";
  htmlTemplate.tablaSolicitud = registrosAprobados;
  htmlTemplate.aprobadoresEmail = aprobadoresEmail;
  htmlTemplate.mostrarCampoAprobador = 1;
  htmlTemplate.paraAprobar = false;
  htmlTemplate.numberAprobadores = 1;
  htmlTemplate.scriptUrl = scriptUrl; // Pasar la URL del script a la plantilla
  htmlTemplate.linkCotizacion = linkCotizacion;
  htmlTemplate.linkCapex = linkCapex;

  htmlTemplate.nombreCargoAprobador = aprobadoresEmail;

  htmlTemplate.totalCompra = totalCompra.toFixed(2);

  var html = htmlTemplate.evaluate().getContent();

  const correosDeCompras = obtenerCorreosCompras();
  if (correosDeCompras.length > 0) {
    correosDeCompras.forEach((email) => {
      //Enviar correo inicial para el área de compras 
      GmailApp.sendEmail(email, ("NUEVA SOLICITUD DE COMPRA APROBADA CON ID" + registrosAprobados[0][0]) + "", "ESTIMADO(A), ESTA SOLICITUD FUÉ APROBADA", {
        htmlBody: html
      });
    });

  }

}

// Función para enviar email para su aprobación
function enviarEmail(retornoDeResitrarProductos, solicitudId, formatSolicitante, esAviso) {

  var htmlTemplate = HtmlService.createTemplateFromFile("tablaRequisitosEmail");
  //var scriptUrl = ScriptApp.getService().getUrl(); // Obtén la URL del script
  var scriptUrl = obtenerUrl();
  var listProductos = retornoDeResitrarProductos.productos;

  // Formatear los precios y subtotales de los productos
  listProductos = listProductos.map((p) => {
    return {
      ...p,
      precio: formatearMoneda(p.precio),
      subtotal: formatearMoneda(p.subtotal)
    };
  });

  // Obtener el enlace del documento PDF desde la columna 21
  var linkCotizacion = listProductos[0].linkCotizacion || "";

  htmlTemplate.listProductos = listProductos;
  htmlTemplate.totalCompra = formatearMoneda(retornoDeResitrarProductos.totalCompra);
  htmlTemplate.solicitudId = solicitudId;
  htmlTemplate.emisor = formatSolicitante.solicitante.names + " - " + formatSolicitante.solicitante.cargo;
  htmlTemplate.razonDeCompra = listProductos[0].razonCompra;
  htmlTemplate.fechaSolicitud = listProductos[0].fechaRegistro;
  htmlTemplate.justificacion = listProductos[0].justificacion;
  htmlTemplate.descriptionCompra = listProductos[0].descripcionCompra;
  htmlTemplate.activo = listProductos[0].requiereCapex || "NO";
  htmlTemplate.tipoInversion = listProductos[0].tipoCapex || "";
  htmlTemplate.centroDeCosto = listProductos[0].centroCosto;
  htmlTemplate.observaciones = listProductos[0].observaciones || "";
  htmlTemplate.mostrarCampoAprobador = 0;
  htmlTemplate.paraAprobar = true;
  htmlTemplate.nombreCargoAprobador = '';
  htmlTemplate.numberAprobadores = 1;
  htmlTemplate.scriptUrl = scriptUrl; // Pasar la URL del script a la plantilla
  htmlTemplate.linkCotizacion = linkCotizacion;
  htmlTemplate.linkCapex = "";

  var destinatario = determinarDestinatario(retornoDeResitrarProductos.totalCompra, formatSolicitante);

  var destinatarioId = destinatario.solicitante.id;

  htmlTemplate.aprobadoresEmail = destinatarioId;

  // Configurar el correo para el solicitante
  if (!esAviso) {
    var html = htmlTemplate.evaluate().getContent();

    //Correo para la persona encargada de la aprobación
    GmailApp.sendEmail(destinatario.solicitante.email, ("SOLICITUD DE COMPRA " + listProductos[0].solicitudId), "MENSAJE DEL EMAIL", {
      htmlBody: html
    });

    htmlTemplate.mostrarCampoAprobador = 0;
    htmlTemplate.paraAprobar = false;
    var htmlParaSolicitante = htmlTemplate.evaluate().getContent();

    // Enviar correo para el solicitante
    GmailApp.sendEmail(formatSolicitante.solicitante.email, ("EL REGISTRO DE TU SOLICITUD DE COMPRA " + listProductos[0].solicitudId) + " FUE EXITOSA", "ESTIMADO, TU SOLICITUD DE COMPRA ESTÁ EN CURSO", {
      htmlBody: htmlParaSolicitante
    });

    const correosDeCompras = obtenerCorreosCompras();

    correosDeCompras.forEach((email) => {
      //Enviar correo inicial para el área de compras 
      GmailApp.sendEmail(email, ("NUEVA SOLICITUD DE COMPRA GENERADA CON ID" + listProductos[0].solicitudId) + "", "ESTIMADO(A), ESTA SOLICITUD ESTÁ PENDIENTE DE APROBACIÓN", {
        htmlBody: htmlParaSolicitante
      });
    });

  } else {
    var html = htmlTemplate.evaluate().getContent();
    //Correo para la persona encargada de la aprobación
    GmailApp.sendEmail(destinatario.solicitante.email, ("NOTIFICACIÓN DE APROBACIÓN PARA LA SOLICITUD DE COMPRA " + listProductos[0].solicitudId), "MENSAJE DEL EMAIL", {
      htmlBody: html
    });
  }
}

function formatearMoneda(totalCompra) {
  if (totalCompra === undefined || totalCompra === null) {
    return "0.00"; // Retorna un valor por defecto si totalCompra es indefinido o nulo

  }
  var totalFormateado = totalCompra.toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
  return totalFormateado;
}

//Extraer lista de correos para el área de compras 
function obtenerCorreosCompras() {

  var correos = [];
  var data = obtenerDatos("Compras");

  // Comenzar desde el índice 1 para omitir el encabezado
  for (let i = 1; i < data.length; i++) {
    correos.push(data[i][1]); // Asume que los correos están en la segunda columna
  }
  console.log("Estos son los correos de compras: " + correos);
  return correos;
}


//Determinar el destinatario / persona que aprobará la solicitud 
function determinarDestinatario(totalCompra, formatoSolicitante) {

  //Verificar el código del cargo del solicitante 
  var cargo = formatoSolicitante.solicitante.codigoCargo;

  if (cargo == 2) { //Solicitud hecha por el gerente de área
    //Enviar correo hacia el mismo gerente de área
    return formatoSolicitante;

  } else if (cargo == 3) {  //Solicitud hecha por el jefe de área
    if (totalCompra <= 500) {
      //Enviar correo hacia el mismo jefe de área
      return formatoSolicitante;

    } else {
      //enviar correo hacia el gerente de área 
      var jefeId = formatoSolicitante.solicitante.jefe;

      var userReceptor = cargarDataUsers(jefeId);
      var formatoUserReceptor = transformarData(userReceptor);
      return formatoUserReceptor;
    }

  } else if (cargo == 4) {  // Solicitud hecha por otros cargos

    var jefeId = formatoSolicitante.solicitante.jefe;
    var userReceptor = cargarDataUsers(jefeId);
    var formatoUserReceptor = transformarData(userReceptor);

    if (totalCompra <= 500) {
      //Enviar correo hacia el jefe de área
      return formatoUserReceptor;

    } else {
      //En caso de que no tenga jefe se envía al gerente de área
      if (formatoUserReceptor.solicitante.codigoCargo == 2) {
        //Enviar correo al gerente de Área
        return formatoUserReceptor;
      } else {
        //Enviar correo hacia el gerente de área
        var gerenteAreaId = formatoUserReceptor.solicitante.jefe;
        var userGerenteArea = cargarDataUsers(gerenteAreaId);
        var formatoGerenteArea = transformarData(userGerenteArea);
        return formatoGerenteArea;
      }
    }
  }
}

function obtenerCotizacionDeDrive(cotizacionUrl) {
  var fileId = cotizacionUrl.match(/[-\w]{25,}/); // Extrae el ID del archivo del enlace
  var cotizacionFile = DriveApp.getFileById(fileId);
  return cotizacionFile;
}

/*function obtenerCapexSinFirmarDeDrive(capexUrl) {
  var match = capexUrl.match(/[-\w]{25,}/); // Extrae el ID del archivo del enlace
  if (match && match[0]) { // Verifica si se encontró alguna coincidencia
    var fileId = match[0]; // Accede al primer elemento del array de coincidencias, que es el ID del archivo
    var capexFile = DriveApp.getFileById(fileId); // Obtiene el archivo usando el ID
    return capexFile;
  } else {
    throw new Error('No se pudo extraer el ID del archivo de la URL proporcionada.');
  }
} */
function obtenerCapexSinFirmarDeDrive(capexUrl) {
  var match = capexUrl ? capexUrl.match(/[-\w]{25,}/) : null; // Extrae el ID del archivo
  if (match && match[0]) {
    var fileId = match[0];
    try {
      var capexFile = DriveApp.getFileById(fileId); // Intenta obtener el archivo
      return capexFile; // Si lo encuentra, lo devuelve
    } catch (e) {
      console.warn("Archivo no encontrado en Drive para ID: " + fileId); // Mensaje en consola para depuración
      return ""; // Devuelve cadena vacía y continúa el flujo
    }
  }
  return ""; // Devuelve cadena vacía si la URL es inválida
}


// Obtener los últimos registros de una solicitud
function obtenerUltimosRegistros(solicitudId) {
  var data = obtenerDatos("index");
  var filteredData = data.reverse().filter((row) => row[0] == solicitudId);
  return filteredData;
}

//obtener los registros por id en tipo array
function obtenerRegistrosPorId(solicitudId) {
  //var libro = SpreadsheetApp.getActiveSpreadsheet();
  //var hoja = libro.getSheetByName("index");
  //var data = hoja.getDataRange().getValues();
  var data = obtenerDatos("index");
  var registros = [];

  data.forEach((row) => {
    if (row[0] == solicitudId) {
      registros.push(row);
    }
  });
  console.log("Este es el array de registros en la función obtenerRegistrosPorId: " + JSON.stringify(registros))
  return JSON.stringify(registros);

}


function seguimientoSolicitudPorId(solicitudId) {

  var registrosSolicitud = obtenerUltimosRegistros(solicitudId);
  if (!registrosSolicitud || registrosSolicitud.length === 0) {
    console.error("Error: No se encontraron registros para la solicitud ID " + solicitudId);
    return "Error: Registros no encontrados.";
  }

  var totalCompra = costoTotalSolicitud(solicitudId);

  var solicitanteEmail = registrosSolicitud[0][2] ? registrosSolicitud[0][2].trim() : null;
  if (!solicitanteEmail) {
    console.error("Error: El email del solicitante está vacío.");
    return "Error: Email del solicitante no encontrado.";
  }


  console.log("Email de la perona que realizó la solicitud: " + solicitanteEmail)

  //Obtener el solicitante de la solicitud por email
  var solicitanteData = cargarDataUsersPorEmail(solicitanteEmail);
  console.log("Usaurio consultado en la base de datos: " + solicitanteEmail)
  if (!solicitanteData) {
    console.error("Error: No se encontraron datos para el email " + solicitanteEmail);

    console.log("Este es el solicitante de la solicitud: " + JSON.stringify(aprobadores));
    return "Error: Datos no encontrados.";
  }

  var formatSolicitante = transformarData(solicitanteData);

  console.log("Este es el solicitante de la solicitud: " + JSON.stringify(formatSolicitante.solicitante));
  console.log(" Este es su nombre: " + formatSolicitante.solicitante.names + " " + formatSolicitante.solicitante.jefe)

  var aprobadores = [{
    nombreAprobador: '',
    cargo: '',
    fecha: '',
    estado: ''
  }];

  //Solicitud con un monto total menor a 500 y hecha por un empleado con tipo de cargo otros
  if (totalCompra <= 500 && formatSolicitante.solicitante.codigoCargo == 4) {

    // La solicitud debe ser aprobada por el jefe de área
    var jefeAprobador = cargarDataUsers(formatSolicitante.solicitante.jefe);
    var jefeAprobadorTrans = transformarData(jefeAprobador);

    aprobadores.push({
      nombreAprobador: jefeAprobadorTrans.solicitante.names,
      cargo: jefeAprobadorTrans.solicitante.cargo,
      fecha: registrosSolicitud[0][17] ? registrosSolicitud[0][17] : '',
      estado: registrosSolicitud[0][15]
    })

  } else if (totalCompra <= 500 && formatSolicitante.solicitante.codigoCargo != 4) {
    // Solicitud hecha por un jefe 
    aprobadores.push({
      nombreAprobador: formatSolicitante.solicitante.names,
      cargo: formatSolicitante.solicitante.cargo,
      fecha: registrosSolicitud[0][17] ? registrosSolicitud[0][17] : '',
      estado: registrosSolicitud[0][15]
    })

  } else if (totalCompra > 500 && formatSolicitante.solicitante.codigoCargo == 4) {

    // Llamamos al gerente general, pero para ello primero llamamos al jefe
    var jefeAprobador = cargarDataUsers(formatSolicitante.solicitante.jefe);
    var jefeAprobadorTrans = transformarData(jefeAprobador);

    // El jefe de área es el gerente de área
    if (jefeAprobadorTrans.solicitante.codigoCargo == 2) {
      aprobadores.push({
        nombreAprobador: jefeAprobadorTrans.solicitante.names,
        cargo: jefeAprobadorTrans.solicitante.cargo,
        fecha: registrosSolicitud[0][17] ? registrosSolicitud[0][17] : '',
        estado: registrosSolicitud[0][15]
      })
    } else {

      //Cargamos al gerente de área 
      var gerenteAprobador = cargarDataUsers(jefeAprobadorTrans.solicitante.jefe);
      var gerenteAprobadorTrans = transformarData(gerenteAprobador);
      aprobadores.push({
        nombreAprobador: gerenteAprobadorTrans.solicitante.names,
        cargo: gerenteAprobadorTrans.solicitante.cargo,
        fecha: registrosSolicitud[0][17] ? registrosSolicitud[0][17] : '',
        estado: registrosSolicitud[0][15]
      })
    }


    //Cargamos los datos del gerente general
    var gerenteGeneral = cargarDataUsers(jefeAprobadorTrans.solicitante.codigoCargo == 2 ? jefeAprobadorTrans.solicitante.jefe : gerenteAprobadorTrans.solicitante.jefe);
    var gerenteGeneralTrans = transformarData(gerenteGeneral);
    aprobadores.push({
      nombreAprobador: gerenteGeneralTrans.solicitante.names,
      cargo: gerenteGeneralTrans.solicitante.cargo,
      fecha: registrosSolicitud[0][20] ? registrosSolicitud[0][20] : '',
      estado: registrosSolicitud[0][18]
    })

  } else if (totalCompra > 500 && formatSolicitante.solicitante.codigoCargo == 3) { //Solicitud hecha por el jefe de área
    //Cargamos los datos del gerente de área
    var gerenteArea = cargarDataUsers(formatSolicitante.solicitante.jefe);
    var gerenteAreaTrans = transformarData(gerenteArea);

    aprobadores.push({
      nombreAprobador: gerenteAreaTrans.solicitante.names,
      cargo: gerenteAreaTrans.solicitante.cargo,
      fecha: registrosSolicitud[0][17] ? registrosSolicitud[0][17] : '',
      estado: registrosSolicitud[0][15]
    })

    //Cargamos los datos del gerente general
    var gerenteGeneral = cargarDataUsers(gerenteAreaTrans.solicitante.jefe);
    var gerenteGeneralTrans = transformarData(gerenteGeneral);
    aprobadores.push({
      nombreAprobador: gerenteGeneralTrans.solicitante.names,
      cargo: gerenteGeneralTrans.solicitante.cargo,
      fecha: registrosSolicitud[0][20] ? registrosSolicitud[0][20] : '',
      estado: registrosSolicitud[0][18]
    })

  } else if (totalCompra > 500 && formatSolicitante.solicitante.codigoCargo == 2) { // Solicitud hecha por un gerente de área 

    //var solicitanteData = cargarDataUsersPorEmail(solicitanteEmail);
    //var formatSolicitante = transformarData(solicitanteData); formatSolicitante
    aprobadores.push({
      nombreAprobador: formatSolicitante.solicitante.names,
      cargo: formatSolicitante.solicitante.cargo,
      fecha: registrosSolicitud[0][17] ? registrosSolicitud[0][17] : '',
      estado: registrosSolicitud[0][15]
    })

    // Cargamos al gerente general
    var gerenteGeneral = cargarDataUsers(formatSolicitante.solicitante.jefe);
    var gerenteGeneralTrans = transformarData(gerenteGeneral);
    aprobadores.push({
      nombreAprobador: gerenteGeneralTrans.solicitante.names,
      cargo: gerenteGeneralTrans.solicitante.cargo,
      fecha: registrosSolicitud[0][20] ? registrosSolicitud[0][20] : '',
      estado: registrosSolicitud[0][18]
    })
  }

  console.log("Lista de aprobadores: " + aprobadores);
  return JSON.stringify(aprobadores);
}

// FUNCIÓN PARA TRANSFORMAR EL NOMBRE DE UN APROBADOR
function convertEmailANombre(aprobadoresEmail) {
  var nombresAprobadores = '';

  for (var i = 0; i < aprobadoresEmail.length; i++) {
    var email = aprobadoresEmail[i];
    var partes = email.split('@')[0].split('.');
    var nombre = partes[0];
    var apellido = partes[1];

    // Capitalizar la primera letra del nombre y apellido
    nombre = nombre.charAt(0).toUpperCase() + nombre.slice(1);
    apellido = apellido.charAt(0).toUpperCase() + apellido.slice(1);

    nombresAprobadores += nombre + ' ' + apellido;

    // Añadir una coma y espacio si no es el último elemento
    if (i < aprobadoresEmail.length - 1) {
      nombresAprobadores += ', ';
    }
  }

  return nombresAprobadores;
}

//FUNCIÓN PARA CONVERTIR UN STRING A UN ARRAY
function convertOfStringToArray(aprobadoresEmail) {
  var arrayAprobadoresEmail = [];

  //Convertimos el string en un array
  if (typeof aprobadoresEmail === 'string') {
    arrayAprobadoresEmail = aprobadoresEmail.split(',');
  }
  return arrayAprobadoresEmail;
}

//Cargar data de usuarios 
function cargarDataUsers(userId) {
  var data = obtenerDatos("users");
  var flujoData = null;

  data.forEach((row) => {
    if (row[0] == userId) {
      flujoData = row; // Guarda el registro completo
    }
  });
  return flujoData; // Devuelve el registro encontrado o null si no se encuentra
}

//Cargar data de usuarios 
function cargarDataUsersPorEmail(userEmail) {
  var data = obtenerDatos("users");
  var flujoData = null;

  data.forEach((row) => {
    if (row[2] == userEmail) {
      console.log("SI lo encontró en la función cargarDataUsersPorEmail: " + row);
      flujoData = row; // Guarda el registro completo
    }
  });
  console.log("SI lo encontró en la función cargarDataUsersPorEmail: " + flujoData);
  return flujoData; // Devuelve el registro encontrado o null si no se encuentra
}

//Buscar el solicitante en la hoja de google sheet por id 
function buscarSolicitantePorId(solicitanteId) {
  //var data = sheet.getDataRange().getValues();
  var data = obtenerDatos("users");
  var flujoData = null;

  data.forEach((row) => {
    if (row[0] == solicitanteId) {
      flujoData = row; // Guarda el registro completo
    }
  });

  return flujoData; // Devuelve el registro encontrado o null si no se encuentra
}

//Transformar el objeto encontrado en un diccionario clave-valor
function transformarData(data) {

  if (!data || data.length === 0) {
    console.error("Error: Data inválida en transformarData.");
    return null;
  }

  // Transformar la fila en un objeto estructurado
  var formatSolicitante = {
    solicitante: {
      id: data[0],
      //username: data[1],
      //password: data[2],
      names: data[1],
      email: data[2],
      area: data[3],
      cargo: data[4],
      jefe: data[5],
      estado: data[6],
      codigoCargo: data[7]
    }
  };
  return formatSolicitante;
}

//Función que extrae la url del proyecto 
function obtenerUrl() {
  var scriptUrl = PropertiesService.getScriptProperties().getProperty('SCRIPT_URL');

  if (!scriptUrl) {
    throw new Error("La url del proyecto no está configurada. Ejecute guardarUrlScript() después del despliegue.")
  }

  return scriptUrl;
}
