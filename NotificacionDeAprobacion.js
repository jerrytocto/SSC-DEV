
// Función para verificar las solicitudes pendientes de aprobación
function verificarSolicitudes() {
    var solicitudes = obtenerTodasSolicitudes();
    var solicitudesProcesadas = new Set();

    for (var i = 0; i < solicitudes.length; i++) {
        var idSolicitud = solicitudes[i][0];
        if (!solicitudesProcesadas.has(idSolicitud)) {
            if (solicitudes[i][15] == "Pendiente") {
                procesarSolicitud(solicitudes[i], 15, i + 1);
            } else if (solicitudes[i][18] == "Pendiente" && solicitudes[i][15] != "Pendiente") {
                procesarSolicitud(solicitudes[i], 18, i + 1);
            }
            solicitudesProcesadas.add(idSolicitud);
        }
    }
}

function procesarSolicitud(solicitud, columnaEstado, rowIndex) {
    //var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("index");

    var sheet = conectarHoja("index");
    var fechaActual = new Date();
    var fechaComparacion, fechaUltimaNotificacion, contadorNotificaciones;

    if (columnaEstado == 15) {
        // Estado pendiente en la columna 15 (aprobación por el gerente de área)
        fechaComparacion = new Date(solicitud[4]); // Fecha de emisión de la solicitud
        fechaUltimaNotificacion = solicitud[26] || null; // Columna 26, índice 25
        contadorNotificaciones = solicitud[27] || 0;     // Columna 26, índice 25
    } else if (columnaEstado == 18) {
        // Estado pendiente en la columna 18 (aprobación por el gerente general)
        fechaComparacion = new Date(solicitud[17]); // Fecha de aprobación por el gerente de área
        fechaUltimaNotificacion = solicitud[28] || null; // Columna 27, índice 26
        contadorNotificaciones = solicitud[29] || 0;     // Columna 28, índice 27
    }

    // Determinar si se debe enviar una notificación
    if (!fechaUltimaNotificacion) {
        // Primera notificación
        if (calcularTiempoTranscurrido(fechaComparacion, fechaActual)) {
            enviarNotificacion(solicitud, columnaEstado);
            actualizarRegistroNotificacion(sheet, rowIndex, fechaActual, 1, columnaEstado);
        }
    } else {
        // Notificaciones subsiguientes
        if (calcularTiempoTranscurrido(new Date(fechaUltimaNotificacion), fechaActual)) {
            enviarNotificacion(solicitud, columnaEstado);
            actualizarRegistroNotificacion(sheet, rowIndex, fechaActual, contadorNotificaciones + 1, columnaEstado);
        }
    }
}

function actualizarRegistroNotificacion(sheet, rowIndex, fechaActual, contador, columnaEstado) {
    var fechaNotificacionCol, contadorCol;

    if (columnaEstado == 15) {
        fechaNotificacionCol = 27; // Columna para "Fecha Última Notificación (Gerente Área)"
        contadorCol = 28;          // Columna para "Contador Notificaciones (Gerente Área)"
    } else if (columnaEstado == 18) {
        fechaNotificacionCol = 29; // Columna para "Fecha Última Notificación (Gerente General)"
        contadorCol = 30;          // Columna para "Contador Notificaciones (Gerente General)"
    }

    // Verificar y crear encabezados si las columnas no tienen encabezados
    if (sheet.getRange(1, fechaNotificacionCol).isBlank()) {
        sheet.getRange(1, fechaNotificacionCol).setValue(
            columnaEstado == 15 ? "Fecha Última Notificación (Gerente Área)" : "Fecha Última Notificación (Gerente General)"
        );
    }

    if (sheet.getRange(1, contadorCol).isBlank()) {
        sheet.getRange(1, contadorCol).setValue(
            columnaEstado == 15 ? "Contador Notificaciones (Gerente Área)" : "Contador Notificaciones (Gerente General)"
        );
    }

    // Obtener el ID de la solicitud actual
    var solicitudId = sheet.getRange(rowIndex, 1).getValue(); // Suponiendo que el ID está en la primera columna

    // Recorrer las filas para encontrar todas las filas con el mismo ID de solicitud
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();

    for (var i = rowIndex - 1; i < values.length; i++) {
        if (values[i][0] == solicitudId) {
            // Actualizar la fecha de notificación y el contador de notificaciones
            sheet.getRange(i + 1, fechaNotificacionCol).setValue(fechaActual);
            sheet.getRange(i + 1, contadorCol).setValue(contador);
        }
    }
}


function calcularTiempoTranscurrido(fechaAnterior, fechaVerificacion) {
    var diferencia = fechaVerificacion.getTime() - fechaAnterior.getTime();
    var diasTranscurridos = diferencia / (1000 * 60 * 60 * 24);
    return diasTranscurridos >= 2; // 2 días para todas las notificaciones
}


function enviarNotificacion(solicitud, columnaEstado) {
    var solicitudId = solicitud[0];
    console.log("Verificando solicitud con ID en la función enviarNotificación:", solicitudId);

    var fechaEmisionSolicitud = columnaEstado == 15 ? solicitud[4] : solicitud[17];
    var registrosPorSolicitud = obtenerUltimosRegistros(solicitudId);

    enviarEmailDeAviso(registrosPorSolicitud, columnaEstado);
}

function crearDisparadorDiario() {
    eliminarDisparadoresExistentes();
    ScriptApp.newTrigger('verificarSolicitudes')
        .timeBased()
        .everyDays(1)
        .atHour(5)
        .nearMinute(0)
        .create();
    console.log("Disparador diario creado exitosamente para ejecutarse a las 5:00 a.m.");
}

function eliminarDisparadoresExistentes() {
    var disparadores = ScriptApp.getProjectTriggers();
    for (var i = 0; i < disparadores.length; i++) {
        if (disparadores[i].getHandlerFunction() === 'verificarSolicitudes') {
            ScriptApp.deleteTrigger(disparadores[i]);
        }
    }
}

//Función para obtener las solicitudes por id de la función ultimos registros del archivo index
function obtenerTodasSolicitudes() {
    // Retorna todas las solicitudes de la hoja index
    //var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("index");
    //var data = sheet.getDataRange().getValues();
    var data = obtenerDatos("index");
    return data; // Retorna todas las solicitudes
}

//Función para enviar el email de aviso
function enviarEmailDeAviso(registrosPorSolicitud, columnaEstado) {

    // Obtener el id de la solicitud
    var idSolicitud = registrosPorSolicitud[0][0];

    // Obtener el costo total de la solicitud 
    var costoTotal = costoTotalSolicitud(idSolicitud);

    // Obtener el solicitante por email
    var solicitante = cargarDataUsersPorEmail(registrosPorSolicitud[0][2]);

    // Validar que el solicitante en la hoja users 
    if (solicitante == null) {
        console.log("El solicitante no se encuentra en la hoja users");
        return;
    }

    // Transformar la data del solicitante 
    var formatSolicitante = transformarData(solicitante);

    if (columnaEstado == 15) {
        // Enviar correo al jefe o gerente de área 
        enviarEmail(costoTotal, idSolicitud, formatSolicitante, true);

    } else if (columnaEstado == 18) {

        // Obtener el link de la cotización 
        var cotizacion = registrosPorSolicitud[0][21] !="" ? registrosPorSolicitud[0][21] : "";

        // Verificar quien es la persona que realizó la primera aprobación 
        var aprobador = determinarDestinatario(costoTotal, formatSolicitante);

        // Obtener el id del siguiente aprobador 
        var siguienteAprobadorId = aprobador.solicitante.jefe;
        var siguienteAprobador = cargarDataUsers(siguienteAprobadorId);
        var formatSiguienteAprobador = transformarData(siguienteAprobador);

        // Tomar el email del siguiente aprobador
        var emailSiguienteAprobador = formatSiguienteAprobador.solicitante.email;

        enviarCorreoGerenteGeneral(
            registrosPorSolicitud,
            aprobador,
            emailSiguienteAprobador,
            costoTotal,
            cotizacion,
            true
        )
    }
}


