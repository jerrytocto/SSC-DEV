/**
 * @OnlyCurrentDoc
 * @NotOnlyCurrentDoc
 */
function getOAuthToken() {
  DriveApp.getFiles();
  UrlFetchApp.fetch('https://www.google.com');
}

function generateCapex(totalCompra, registrosAprobados, aprobadoresEmail, conFirma, nombreAreaGG) {

  //Identificadores de los documentos a usar (id)
  var capexPlantillaId = "19Mw1NDbHLUGMxLnLRk9brWA6VcMj-CNaqIfT9Txg1jc"; // Id de la plantilla de capex
  var pdfId = "1oMCyoT7fnurRZDyr_EhmyKr_oOC4Yl6n";      // Id de la carpeta de PDFs
  var tempId = "1TCwbppsI3LmqYWgLltwTjypn2fYtLhY2";     // Id de la carpeta temporal
  var idCarpetaParaVerQR = "1sLK9Z_wq2Iqh_HfqJRMLxPUVXzeAPWpo";     // Id de la nueva carpeta para el PDF inicial
  var idCarpetaCapexDoc = "1KUyNuFJDyHLFqafo-Mlv9jSs2OZitSYf";

  //Para poder realizar los cambios en Google Docs
  var capexPlantilla = DriveApp.getFileById(capexPlantillaId);
  var carpetaPdf = DriveApp.getFolderById(pdfId);
  var carpetaTemp = DriveApp.getFolderById(tempId);
  var carpetaParaVerQR = DriveApp.getFolderById(idCarpetaParaVerQR);

  // Hacer una copia de la plantilla en la carpeta temporal
  var copiaPlantilla = capexPlantilla.makeCopy(carpetaTemp);
  var copiaId = copiaPlantilla.getId();
  var doc = DocumentApp.openById(copiaId);
  var body = doc.getBody();

  console.log("Listado de productus aprobados que debería ir en el capex: " + registrosAprobados);
  // Datos generales (asumiendo que los datos generales están en la primera fila)
  //var solicitanteDB = registrosAprobados[0][1];
  var razonCompra = registrosAprobados[0][3];
  var fechaRegistro = registrosAprobados[0][4];
  var justificacion = registrosAprobados[0][6];
  var prioridad = registrosAprobados[0][5];
  var solicitudId = registrosAprobados[0][0];
  var observaciones = registrosAprobados[0][14] ? registrosAprobados[0][14] : "";
  var usuarioEmail = registrosAprobados[0][2];
  var descriptionCompra = registrosAprobados[0][23] ? registrosAprobados[0][23] : "";
  var tipoCapex = registrosAprobados[0][31] ? registrosAprobados[0][31] : "";

  var solicitante = cargarDataUsersPorEmail(usuarioEmail);
  console.log("Solicitante: " + solicitante);

  if (solicitante) {
    var formatSolicitante = transformarData(solicitante);
    var mostrarUsuarioCapex = formatSolicitante.solicitante.names + " - " + formatSolicitante.solicitante.cargo;
  } else {
    var mostrarUsuarioCapex = registrosAprobados[0][1];  // Valor alternativo en caso de que solicitante no esté definido
  }

  //Obtener fecha de creación del capex
  var date = new Date();
  var dia = date.getDate();
  var mes = date.getMonth() + 1;
  var anio = date.getFullYear();
  var horaCompleta = date.getHours() + ":" + date.getMinutes();

  // Reemplazar datos generales en el documento
  var body = doc.getBody();
  //body.replaceText("{{solicitante}}", solicitante);
  body.replaceText("{{tipoDeGasto}}", tipoCapex);
  body.replaceText("{{descripcionCompra}}", descriptionCompra);
  body.replaceText("{{justificacion}}", justificacion);
  body.replaceText("{{prioridad}}", prioridad);
  body.replaceText("{{totalCompra}}", totalCompra.toFixed(2));
  body.replaceText("{{solicitante}}", mostrarUsuarioCapex);
  body.replaceText("{{solicitudId}}", solicitudId);
  body.replaceText("{{dia}}", dia);
  body.replaceText("{{mes}}", mes);
  body.replaceText("{{anio}}", anio);
  //body.replaceText("{{razonCompra}}", razonCompra.toUpperCase());
  body.replaceText("{{observaciones}}", observaciones);

  if (conFirma) {
    console.log("Se tiene que firmar")
    body.replaceText("{{gerenteGeneral}}", nombreAreaGG);
    body.replaceText("{{date}}", dia + "/" + mes + "/" + anio + "  " + horaCompleta);
    var nombreArchivo = `CAPEX_FIRM_${solicitudId}_${dia}-${mes}-${anio}.pdf`;
  } else {
    var nombreArchivo = `CAPEX_${solicitudId}_${dia}-${mes}-${anio}.pdf`;
  }

  // Crear un bloque de texto repetible
  var productosPlaceholder = "{{#productos}}";
  var startIndex = body.getText().indexOf(productosPlaceholder);
  var endIndex = body.getText().indexOf("{{/productos}}") + "{{/productos}}".length;
  var blockText = body.getText().substring(startIndex, endIndex);
  var productTemplate = blockText.replace(productosPlaceholder, "").replace("{{/productos}}", "");

  // Crear un bloque de texto repetible
  //var productContent = "";
  var descripAllProducts = "";
  for (var i = 0; i < registrosAprobados.length; i++) {
    var producto = registrosAprobados[i];
    var cantidad = producto[11];
    var centroDeCosto = producto[10];
    var equipo = producto[7];
    var marca = producto[8];
    var especificaciones = normalizarEspecificaciones(producto[9]);
    var precio = producto[12];
    var subtotal = producto[13];
    var justifyCompra = producto[22] ? producto[22] : "";

    //var productEntry = + cantidad + "  " + equipo.toUpperCase() + " " + marca.toUpperCase() + ((i + 1 < registrosAprobados.length) ? " ," : "");
    // Formato del descripAllProductsEntry
    var descripAllProductsEntry =
      `Centro de costo: ${centroDeCosto.toUpperCase()}\n` +
      `    - ${cantidad} ${equipo.toUpperCase()} ${marca.toUpperCase()} (${justifyCompra})
               ${especificaciones}\n` +
      `      Unit Price: US$: ${precio}\n` +
      `      Sub Total: US$: ${subtotal}\n` +
      `${i + 1 < registrosAprobados.length ? '--------------------------------------------------\n' : ''}`;

    //productContent += productEntry;
    descripAllProducts += descripAllProductsEntry;
  }

  // Función para normalizar las especificaciones
  function normalizarEspecificaciones(especificaciones) {
    // Reemplazar saltos de línea, comas y puntos y comas con un espacio
    return especificaciones.replace(/[\n,;]+/g, ' ').trim();
  }

  body.replaceText("{{DESCRIPTALLPRODUCTS}}", descripAllProducts);
  var urlDocumento = copiaPlantilla.getUrl();
  var qrCodeUrl = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodeURIComponent(urlDocumento)}`;

  var response = UrlFetchApp.fetch(qrCodeUrl);
  var qrBlob = response.getBlob().setName("qrCode.png");

  var qrParagraph = body.findText("{{codigoQR}}").getElement().getParent().asParagraph();
  qrParagraph.clear();
  qrParagraph.appendInlineImage(qrBlob);

  // Guardar el documento nuevamente como PDF
  doc.saveAndClose();

  return { link: urlDocumento };
}

function firmarCapex(encargadoFirmar, registros) {
  var solicitudId = registros[0][0];
  var linkCapex = registros[0][24];

  var pdfFolderId = "1oMCyoT7fnurRZDyr_EhmyKr_oOC4Yl6n";

  if (linkCapex) {
    var date = new Date();
    var dia = date.getDate();
    var mes = date.getMonth() + 1;
    var anio = date.getFullYear();
    var horaCompleta = date.getHours() + ":" + date.getMinutes();

    var docId = linkCapex.match(/[-\w]{25,}/)[0]; // Extraer ID del documento
    var doc = DocumentApp.openById(docId);
    var body = doc.getBody();

    body.replaceText("{{gerenteGeneral}}", encargadoFirmar);
    body.replaceText("{{date}}", dia + "/" + mes + "/" + anio + " " + horaCompleta);

    doc.saveAndClose();

    var file = DriveApp.getFileById(docId);
    var pdfBlob = file.getAs(MimeType.PDF);
    var nombreArchivo = `CAPEX_FIRM_${solicitudId}_${dia}-${mes}-${anio}.pdf`;
    pdfBlob.setName(nombreArchivo);

    var folder = DriveApp.getFolderById(pdfFolderId);
    var archivoPDF = folder.createFile(pdfBlob);

    return { link: archivoPDF.getUrl() };
  }
}
