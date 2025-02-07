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
  var capexPlantillaId = "1rLn1NCmShzZeHpv8A9mBKAve1yquwM4f"; // Id de la plantilla de capex
  var pdfId = "1x601kZDbaIx2nSuT2nluGpUr7W9g9i8k";      // Id de la carpeta de PDFs
  var tempId = "1F7ihMUVQZlsEyzxg90l8LBI-jzqYTJRP";     // Id de la carpeta temporal 
  var idCarpetaParaVerQR = "1Gb_2YEB1OAUwYKSjoKSmh7lZu5xhgfic";     // Id de la nueva carpeta para el PDF inicial
  var idCarpetaCapexDoc = "1KUyNuFJDyHLFqafo-Mlv9jSs2OZitSYf";

  //Para poder realizar los cambios en Google Docs
  var capexPlantilla = DriveApp.getFileById(capexPlantillaId);
  var carpetaPdf = DriveApp.getFolderById(pdfId);
  var carpetaTemp = DriveApp.getFolderById(tempId);
  var carpetaParaVerQR = DriveApp.getFolderById(idCarpetaParaVerQR);
  var carpetaCapexDoc = DriveApp.getFolderById(idCarpetaCapexDoc);

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

  // Guardar el documento para obtener el enlace
  doc.saveAndClose();
  var archivoPdf = carpetaParaVerQR.createFile(copiaPlantilla.getAs('application/pdf')).setName(nombreArchivo);
  var linkPdf = archivoPdf.getUrl();

  // Generar el código QR usando la API de goqr.me
  var qrCodeUrl = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodeURIComponent(linkPdf)}`;
  var response = UrlFetchApp.fetch(qrCodeUrl);
  var qrBlob = response.getBlob().setName("qrCode.png");

  // Volver a abrir el documento y agregar el QR
  doc = DocumentApp.openById(copiaId);
  body = doc.getBody();
  var qrParagraph = body.findText("{{codigoQR}}").getElement().getParent().asParagraph();
  qrParagraph.clear();
  qrParagraph.appendInlineImage(qrBlob);

  // Guardar el documento nuevamente como PDF
  doc.saveAndClose();
  //archivoPdf.setTrashed(true); // Eliminar el PDF inicial
  //var archivoPdfFinal = carpetaCapexDoc.createFile(copiaPlantilla.getAs('application/doc')).setName(nombreArchivo);
  //var linkPdfFinal = archivoPdfFinal.getUrl();

  var copiaDocumento = copiaPlantilla.makeCopy(nombreArchivo, carpetaCapexDoc);
  var linkCapexDocFinal = copiaDocumento.getUrl();

  // Configurar los permisos del archivo PDF final
  //archivoPdfFinal.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Eliminar la copia temporal del documento
  DriveApp.getFileById(copiaId).setTrashed(true);

  return {link: linkCapexDocFinal };
  //return { pdf: archivoPdf, link: linkPdf };

  console.log("Este es una subida de prueba")

}

