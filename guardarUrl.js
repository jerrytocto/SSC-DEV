

/*function guardarUrlScript() {
    var scriptUrl = ScriptApp.getService().getUrl(); // Obtiene la URL del script desplegado
    if (scriptUrl) {
        PropertiesService.getScriptProperties().setProperty('SCRIPT_URL', scriptUrl); // Guarda la URL
        console.log("URL guardada:", scriptUrl); // Registro de la URL en el log para confirmar
    } else {
        console.error("No se pudo obtener la URL del script. Verifica si el script está desplegado.");
    }
} */


function guardarUrlProduccion(urlProduccion) {
    if (urlProduccion) {
        PropertiesService.getScriptProperties().setProperty('SCRIPT_URL', urlProduccion);
        console.log("URL de producción guardada:", urlProduccion);
    } else {
        console.error("No se proporcionó una URL de producción válida.");
    }
}

/**
 * Guardar la URL de producción del script después de cada despliegue.
 * Reemplace el valor en 'guardarUrlActual' con la URL de la nueva versión desplegada.
 */
function guardarUrlActual() {
    guardarUrlProduccion('https://script.google.com/macros/s/AKfycbxNtrsE6xET6BL4YADOktXSMjJihOVaJLsyVZpjzbMG26rGIWRhZ4VOJvJ26mj5L7fIeg/exec');
}

