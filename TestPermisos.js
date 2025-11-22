/**
 * Función de prueba para forzar la solicitud de permisos de UrlFetchApp.
 * Ejecuta esta función manualmente desde el editor de Apps Script.
 * Te pedirá autorizar el permiso: https://www.googleapis.com/auth/script.external_request
 */
function testPermisosUrlFetch() {
    try {
        // Hacer una llamada simple a una API pública para forzar el permiso
        const response = UrlFetchApp.fetch('https://www.google.com');
        Logger.log('✅ Permisos de UrlFetchApp autorizados correctamente');
        Logger.log('Código de respuesta: ' + response.getResponseCode());
        return '✅ Permisos autorizados. Ahora puedes usar el OCR con Gemini.';
    } catch (e) {
        Logger.log('❌ Error: ' + e.toString());
        return '❌ Error al autorizar permisos: ' + e.message;
    }
}
