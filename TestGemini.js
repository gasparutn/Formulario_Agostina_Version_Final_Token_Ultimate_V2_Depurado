/**
 * Función de prueba para verificar la conexión con Gemini API.
 * Ejecuta esta función manualmente desde el editor de Apps Script.
 */
function testGeminiAPI() {
    try {
        // Probar con un prompt simple de texto (sin imagen)
        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${GEMINI_API_KEY}`;

        const payload = {
            contents: [{
                parts: [{
                    text: "Di 'Hola' en JSON con formato: {\"mensaje\": \"tu respuesta\"}"
                }]
            }]
        };

        const options = {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };

        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();

        Logger.log('Código de respuesta: ' + responseCode);
        Logger.log('Respuesta completa: ' + responseText);

        if (responseCode === 200) {
            return '✅ Gemini API funciona correctamente';
        } else {
            return `❌ Error ${responseCode}: ${responseText}`;
        }
    } catch (e) {
        Logger.log('Error: ' + e.toString());
        return '❌ Error: ' + e.message;
    }
}
