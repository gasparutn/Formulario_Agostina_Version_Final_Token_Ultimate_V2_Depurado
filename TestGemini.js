/**
 * Función de prueba MANUAL para verificar Gemini API.
 * Ejecuta esta función desde el editor de Apps Script.
 */
function testGeminiAPI() {
    // ==========================================
    // CONFIGURACIÓN DE PRUEBA
    // ==========================================
    // 1. Busca un ID de archivo real en tu Google Drive (una imagen de comprobante)
    // 2. Pégalo aquí abajo:
    const FILE_ID_PRUEBA = "1ls-8yNwT7aAcOz-dkK2wlmCTvW8O2BnQ"; // ID válido que ya tenemos

    // Precio esperado para la prueba (ej. 10000)
    const PRECIO_ESPERADO = 10000;
    // ==========================================

    if (FILE_ID_PRUEBA === "PON_AQUI_TU_FILE_ID_DE_PRUEBA") {
        Logger.log("⚠️ ERROR: Debes poner un FILE ID real en la variable FILE_ID_PRUEBA");
        return;
    }

    Logger.log("🚀 Iniciando prueba manual de Gemini API...");
    Logger.log("📅 Versión del Script: ACTUALIZADA (gemini-2.5-flash)");
    Logger.log(`📂 Archivo ID: ${FILE_ID_PRUEBA}`);
    Logger.log(`💰 Precio Esperado: ${PRECIO_ESPERADO}`);

    try {
        // Llamar a la función real del servicio
        const resultado = analizarComprobanteIA(FILE_ID_PRUEBA, PRECIO_ESPERADO);

        Logger.log("---------------------------------------------------");
        Logger.log("📊 RESULTADO DE LA PRUEBA:");
        Logger.log("---------------------------------------------------");
        Logger.log(`Exito: ${resultado.exito ? "✅ SÍ" : "❌ NO"}`);

        if (resultado.error) {
            Logger.log(`❌ Error devuelto: ${resultado.error}`);
        }

        Logger.log(`💵 Monto Detectado: $${resultado.monto_total}`);
        Logger.log(`👥 Cantidad Personas: ${resultado.cantidad_personas}`);
        Logger.log(`📝 Observación: ${resultado.observacion}`);

        if (resultado.raw_text) {
            Logger.log("---------------------------------------------------");
            Logger.log("📜 Texto Crudo Detectado (primeros 200 chars):");
            Logger.log(resultado.raw_text.substring(0, 200) + "...");
        }
        Logger.log("---------------------------------------------------");

    } catch (e) {
        Logger.log(`❌ EXCEPCIÓN NO CONTROLADA: ${e.toString()}`);
    }
}
