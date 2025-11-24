/**
 * Analiza un comprobante usando Google Gemini Vision.
 * @param {string} fileId - ID del archivo en Google Drive.
 * @param {number} precioEsperado - Monto que esperamos encontrar por persona.
 * @return {Object} Resultado estandarizado con información del análisis.
 */
function analizarComprobanteIA(fileId, precioEsperado) {
    Logger.log(`[Gemini] Iniciando análisis para FileID: ${fileId}, Esperado: ${precioEsperado}`);
    Logger.log(`[Gemini] Usando Modelo: ${GEMINI_MODEL}`);

    try {
        // 1. Obtener el archivo de Drive
        const file = DriveApp.getFileById(fileId);
        const mimeType = file.getMimeType();
        const blob = file.getBlob();
        const fileSize = blob.getBytes().length;

        // Validar tamaño (máximo 20MB para Gemini)
        const MAX_SIZE = 20 * 1024 * 1024; // 20MB
        if (fileSize > MAX_SIZE) {
            Logger.log(`[Gemini] Archivo demasiado grande: ${fileSize} bytes (máx: ${MAX_SIZE})`);
            return {
                exito: false,
                error: `Archivo demasiado grande (${Math.round(fileSize / 1024 / 1024)}MB). Máximo: 20MB`
            };
        }

        const base64Data = Utilities.base64Encode(blob.getBytes());
        Logger.log(`[Gemini] Archivo: ${mimeType}, Tamaño: ${Math.round(fileSize / 1024)}KB`);

        // Validar tipo de archivo soportado por Gemini
        if (!['image/jpeg', 'image/png', 'image/webp', 'application/pdf'].includes(mimeType)) {
            return {
                exito: false,
                error: `Formato no soportado por IA: ${mimeType}. Use JPG, PNG o PDF.`
            };
        }

        // 2. Construir el Prompt mejorado para detectar pagos múltiples
        const prompt = `
      Analiza este comprobante de pago (transferencia bancaria).
      El precio unitario esperado por persona es: $${precioEsperado}.
      
      Busca el monto total transferido y determina si corresponde a 1 persona o a un grupo familiar (múltiplos del precio esperado).
      
      Responde EXCLUSIVAMENTE con un objeto JSON con esta estructura:
      {
        "monto_total": number,
        "cantidad_personas": number,
        "monto_por_persona": number,
        "nombres_detectados": [],
        "moneda": "ARS",
        "observacion": "texto explicativo"
      }
      
      Si detectas 2 o más personas, divide el monto_total por cantidad_personas.
      Si solo ves 1 persona o no hay indicios, cantidad_personas = 1.
      
      IMPORTANTE: El campo "observacion" debe estar escrito SIEMPRE en ESPAÑOL.
    `;

        // 3. Configurar la petición a Gemini API
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;

        const payload = {
            contents: [{
                parts: [
                    { text: prompt },
                    {
                        inline_data: {
                            mime_type: mimeType,
                            data: base64Data
                        }
                    }
                ]
            }],
            generationConfig: {
                response_mime_type: "application/json"
            }
        };

        const options = {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };

        // 4. Llamar a la API
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();

        if (responseCode !== 200) {
            Logger.log(`[Gemini] Error API (${responseCode}): ${responseText}`);

            // Logging adicional para error 400
            if (responseCode === 400) {
                Logger.log(`[Gemini] Detalles del error 400:`);
                Logger.log(`  - Modelo: ${GEMINI_MODEL}`);
                Logger.log(`  - MimeType: ${mimeType}`);
                Logger.log(`  - Tamaño Base64: ${base64Data.length} caracteres`);
                Logger.log(`  - URL: ${url.substring(0, 100)}...`);

                try {
                    const errorDetail = JSON.parse(responseText);
                    Logger.log(`  - Error detallado: ${JSON.stringify(errorDetail)}`);
                } catch (e) {
                    Logger.log(`  - Respuesta raw: ${responseText.substring(0, 500)}`);
                }
            }

            return { exito: false, error: `Error Gemini (${responseCode})` };
        }

        // 5. Procesar Respuesta
        const jsonResponse = JSON.parse(responseText);

        // Verificar estructura de respuesta de Gemini
        if (!jsonResponse.candidates || jsonResponse.candidates.length === 0 || !jsonResponse.candidates[0].content) {
            Logger.log(`[Gemini] Respuesta vacía o estructura inesperada: ${responseText}`);
            return { exito: false, error: "Gemini no devolvió candidatos." };
        }

        const rawText = jsonResponse.candidates[0].content.parts[0].text;
        Logger.log(`[Gemini] Respuesta Raw: ${rawText}`);

        // Limpiar bloques de código markdown si existen (```json ... ```)
        const jsonString = rawText.replace(/```json/g, '').replace(/```/g, '').trim();
        const resultado = JSON.parse(jsonString);

        // Normalizar datos numéricos (quitar separadores de miles si vienen como string)
        if (typeof resultado.monto_total === 'string') {
            resultado.monto_total = parseFloat(resultado.monto_total.replace(/\./g, '').replace(',', '.'));
        }

        // Devolver datos crudos para que Comprobantes.js aplique la lógica de negocio
        return {
            exito: true, // Siempre true si Gemini respondió JSON válido
            monto_total: resultado.monto_total,
            cantidad_personas_detectadas: resultado.cantidad_personas || 1,
            nombres_detectados: resultado.nombres_detectados || [],
            moneda: resultado.moneda || "ARS",
            observacion: resultado.observacion || "Análisis completado.",
            raw_text: rawText
        };

    } catch (e) {
        Logger.log(`[Gemini] Error CRÍTICO: ${e.toString()}`);
        return {
            exito: false,
            error: `Error interno: ${e.toString()}`
        };
    }
}
