/**
 * Analiza un comprobante usando Google Gemini Vision.
 * @param {string} fileId - ID del archivo en Google Drive.
 * @param {number} precioEsperado - Monto que esperamos encontrar por persona.
 * @return {Object} Resultado estandarizado con información del análisis.
 */
function analizarComprobanteIA(fileId, precioEsperado) {
    Logger.log(`[Gemini] Iniciando análisis para FileID: ${fileId}, Esperado: ${precioEsperado}`);

    try {
        // 1. Obtener el archivo de Drive
        const file = DriveApp.getFileById(fileId);
        const mimeType = file.getMimeType();
        const blob = file.getBlob();
        const base64Data = Utilities.base64Encode(blob.getBytes());

        // Validar tipo de archivo soportado por Gemini
        if (!['image/jpeg', 'image/png', 'image/webp', 'application/pdf'].includes(mimeType)) {
            return {
                exito: false,
                error: `Formato no soportado por IA: ${mimeType}. Use JPG, PNG o PDF.`
            };
        }

        // 2. Construir el Prompt mejorado para detectar pagos múltiples
        const prompt = `
      Actúa como un sistema experto de validación de pagos.
      Analiza este comprobante de pago y extrae:
      
      1. El monto TOTAL pagado (número sin símbolos, ej: 990000)
      2. Cuántas personas/beneficiarios están incluidos en el pago
      3. Los nombres de las personas si están visibles
      
      Monto esperado por persona: ${precioEsperado}
      
      Responde ÚNICAMENTE con JSON válido (sin markdown):
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
            }]
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
            return { exito: false, error: `Error Gemini (${responseCode})` };
        }

        // 5. Procesar Respuesta
        const jsonResponse = JSON.parse(responseText);
        const candidates = jsonResponse.candidates;

        if (!candidates || candidates.length === 0) {
            return { exito: false, error: "Gemini no devolvió candidatos." };
        }

        const contentText = candidates[0].content.parts[0].text;
        Logger.log(`[Gemini] Respuesta Raw: ${contentText}`);

        // Limpiar y parsear el JSON devuelto por Gemini
        const cleanJson = contentText.replace(/```json/g, '').replace(/```/g, '').trim();
        const datos = JSON.parse(cleanJson);

        // 6. Normalizar y validar
        if (datos.monto_total !== null && datos.monto_total !== undefined) {
            // Normalizar montos
            let montoTotal = typeof datos.monto_total === 'string'
                ? parseFloat(datos.monto_total.replace(/[^0-9.-]/g, ''))
                : datos.monto_total;

            let montoPorPersona = typeof datos.monto_por_persona === 'string'
                ? parseFloat(datos.monto_por_persona.replace(/[^0-9.-]/g, ''))
                : datos.monto_por_persona;

            // Normalizar precio esperado (eliminar puntos como separadores de miles)
            let precioNormalizado = typeof precioEsperado === 'string'
                ? parseFloat(precioEsperado.replace(/\./g, '').replace(',', '.'))
                : precioEsperado;

            const cantidadPersonas = datos.cantidad_personas || 1;

            // Validar coincidencia
            const diferencia = Math.abs(montoPorPersona - precioNormalizado);
            const coincide = diferencia <= 1;

            Logger.log(`[Gemini] Total=${montoTotal}, Personas=${cantidadPersonas}, PorPersona=${montoPorPersona}, Esperado=${precioNormalizado}, Diferencia=${diferencia}, Coincide=${coincide}`);

            return {
                exito: true,
                textoEncontrado: `$${montoPorPersona}`,
                montoTotal: montoTotal,
                cantidadPersonas: cantidadPersonas,
                montoPorPersona: montoPorPersona,
                nombresDetectados: datos.nombres_detectados || [],
                coincidencia: coincide,
                esPagoMultiple: cantidadPersonas > 1,
                raw: datos
            };
        } else {
            return {
                exito: false,
                error: "No se detectó monto",
                raw: datos
            };
        }

    } catch (e) {
        Logger.log(`[Gemini] Excepción: ${e.toString()}`);
        return { exito: false, error: "Error Técnico IA: " + e.message };
    }
}
