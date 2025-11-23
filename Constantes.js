const SPREADSHEET_ID = "1Ru-XGng2hYJbUvl-H2IA7aYQx7Ju-jk1LT1fkYOnG0w";

const NOMBRE_HOJA_BUSQUEDA = "Base de Datos";
const NOMBRE_HOJA_REGISTRO = "Registros";
const NOMBRE_HOJA_CONFIG = "Config";
const NOMBRE_HOJA_PREVENTA = "PRE-VENTA";

const FOLDER_ID_FOTOS = "1S2SbkuYdvcLFZYoHacfgwEU80kAN094l";
const FOLDER_ID_FICHAS = "1aDsTTDWHiDFUeZ8ByGp8_LY3fdzVQomu";
const FOLDER_ID_COMPROBANTES = "169EISq4RsDetQ0H3B17ViZFfe25xPcMM";

// =========================================================
// CONSTANTES "Base de Datos" (Sin cambios)
// =========================================================
const COL_HABILITADO_BUSQUEDA = 2; // Col B
const COL_NOMBRE_BUSQUEDA = 3; // Col C
const COL_APELLIDO_BUSQUEDA = 4; // Col D
const COL_FECHA_NACIMIENTO_BUSQUEDA = 5; // Col E
const COL_DNI_BUSQUEDA = 7; // Col G
const COL_OBRASOCIAL_BUSQUEDA = 8; // Col H
const COL_COLEGIO_BUSQUEDA = 9; // Col I
const COL_RESPONSABLE_BUSQUEDA = 10; // Col J
const COL_TELEFONO_BUSQUEDA = 11; // Col K

// =========================================================
// CONSTANTES "Registros" (NUEVA ESTRUCTURA - 48 COLUMNAS)
// =========================================================
const COL_NUMERO_TURNO = 1; // A: N°
const COL_MARCA_TEMPORAL = 2; // B: Marca temporal
const COL_MARCA_N_E_A = 3; // C: N(Normal) E(Extendida)
const COL_ESTADO_NUEVO_ANT = 4; // D: Inscripto
const COL_EMAIL = 5; // E: Email
const COL_NOMBRE = 6; // F: Nombres
const COL_APELLIDO = 7; // G: Apellidos
const COL_FECHA_NACIMIENTO_REGISTRO = 8; // H: Fecha de Nacimiento
const COL_GRUPOS = 9; // I: GRUPOS
const COL_DNI_INSCRIPTO = 10; // J: DNI INSCRIPTO
const COL_OBRA_SOCIAL = 11; // K: Obra Social
const COL_COLEGIO_JARDIN = 12; // L: Colegio/Jardín
const COL_ADULTO_RESPONSABLE_1 = 13; // M: Responsable 1
const COL_DNI_RESPONSABLE_1 = 14; // N: DNI Responsable 1
const COL_TEL_RESPONSABLE_1 = 15; // O: Teléfono Responsable 1
const COL_ADULTO_RESPONSABLE_2 = 16; // P: Responsable 2
const COL_DNI_RESPONSABLE_2 = 17; // Q: DNI Responsable 2
const COL_TEL_RESPONSABLE_2 = 18; // R: Teléfono Responsable 2
const COL_PERSONAS_AUTORIZADAS = 19; // S: Personas Autorizadas
const COL_PRACTICA_DEPORTE = 20; // T: Actividad extracurricular
const COL_ESPECIFIQUE_DEPORTE = 21; // U: Especifique cual actividad
const COL_TIENE_ENFERMEDAD = 22; // V: Enfermedades preexistentes
const COL_ESPECIFIQUE_ENFERMEDAD = 23; // W: Enfermedad medicaciones especiales
const COL_ES_ALERGICO = 24; // X: Alérgica/o
const COL_ESPECIFIQUE_ALERGIA = 25; // Y: Especifique la Alergia
const COL_APTITUD_FISICA = 26; // Z: Certificado Aptitud Física
const COL_FOTO_CARNET = 27; // AA: Foto Carnet
const COL_JORNADA = 28; // AB: Jornada
const COL_SOCIO = 29; // AC: SOCIO?
const COL_METODO_PAGO = 30; // AD: Método de Pago
const COL_MODO_PAGO_CUOTA = 31; // AE: Modo Pago Cuotas
const COL_PRECIO = 32; // AF: Precio total
const COL_CUOTA_1 = 33; // AG: Cuota 1
const COL_CUOTA_2 = 34; // AH: Cuota 2
const COL_CUOTA_3 = 35; // AI: Cuota 3
const COL_CANTIDAD_CUOTAS = 36; // AJ: Cant. Cuotas
const COL_ESTADO_PAGO = 37; // AK: Estado de Pago
const COL_MONTO_A_PAGAR = 38; // AL: Pagó/Acumulado
const COL_PAGADOR_NOMBRE_MANUAL = 39; // AM: Nombre y Apellido (Pagador)
const COL_PAGADOR_DNI_MANUAL = 40; // AN: DNI Pagador
const COL_COMPROBANTE_MANUAL_TOTAL_EXT = 41; // AO: Comprobante 1 pago individual o Familiar
const COL_COMPROBANTE_MANUAL_CUOTA1 = 42; // AP: Comprobante C1
const COL_COMPROBANTE_MANUAL_CUOTA2 = 43; // AQ: Comprobante C2
const COL_COMPROBANTE_MANUAL_CUOTA3 = 44; // AR: Comprobante C3
const COL_ENVIAR_EMAIL_MANUAL = 45; // AS: Enviar Email
const COL_FECHA_HORA = 46; // AT: Fecha y Hora
const COL_RESERVA_1 = 47; // AU: Reserva 1
const COL_VINCULO_PRINCIPAL = 48; // AV: VÍNCULO

// =========================================================
// CONSTANTES "PRE-VENTA" (Sin cambios)
// =========================================================
const COL_PREVENTA_EMAIL = 3; // Col C
const COL_PREVENTA_NOMBRE = 4; // Col D
const COL_PREVENTA_APELLIDO = 5; // Col E
const COL_PREVENTA_DNI = 6; // Col F
const COL_PREVENTA_FECHA_NAC = 7; // Col G
const COL_PREVENTA_GUARDA = 9; // Col I

// =========================================================
// CONSTANTES GEMINI AI
// =========================================================
const GEMINI_API_KEY = "AIzaSyC_dW30ZJzaAt1dwvneSg9WSpV1HRuUHUg";
const GEMINI_MODEL = "gemini-2.5-flash";
