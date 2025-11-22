/**
 * @fileoverview Contiene todas las funciones de validación centralizadas para el proyecto.
 */

/**
 * Valida que los datos esenciales para subir un comprobante estén presentes.
 * @param {string} dni - DNI del inscripto.
 * @param {object} fileData - Objeto con la información del archivo.
 * @param {string[]} cuotasSeleccionadas - Array de las cuotas que se están pagando.
 * @returns {{esValido: boolean, mensaje: string}}
 */
function validarDatosComprobante(dni, fileData, cuotasSeleccionadas) {
  if (!dni || !fileData || !cuotasSeleccionadas || cuotasSeleccionadas.length === 0) {
    return {
      esValido: false,
      mensaje: "Faltan datos (DNI, archivo o tipo de comprobante).",
    };
  }
  return { esValido: true, mensaje: "" };
}

/**
 * Valida que los datos del adulto pagador (nombre y DNI) estén presentes.
 * @param {object} datosExtras - Objeto con datos adicionales del pagador.
 * @returns {{esValido: boolean, mensaje: string}}
 */
function validarDatosPagador(datosExtras) {
  if (!datosExtras || !datosExtras.nombrePagador || !datosExtras.dniPagador) {
    return {
      esValido: false,
      mensaje: "Faltan los datos del adulto pagador (Nombre o DNI).",
    };
  }
  return { esValido: true, mensaje: "" };
}

/**
 * Valida que el DNI del pagador tenga el formato correcto (8 dígitos numéricos).
 * @param {string} dniPagador - El DNI del pagador.
 * @returns {{esValido: boolean, mensaje: string}}
 */
function validarFormatoDni(dniPagador) {
  if (!/^[0-9]{8}$/.test(dniPagador)) {
    return {
      esValido: false,
      mensaje: "El DNI debe tener 8 dígitos numéricos.",
    };
  }
  return { esValido: true, mensaje: "" };
}

/**
 * Valida si una cuota específica ya tiene un comprobante registrado.
 * @param {boolean} pagandoCuota - Flag que indica si se intenta pagar esta cuota ahora.
 * @param {string|any} comprobanteExistente - El valor de la celda del comprobante actual.
 * @param {number} numeroCuota - El número de la cuota (1, 2, o 3).
 * @returns {void}
 * @throws {Error} Si la cuota ya tiene un comprobante.
 */
function validarComprobanteExistente(pagandoCuota, comprobanteExistente, numeroCuota) {
  if (pagandoCuota && comprobanteExistente && String(comprobanteExistente).trim() !== "") {
    throw new Error(`La Cuota ${numeroCuota} ya tiene un comprobante registrado. No se puede volver a pagar.`);
  }
}

/**
 * Valida que una dirección de email sea formalmente correcta.
 * @param {string} email - La dirección de email a validar.
 * @returns {{esValido: boolean, mensaje: string}}
 */
function validarEmail(email) {
  if (!email || !email.includes("@")) {
    return {
      esValido: false,
      mensaje: "No se encontró un email válido para este registro. Por favor, contacte a la administración.",
    };
  }
  return { esValido: true, mensaje: "" };
}

/**
 * Valida que la fecha de nacimiento esté dentro del rango permitido.
 * @param {string} fechaNacimiento - La fecha en formato 'YYYY-MM-DD'.
 * @returns {{esValido: boolean, mensaje: string}}
 */
function validarFechaNacimiento(fechaNacimiento) {
  if (!fechaNacimiento || fechaNacimiento < "2010-01-01" || fechaNacimiento > "2023-12-31") {
    return {
      esValido: false,
      mensaje: "La fecha de nacimiento del inscripto debe estar entre 01/01/2010 y 31/12/2023.",
    };
  }
  return { esValido: true, mensaje: "" };
}

/**
 * Valida que el DNI de un hermano no sea igual al del inscripto principal.
 * @param {string} dniHermano - DNI limpio del hermano.
 * @param {string} dniPrincipal - DNI limpio del principal.
* @returns {{esValido: boolean, mensaje: string}}
 */
function validarDniHermanoDistinto(dniHermano, dniPrincipal) {
  if (dniHermano === dniPrincipal) {
    return {
      esValido: false,
      mensaje: "El DNI del hermano/a no puede ser igual al del inscripto principal.",
    };
  }
  return { esValido: true, mensaje: "" };
}