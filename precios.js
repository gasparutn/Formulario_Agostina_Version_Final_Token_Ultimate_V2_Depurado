/**
 * Archivo: precios.gs
 * Contiene la lógica centralizada para calcular los precios de inscripción
 * basado en la grilla de la hoja "Config".
 *
 * (MODIFICADO v7 - LÓGICA DUAL)
 * - (MODIFICACIÓN CLAVE) Se implementa la lógica de precios dual:
 * 1. (CASO PRE-VENTA): Si 'esPreventa' es true (Filas 59, 69), el precio de la celda
 * es el TOTAL FAMILIAR y SE DIVIDE por 'totalHijos'.
 * 2. (CASO GENERAL): Si 'esPreventa' es false (Filas 17, 27, 37, 47), el precio
 * de la celda es el PRECIO INDIVIDUAL y SE TOMA DIRECTAMENTE.
 *
 * - 'obtenerPrecioYConfiguracion' ahora acepta 'totalHijos'.
 * - '_obtenerIndiceHijo' sigue devolviendo el índice y el total.
 */

/**
 * Función principal para obtener el precio basado en la grilla de Config.
 * (MODIFICADA) Ahora acepta 'totalHijos' para calcular el precio individual.
 *
 * @param {object} datos - Objeto con los datos del inscripto.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hojaConfig - La hoja de "Config" ya abierta.
 * @param {number} indiceHijo - El índice de precio a usar (0, 1, o 2).
 * @param {number} totalHijos - El conteo total de hijos (para la división).
 * @returns {{precio: number, montoAPagar: number, cantidadCuotas: number, valorCuota: number}}
 */
function obtenerPrecioYConfiguracion(
  datos,
  hojaConfig,
  indiceHijo = 0,
  totalHijos = 1
) {
  // --- 1. Extraer datos ---
  const esPreventa = datos.esPreventa === true || datos.esPreventa === "true";
  const jornada = datos.jornada;
  const metodoPago = datos.metodoPago;
  const esSocio = datos.esSocio === "SÍ";

  let tipoPagoParaPrecio = "";
  if (metodoPago === "Pago en Cuotas" && !esPreventa) {
    tipoPagoParaPrecio = datos.subMetodoCuotas || "Transferencia";
  } else {
    tipoPagoParaPrecio = metodoPago;
  }

  Logger.log(
    `Calculando precio para indiceHijo: ${indiceHijo}, totalHijos: ${totalHijos}. Jornada=${jornada}, Metodo=${metodoPago}, Socio=${esSocio}, esPreventa=${esPreventa}`
  );

  let precioDeTabla = 0; // El valor EXACTO de la celda
  let celdaPrecio = "";
  let baseRow = 0;
  let precioEsCuota = false; // TRUE si la celda contiene el valor de 1 cuota; FALSE si contiene el total familiar.
  let cantidadCuotas = 0;

  try {
    // --- 2. Lógica de Mapeo de Columnas (E, F, G, H) ---
    let colLetra = "";
    if (esSocio) {
      if (tipoPagoParaPrecio.includes("Efectivo")) {
        colLetra = "E";
      } else {
        colLetra = "F";
      }
    } else {
      // No Socio
      if (tipoPagoParaPrecio.includes("Efectivo")) {
        colLetra = "G";
      } else {
        colLetra = "H";
      }
    }

    // --- 3. Lógica de Mapeo de Filas (AQUÍ ESTÁ EL CAMBIO) ---
    if (esPreventa) {
      // --- LÓGICA DE PRE-VENTA ---
      // La tabla (filas 59/69) almacena el PRECIO TOTAL FAMILIAR.
      precioEsCuota = false;
      cantidadCuotas = 1; // Pre-Venta siempre es 1 pago

      if (jornada === "Jornada Normal") {
        baseRow = 59;
      } else {
        baseRow = 69;
      }
      Logger.log(`Lógica PRE-VENTA. Fila base: ${baseRow}.`);
    } else {
      // --- LÓGICA GENERAL (EXISTENTE) ---
      if (jornada === "Jornada Normal") {
        if (metodoPago === "Pago en Cuotas") {
          baseRow = 37;
          precioEsCuota = true; // La tabla (37) almacena el valor de *una* cuota (ya individual)
          cantidadCuotas = 3;
        } else {
          baseRow = 17;
          precioEsCuota = false; // La tabla (17) almacena el *precio individual por nivel*
          cantidadCuotas = 1;
        }
      } else {
        // Asumimos "Jornada Normal extendida"
        if (metodoPago === "Pago en Cuotas") {
          baseRow = 47;
          precioEsCuota = true;
          cantidadCuotas = 3;
        } else {
          baseRow = 27;
          precioEsCuota = false;
          cantidadCuotas = 1;
        }
      }
      Logger.log(`Lógica GENERAL. Fila base: ${baseRow}`);
    }

    // --- 4. Determina la fila exacta según el índice del hijo ---
    let rowNum = baseRow + indiceHijo; // indiceHijo ya es 0, 1, o 2.

    // --- 5. Obtener el Precio de la Celda ---
    if (colLetra && rowNum > 0) {
      celdaPrecio = colLetra + rowNum;
      Logger.log(
        `Obteniendo precio de la celda de Config: ${celdaPrecio}. Es valor/cuota? ${precioEsCuota}`
      );

      const valorCelda = hojaConfig.getRange(celdaPrecio).getValue();

      if (typeof valorCelda === "number") {
        precioDeTabla = valorCelda;
      } else if (typeof valorCelda === "string") {
        const precioLimpio = valorCelda
          .replace(/[$.]/g, "")
          .split(" ")[0]
          .replace(/,/g, ".");
        precioDeTabla = parseFloat(precioLimpio) || 0;
      }
    } else {
      Logger.log(
        `No se pudo determinar la celda. colLetra=${colLetra}, rowNum=${rowNum}`
      );
    }
  } catch (e) {
    Logger.log(
      `Error en obtenerPrecioYConfiguracion: ${e.message}. Celda: ${celdaPrecio}. Stack: ${e.stack}`
    );
    return { precio: 0, montoAPagar: 0, cantidadCuotas: 0, valorCuota: 0 };
  }

  // =========================================================
  // --- 6. (MODIFICACIÓN CLAVE) Lógica de Precios DUAL ---
  // =========================================================

  let precioTotal = 0; // El precio individual final
  let valorCuota = 0;
  let montoAPagarInicial = 0;
  const divisor = Math.max(totalHijos, 1); // Asegura que nunca se divida por 0

  if (precioEsCuota) {
    // CASO 1: General - Pago en Cuotas (Filas 37, 47)
    // precioDeTabla es el valor de 1 cuota individual. NO se divide.
    valorCuota = precioDeTabla;
    precioTotal = precioDeTabla * 3;
    montoAPagarInicial = 0;
    cantidadCuotas = 3;
  } else if (!precioEsCuota && esPreventa) {
    // CASO 2: Pre-Venta (Filas 59, 69)
    // precioDeTabla es el TOTAL FAMILIAR. Debe dividirse.
    const precioTotalFamilia = precioDeTabla;
    precioTotal = parseFloat((precioTotalFamilia / divisor).toFixed(2)); // Precio individual
    montoAPagarInicial = precioTotal;
    cantidadCuotas = 1;
    valorCuota = 0;
  } else {
    // (!precioEsCuota && !esPreventa)
    // CASO 3: General - Pago Único (Filas 17, 27)
    // precioDeTabla es el PRECIO INDIVIDUAL (por nivel de hermano). NO se divide.
    precioTotal = precioDeTabla; // Tomar el precio directo
    montoAPagarInicial = precioTotal;
    cantidadCuotas = 1;
    valorCuota = 0;
  }

  // (Lógica de seguridad de la v17)
  // Si el frontend falló y mandó 'Pago en Cuotas' para Pre-Venta, forzarlo a 1 pago.
  if (esPreventa) {
    cantidadCuotas = 1;
    valorCuota = 0;
    // El precioTotal ya fue calculado (dividido) correctamente en el CASO 2.
    montoAPagarInicial = precioTotal;
  }

  Logger.log(
    `Precio Total (AE) [Individual]: ${precioTotal}, Monto a Pagar (AK): ${montoAPagarInicial}, Valor Cuota (AF/AG/AH): ${valorCuota}, Cuotas: ${cantidadCuotas}`
  );

  return {
    precio: precioTotal, // Col AE (Precio Total Individual)
    montoAPagar: montoAPagarInicial, // Col AK (Monto a Pagar)
    cantidadCuotas: cantidadCuotas,
    valorCuota: valorCuota, // (Propiedad para Pagos.js)
  };
}

/**
 * Función helper para encontrar el índice de un hermano (0, 1, 2+)
 * cuando está actualizando sus datos.
 *
 * (MODIFICADO v16) Ahora devuelve un objeto { indice, total }.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hojaRegistro - La hoja de "Registros".
 * @param {number} fila - La fila del DNI que se está actualizando.
 * @returns {{indice: number, total: number}} - El índice de precio (0,1,2) y el total de miembros.
 */
function _obtenerIndiceHijo(hojaRegistro, fila) {
  try {
    const vinculoPrincipal = hojaRegistro
      .getRange(fila, COL_VINCULO_PRINCIPAL)
      .getValue();

    // Si no tiene vínculo (o es el principal sin hermanos), es el índice 0, total 1.
    if (!vinculoPrincipal) {
      return { indice: 0, total: 1 };
    }

    // Si tiene vínculo, buscar a toda la familia
    const rangoVinculos = hojaRegistro.getRange(
      2,
      COL_VINCULO_PRINCIPAL,
      hojaRegistro.getLastRow() - 1,
      1
    );
    const todasLasCeldas = rangoVinculos
      .createTextFinder(vinculoPrincipal)
      .matchEntireCell(true)
      .findAll();

    // Contar el número total de miembros en la familia
    const totalMiembros = todasLasCeldas.length;

    if (totalMiembros <= 0) {
      Logger.log(
        `Error en _obtenerIndiceHijo: Se encontró vínculo ${vinculoPrincipal} pero el conteo es 0.`
      );
      return { indice: 0, total: 1 }; // Fallback seguro
    }

    // =========================================================
    // --- Lógica de índice (v16) ---
    // Devolvemos el índice basado en el conteo total.
    // Si totalMiembros = 1, índice = 0 (Fila 1)
    // Si totalMiembros = 2, índice = 1 (Fila 2)
    // Si totalMiembros = 3, índice = 2 (Fila 3)
    // =========================================================
    const indicePrecio = Math.min(totalMiembros - 1, 2);

    Logger.log(
      `_obtenerIndiceHijo: Fila ${fila} pertenece a familia ${vinculoPrincipal}. Miembros totales: ${totalMiembros}. Índice de precio aplicado: ${indicePrecio}`
    );

    return { indice: indicePrecio, total: totalMiembros };
  } catch (e) {
    Logger.log(`Error en _obtenerIndiceHijo: ${e.message}`);
    return { indice: 0, total: 1 }; // Fallback seguro
  }
}
