/**123
 * Archivo: CalculaEdadGrupoColor.js
 * Contiene la lógica para calcular grupos por FECHA DE CORTE y asignar colores.
 */

/**
 * Determina el grupo basado en una fecha de corte fija (30 de junio).
 * @param {string} fechaNacStr - La fecha de nacimiento (ej. "2017-10-20").
 * @returns {string} - El texto del grupo (ej. "Grupo 8 años").
 */
function obtenerGrupoPorFechaNacimiento(fechaNacStr) {
  if (!fechaNacStr) return "Sin Fecha";

  try {
    // Se usa 'T00:00:00Z' para asegurar que la fecha se interprete como UTC y evitar problemas de zona horaria.
    const fechaNac = new Date(fechaNacStr + "T00:00:00Z");
    const mesCorte = 5; // Junio (los meses en JS van de 0 a 11)
    const diaCorte = 30;

    if (fechaNac >= new Date(Date.UTC(2022, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2023, mesCorte, diaCorte))) return "Grupo 3 años";
    if (fechaNac >= new Date(Date.UTC(2021, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2022, mesCorte, diaCorte))) return "Grupo 4 años";
    if (fechaNac >= new Date(Date.UTC(2020, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2021, mesCorte, diaCorte))) return "Grupo 5 años";
    if (fechaNac >= new Date(Date.UTC(2019, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2020, mesCorte, diaCorte))) return "Grupo 6 años";
    if (fechaNac >= new Date(Date.UTC(2018, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2019, mesCorte, diaCorte))) return "Grupo 7 años";
    if (fechaNac >= new Date(Date.UTC(2017, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2018, mesCorte, diaCorte))) return "Grupo 8 años";
    if (fechaNac >= new Date(Date.UTC(2016, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2017, mesCorte, diaCorte))) return "Grupo 9 años";
    if (fechaNac >= new Date(Date.UTC(2015, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2016, mesCorte, diaCorte))) return "Grupo 10 años";
    if (fechaNac >= new Date(Date.UTC(2014, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2015, mesCorte, diaCorte))) return "Grupo 11 años";
    if (fechaNac >= new Date(Date.UTC(2013, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2014, mesCorte, diaCorte))) return "Grupo 12 años";
    if (fechaNac >= new Date(Date.UTC(2012, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2013, mesCorte, diaCorte))) return "Grupo 13 años";
    if (fechaNac >= new Date(Date.UTC(2011, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2012, mesCorte, diaCorte))) return "Grupo 14 años";
    if (fechaNac >= new Date(Date.UTC(2010, mesCorte, diaCorte)) && fechaNac < new Date(Date.UTC(2011, mesCorte, diaCorte))) return "Grupo 15 años";

    return "Fuera de rango";
  } catch (e) {
    Logger.log("Error al parsear fecha en obtenerGrupoPorFechaNacimiento: " + fechaNacStr + " | Error: " + e.message);
    return "Error Fecha";
  }
}

/**
 * Aplica el color de fondo a la celda del grupo según la configuración.
 * @param {Sheet} hoja - La hoja de cálculo donde se aplicará el color.
 * @param {number} fila - El número de fila a colorear.
 * @param {string} textoGrupo - El nombre del grupo (ej. "Grupo 8 años").
 * @param {Sheet} hojaConfig - La hoja de configuración que contiene los colores.
 */
function aplicarColorGrupo(hoja, fila, textoGrupo, hojaConfig) {
  try {
    // El rango donde están definidos los grupos y sus colores en la hoja 'Config'
    const rangoGrupos = hojaConfig.getRange("A30:B41");
    const valoresGrupos = rangoGrupos.getValues();
    const coloresGrupos = rangoGrupos.getBackgrounds();

    for (let i = 0; i < valoresGrupos.length; i++) {
      // Compara el texto del grupo de la hoja con el texto del grupo calculado
      if (valoresGrupos[i][0].toString().trim() == textoGrupo.toString().trim()) {
        const color = coloresGrupos[i][1]; // El color está en la segunda columna del rango
        hoja.getRange(fila, COL_GRUPOS).setBackground(color);
        return; // Termina la función una vez que encuentra y aplica el color
      }
    }
    // Si el bucle termina sin encontrar el grupo, registra un aviso.
    Logger.log(`No se encontró color para el grupo: "${textoGrupo}" en la hoja Config!A30:A41`);
  } catch (e) {
    Logger.log(`Error al aplicar color para el grupo ${textoGrupo} en fila ${fila}: ${e.message}`);
  }
}
