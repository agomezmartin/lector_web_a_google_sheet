function importarPorcentajeSP500() { 
  // Nombre de la hoja de destino
  const hojaDestino = "% SP500 profit tracker";
  
  // Obtener la hoja activa
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hojaDestino);
  if (!hoja) {
    throw new Error(`La hoja llamada "% SP500 profit tracker" no existe.`);
  }

  // Obtener la fecha actual y calcular la fecha del día anterior
  const fechaActual = new Date();
  fechaActual.setDate(fechaActual.getDate() - 1); // Restar un día para usar el día anterior
  const dia = fechaActual.getDate();
  const mes = fechaActual.getMonth(); // Enero = 0, Febrero = 1, etc.

  // Para el año 2024, tenemos que ajustar las columnas comenzando desde Noviembre (columna B)
  // Noviembre 2024 será columna B (columna 2), Diciembre columna C (columna 3), etc.
  const columnaInicio = 2; // Columna B es la número 2 (Nov)
  const filaInicio = 1; // Las filas comienzan en 1
  const columnaDestino = columnaInicio + mes - 10; // Ajuste: Noviembre es 0, Diciembre es 1, etc.
  const filaDestino = filaInicio + dia - 1; // Calcular la fila según el día (día - 1)

  // Usar IMPORTHTML para obtener datos de la web
  const url = "https://www.investing.com/funds/us-500-stock-index-inv-eur"; // Cambia esto por la URL real
  const query = "table"; // Puede ser "table" o "list"
  const indice = 10; // Índice del elemento HTML a importar

  const rangoTemporal = hoja.getRange("A45"); // Rango temporal para importar datos
  rangoTemporal.setFormula(`=IMPORTHTML("${url}", "${query}", ${indice})`);

  // Esperar unos segundos para que el valor sea importado
  Utilities.sleep(5000);

  // Leer el valor de E46 después de que la fórmula se haya ejecutado
  const rangoValor = hoja.getRange("E46");
  const valorImportado = rangoValor.getValue();

  // Eliminar el símbolo de porcentaje y convertir el valor a número
  const valorSinPorcentaje = parseFloat(valorImportado.toString().replace('%', '').replace(',', '.').trim());
  
  if (isNaN(valorSinPorcentaje)) {
    throw new Error("El valor en E46 no se ha importado correctamente. Verifica la fórmula IMPORTHTML.");
  }

  // Borrar la fórmula temporal en A45
  rangoTemporal.clearContent();

  // Pegar el valor en la celda correspondiente con formato "solo valores"
  const rangoDestino = hoja.getRange(filaDestino, columnaDestino);
  rangoDestino.setValue(valorSinPorcentaje/100);
}

function configurarTriggerPorcentaje() {
  // Eliminar triggers existentes
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "importarPorcentajeSP500") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Configurar un nuevo trigger para días laborales (martes a sábado). La web muestra el dato del día anterior.
  const diasLaborales = [ScriptApp.WeekDay.TUESDAY, ScriptApp.WeekDay.WEDNESDAY, ScriptApp.WeekDay.THURSDAY, ScriptApp.WeekDay.FRIDAY, ScriptApp.WeekDay.SATURDAY];

  diasLaborales.forEach(dia => {
    ScriptApp.newTrigger("importarPorcentajeSP500")
      .timeBased()
      .everyDays(1)
      .atHour(08) // Hora: 17:00
      .nearMinute(30) // Minuto: 45
      .create();
  });
}
