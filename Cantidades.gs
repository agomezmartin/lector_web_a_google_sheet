function importarCantidadSP500() { 
  // Nombre de la hoja de destino
  const hojaDestino = "€ SP500 profit tracker";
  
  // Obtener la hoja activa
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hojaDestino);
  if (!hoja) {
    throw new Error(`La hoja llamada "€ SP500 profit tracker" no existe.`);
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

  // Recoger los datos necesarios (URL y cookie de autenticación) para obtener el dato de la página web con la fórmula "UrlFetchApp"
  const url = "https://www.investing.com/[URL_completo]";
  
  // Cookie obtenida manualmente después de iniciar sesión
  const cookie = "ses_id=[valor_de_cookie_de_sesión]";  // Reemplaza con el valor real de la cookie

  // Realizar la solicitud HTTP para obtener el contenido de la página con autenticación
  const respuesta = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36",
      "Cookie": cookie  // Incluye el encabezado de la cookie
    }
  });

  // Obtener el contenido de la respuesta
  const html = respuesta.getContentText();

  // Buscar el valor dentro del <td> que contiene "dailyPL"
  const regex = /<td data-column-name="dailyPL"[^>]*data-value="([^"]+)"/;
  const resultado = html.match(regex);
  
  if (resultado && resultado[1]) {
    // Si encontramos el valor, lo guardamos
    const valor = resultado[1];

  // Pegar el valor en la celda correspondiente con formato "solo valores"
  const rangoDestino = hoja.getRange(filaDestino, columnaDestino);
  rangoDestino.setValue(valor);

}
}

function configurarTriggerCantidad() {
  // Eliminar triggers existentes
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "importarCantidadSP500") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Configurar un nuevo trigger para días laborales (martes a sábado). La web muestra el dato del día anterior.
  const diasLaborales = [ScriptApp.WeekDay.TUESDAY, ScriptApp.WeekDay.WEDNESDAY, ScriptApp.WeekDay.THURSDAY, ScriptApp.WeekDay.FRIDAY, ScriptApp.WeekDay.SATURDAY];

  diasLaborales.forEach(dia => {
    ScriptApp.newTrigger("importarCantidadSP500")
      .timeBased()
      .everyDays(1)
      .atHour(08) // Hora: 17:00
      .nearMinute(30) // Minuto: 45
      .create();
  });
}
