function actualizarFormulario() {
  // IDs del archivo de Google Sheets y del formulario
  const spreadsheetId = "1Hkv8CLofSZPFfTty058hJX0sZGU5r66P_oi7j1tPFMo";
  const sheetName = "activosActualizar";
  const sheetName2 = "activosActualizar2";
  const formId = "1DCr5AjmEqk7dfLA1BzRAmKyhWzwjAMuwjbmugUKB8Qk";

  // Mensaje por defecto si no hay datos
  const mensajeDefault = ["No hay pacientes para programar, reprogramar o cancelar"];

  // Obtener las opciones de la primera hoja
  const opcionesSheet1 = (() => {
    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`La hoja ${sheetName} no existe.`);
      return mensajeDefault;
    }

    const lastRow = sheet.getLastRow(); // Última fila con datos
    if (lastRow < 2) {
      Logger.log(`No hay datos en la hoja ${sheetName}.`);
      return mensajeDefault;
    }

    const dataRange = sheet.getRange(2, 2, lastRow - 1); // Columna B desde la fila 2
    const values = dataRange.getValues().flat(); // Obtener los valores como array

    return values.length > 0 ? values : mensajeDefault;
  })();

  // Obtener las opciones de la segunda hoja
  const opcionesSheet2 = (() => {
    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName2);
    if (!sheet) {
      Logger.log(`La hoja ${sheetName2} no existe.`);
      return mensajeDefault;
    }

    const lastRow = sheet.getLastRow(); // Última fila con datos
    if (lastRow < 2) {
      Logger.log(`No hay datos en la hoja ${sheetName2}.`);
      return mensajeDefault;
    }

    const dataRange = sheet.getRange(2, 2, lastRow - 1); // Columna B desde la fila 2
    const values = dataRange.getValues().flat(); // Obtener los valores como array

    return values.length > 0 ? values : mensajeDefault;
  })();

  // Actualizar el formulario
  const form = FormApp.openById(formId);

  // Función interna para actualizar una pregunta
  const actualizarPregunta = (questionId, opciones) => {
    const items = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE);
    const targetQuestion = items.find(item => item.getId() === questionId);
    if (!targetQuestion) {
      Logger.log(`No se encontró la pregunta con ID ${questionId}.`);
      return;
    }

    const question = targetQuestion.asMultipleChoiceItem();
    question.setChoiceValues(opciones);
  };

  // Actualizar la pregunta correspondiente a la primera hoja
  actualizarPregunta(797739091, opcionesSheet1);

  // Actualizar las preguntas correspondientes a la segunda hoja
  actualizarPregunta(2140363120, opcionesSheet2);
  actualizarPregunta(1642252810, opcionesSheet2);

  Logger.log("Opciones actualizadas en el formulario.");
}
