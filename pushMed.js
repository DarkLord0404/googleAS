/** CONFIGURACIÓN **/

const FORMULARIOS = {
  accesosVasculares: "1Qj0LP6Hm9GJHIhpKGu2nd1fGn8xiSEk-CVqkSiMJPXw",
  mortalidad: "1eqw4OLAktam9n1ZoW5FQ0w5pnqp_KqZ0Z12iLH7vWU4",
  cambioTurno: "1capWe9mMiYINp3e-x09dPDbhrr824H867ZSL5OShViM",
  diciembre: "1zt-cNAUkXI6lG1IiBj3dgFpYjiy84p5xxJN_N-SYhuM",
  extra: "1f5Fw9r7wx1PtYgcUTmZQpAds0seUOdXE7-B1ZhSU8XI",
  auto: "1OhXdLTtIO7MU7kV7UB3z7Ks0YV_s3NyEA0gpbB1oHlE",
  protocolos: "1O0xVZ7MLNiRqjNPyWjYYGZMCFvESlpZkvz9QlNPqH8k",
  turnos: "1kk4hVgNmkAaAAAxzABsnKQL_xJIBk5qppNJVQHcT4OE",
  quejas: "1EjQIomLocuKNLtH-cZDGjwQW07vkTALNEz5AoLSuq1k",
  felicitaciones: "1jQXaLoxiHvCXon5XE8RgK9fwhog2Dc5PmZ9PgCwadPU",
  solicitudesMes: "1yBzK3i4xUvZxlPUQG-TItii9nal8QxqyZCDADucllL0",
  hiperglicemias2: "1YWbClZgr03vySqkj26U_eWrHi81FiGajCPeWYVUDzU4",
  bronquiolitis: "1ju_BKG0t-wCQ4xqKZzJKgil_0BjDGHfo2Og8loepzeE",
  cirugiaDiferida: "1STyH2dW-chhl2LmBIlS-cuouv7d9ry3FHWdfHBHPdBI",
  asistencia: "1fs26_kMVuHv274PCj89-C5qBi8Im5RUWnIBFkD1IwXk",
};

const PREGUNTAS = {
  solicita: 797739091,
  reemplaza: 2026226272,
  cirugiaDiferida: 1627793681,
  asistencia: 303307952,
  otros: 79035639
};

const GRUPOS = {
  medicosUrgencias: { hoja: "medicos_urgencias", col: 2 },
  profesionalesActivos: { hoja: "profesionales_activos", col: 1 },
  habilitadosCDT: { hoja: "Habilitados_CDT", col: 1 },
  especialistasActivos: { hoja: "especialistas_activos", col: 1 }
};

/** ACTUALIZACIONES **/
const ACTUALIZACIONES = [
  { nombre: "Accesos Vasculares", formId: FORMULARIOS.accesosVasculares, preguntas: [PREGUNTAS.solicita], grupo: GRUPOS.medicosUrgencias },
  { nombre: "Mortalidad", formId: FORMULARIOS.mortalidad, preguntas: [PREGUNTAS.solicita], grupo: GRUPOS.profesionalesActivos },
  { nombre: "Cambio Turno", formId: FORMULARIOS.cambioTurno, preguntas: [PREGUNTAS.solicita, PREGUNTAS.reemplaza], grupo: GRUPOS.habilitadosCDT },
  { nombre: "Diciembre", formId: FORMULARIOS.diciembre, preguntas: [PREGUNTAS.solicita], grupo: GRUPOS.profesionalesActivos },
  { nombre: "Extra", formId: FORMULARIOS.extra, preguntas: [PREGUNTAS.solicita], grupo: GRUPOS.profesionalesActivos },
  { nombre: "Auto", formId: FORMULARIOS.auto, preguntas: [PREGUNTAS.solicita], grupo: GRUPOS.profesionalesActivos },
  { nombre: "Protocolos", formId: FORMULARIOS.protocolos, preguntas: [PREGUNTAS.solicita], grupo: GRUPOS.profesionalesActivos },
  { nombre: "Turnos", formId: FORMULARIOS.turnos, preguntas: [PREGUNTAS.solicita], grupo: GRUPOS.profesionalesActivos },
  { nombre: "Quejas", formId: FORMULARIOS.quejas, preguntas: [PREGUNTAS.otros], grupo: GRUPOS.medicosUrgencias },
  { nombre: "Felicitaciones", formId: FORMULARIOS.felicitaciones, preguntas: [PREGUNTAS.otros], grupo: GRUPOS.medicosUrgencias },
  { nombre: "Solicitudes Mes", formId: FORMULARIOS.solicitudesMes, preguntas: [PREGUNTAS.solicita], grupo: GRUPOS.profesionalesActivos },
  { nombre: "Hiperglicemias 2", formId: FORMULARIOS.hiperglicemias2, preguntas: [PREGUNTAS.solicita], grupo: GRUPOS.profesionalesActivos },
  { nombre: "Bronquiolitis", formId: FORMULARIOS.bronquiolitis, preguntas: [PREGUNTAS.solicita], grupo: GRUPOS.medicosUrgencias },
  { nombre: "Cirugía Diferida", formId: FORMULARIOS.cirugiaDiferida, preguntas: [PREGUNTAS.cirugiaDiferida], grupo: GRUPOS.medicosUrgencias },
  { nombre: "Asistencia", formId: FORMULARIOS.asistencia, preguntas: [PREGUNTAS.asistencia], grupo: GRUPOS.especialistasActivos }
];

/** FUNCIÓN GENERAL **/

function pushActualizaciones() {
  const ss = SpreadsheetApp.openById("1tm2_0IlJJjLu5EElbbv9hFSjkS41d4g_g8bLUl943pc");
  Logger.clear();
  Logger.log("🚀 Iniciando actualización de formularios…");

  ACTUALIZACIONES.forEach(config => {
    Logger.log(`📄 Procesando formulario "${config.nombre}" (ID: ${config.formId})`);
    const hoja = ss.getSheetByName(config.grupo.hoja);

    if (!hoja) {
      Logger.log(`⚠️ No se encontró la hoja ${config.grupo.hoja}`);
      return;
    }

    const datos = hoja.getRange(3, config.grupo.col, hoja.getLastRow() - 2, 1).getValues();
    const opciones = datos.map(row => row[0]).filter(item => item);
    Logger.log(`Opciones detectadas: ${opciones.length}`);

    const formulario = FormApp.openById(config.formId);

    config.preguntas.forEach(id => {
      Logger.log(`🔎 Buscando pregunta con ID ${id}…`);
      const pregunta = formulario.getItemById(id);

      if (!pregunta) {
        Logger.log(`⚠️ No se encontró la pregunta con ID ${id} en el formulario ${config.formId}`);
        return; // pasa a la siguiente
      }

      try {
        pregunta.asListItem().setChoiceValues(opciones);
        Logger.log(`✅ Pregunta ${id} actualizada con ${opciones.length} opciones.`);
      } catch (e) {
        Logger.log(`❌ Error al actualizar la pregunta ${id}: ${e}`);
      }
    });
  });

  Logger.log("🎉 Actualización finalizada.");
}