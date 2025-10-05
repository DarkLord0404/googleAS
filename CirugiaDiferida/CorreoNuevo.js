function permisos(){
  DriveApp.getFolders();
  DocumentApp.openById(PLANTILLA)
  GmailApp.getDrafts()
}

function CorreoNuevo(datos) {
  const PLANTILLA="15Cm69vosJjmiS4iYLmC25fCBN6UrP2iEvTqvzfPiIKY"
  const ID_CARPETA="1yDMDZdm3KvOH25hymtvN0YrTdvg2ZiOH"
  
const spreadsheetId = "1Hkv8CLofSZPFfTty058hJX0sZGU5r66P_oi7j1tPFMo";
const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
const hoja = spreadsheet.getSheetByName("Respuestas");
const hoja2 = spreadsheet.getSheetByName("activosActualizar2");

// Función para formatear fechas en dd/MM/yyyy
function formatearFecha(fecha) {
  if (!fecha || Object.prototype.toString.call(fecha) !== "[object Date]") {
    return ""; // Devuelve vacío si no es una fecha válida
  }
  const dia = ("0" + fecha.getDate()).slice(-2); // Día con 2 dígitos
  const mes = ("0" + (fecha.getMonth() + 1)).slice(-2); // Mes con 2 dígitos
  const anio = fecha.getFullYear();
  return `${dia}/${mes}/${anio}`;
}

// Obtener valores de la columna C de la hoja 2
const lastRowHoja2 = hoja2.getLastRow(); // Última fila con datos en hoja2
if (lastRowHoja2 < 2) {
  Logger.log("No hay datos para procesar en hoja 2.");
  return; // Salir si no hay registros en la hoja 2
}

// Obtener el rango activo de b2:c de la hoja 2
const dataHoja2 = hoja2.getRange(2, 2, lastRowHoja2 - 1, 3).getValues(); // Desde B2 a C
const columnaC = dataHoja2.map(row => row[0]); // Obtener columna B como array

// Obtener el valor del evento de respuestas
var respuestas = datos.namedValues;

var accion = respuestas["La acción que desea realizar es:"][0]
  if (!accion || accion !== "Asignar la fecha y hora de programación a un paciente") {
    Logger.log("No se requiere nuevo formato.");
    Logger.log(accion);
    return; // Finaliza la ejecución de este bloque de código.
  }

var cadena1 = respuestas["Caso a programar"] 
              ? respuestas["Caso a programar"][0] 
                ? respuestas["Caso a programar"][0] 
                : null 
              : null;

if (!cadena1) {
  Logger.log("Error: El campo 'Caso' no está definido, es vacío o no tiene elementos.");
  return; // Finaliza la función o detiene el script
}

Logger.log(cadena1)
Logger.log(columnaC)

// Buscar coincidencia entre cadena1 y la columna C
let filaEncontradaHoja2 = -1; // Inicialmente no encontrada
for (let i = 0; i < columnaC.length; i++) {
  if (columnaC[i] === cadena1) {
    filaEncontradaHoja2 = i + 2; // Ajustar índice para corresponder con la fila real
    break;
  }
}

if (filaEncontradaHoja2 === -1) {
  Logger.log("No se encontró coincidencia en la columna C de la hoja 2.");
  return; // Salir si no hay coincidencias
}

// Obtener el dato de la columna A de la fila encontrada en la hoja 2
const codigoHoja2 = hoja2.getRange(filaEncontradaHoja2, 1).getValue();

// Validar fila en la hoja 1 (Respuestas) según el código encontrado
const dataHoja1 = hoja.getRange(1, 1, hoja.getLastRow(), 1).getValues().flat(); // Columna A como array
let filaUsar = dataHoja1.indexOf(codigoHoja2) + 1; // Buscar fila en la hoja 1

if (filaUsar === 0) {
  Logger.log("No se encontró el código en la hoja 1 (Respuestas).");
  return; // Salir si no se encuentra el código en la hoja 1
}

Logger.log(`Fila encontrada en hoja 1: ${filaUsar}`);
Logger.log(`Código encontrado: ${codigoHoja2}`);

// Fin del problema_________



  //var solicita = ["Dirección de correo electrónico"][0]

  var nombrePaciente = hoja.getRange(filaUsar,4).getValue();
  
  var documento = hoja.getRange(filaUsar, 5).getValue().toString().trim();
  var telefonos = hoja.getRange(filaUsar, 24).getValue().toString().trim();
  var numCuenta = hoja.getRange(filaUsar, 6).getValue().toString().trim();
  var numIngreso = hoja.getRange(filaUsar, 25).getValue().toString().trim();
  var fechaIngreso = formatearFecha(hoja.getRange(filaUsar,7).getValue());
  var ubicacionPaciente = hoja.getRange(filaUsar,8).getValue();
  var eps = hoja.getRange(filaUsar,9).getValue();
  var regimensalud = hoja.getRange(filaUsar,10).getValue();

  if (!(["Sura", "Sanitas", "Coosalud", "Compensar"].includes(eps))) {
  numCuenta = "Requiere nueva cuenta";
  numIngreso = "Requiere nuevo ingreso";
}

  var especialidad = hoja.getRange(filaUsar,11).getValue();
  var especialista = hoja.getRange(filaUsar,12).getValue();
  var procedimiento = hoja.getRange(filaUsar,13).getValue();
  var fechaDefinicion = formatearFecha(hoja.getRange(filaUsar,14).getValue());
  var requiereMaterial = hoja.getRange(filaUsar,15).getValue();
  if(requiereMaterial=="Sí"){var material = hoja.getRange(filaUsar,16).getValue();} else {var material="No aplica"}
    
  var uciPOP = hoja.getRange(filaUsar,17).getValue();

  var requiereHD = hoja.getRange(filaUsar,18).getValue();
  if(requiereHD =="Sí"){var reserva = hoja.getRange(filaUsar,28).getValue();} else {var reserva="No requiere reserva de hemocomponentes"}

  //var fechaInclusion = respuestas["Marca temporal"][0]
  var programacionqx = "Sí"
  var fechaProgramacionqx = respuestas["Fecha programada del procedimiento"][0]
  var horaProcedimientoqx = respuestas["Hora programada del procedimiento"][0]
  if (programacionqx == "No") {
    var formularioLink = 'https://docs.google.com/forms/d/e/1FAIpQLSc2cMsgk0-42gBJRbUXFtul6i9uVEqyBTEEOjTHrdcFyJqTpg/viewform';
    var fechahoraProcedimientoqx = '<a href="' + formularioLink + '" target="_blank">Pendiente programar, programe aquí</a>';
  } else {
  var fechahoraProcedimientoqx = fechaProgramacionqx + " a las " + horaProcedimientoqx; 
  }
  var fechaAnestesia = formatearFecha(hoja.getRange(filaUsar,22).getValue());
  var fechaAutorizacion = formatearFecha(hoja.getRange(filaUsar,21).getValue());
  var fechaEgreso = formatearFecha(hoja.getRange(filaUsar,23).getValue());
  
  var anestesia = fechaAnestesia || "Pendiente valoración preanestésica";
  var autorizacion = fechaAutorizacion || "Pendiente autorización";
 
  
// Definición del objeto con las estaciones y sus correos
const estacionRaw = hoja.getRange(filaUsar, 26).getValue(); // Valor obtenido directamente de la hoja
const estacion = (typeof estacionRaw === "string" ? estacionRaw.trim() : estacionRaw.toString().trim()) || ""; // Convertir a string y eliminar espacios si no es nulo

if (!estacion) {
  Logger.log(`El valor de 'estacion' es vacío o inválido. Se encontró: ${estacionRaw}`);
}

// Objeto de correos por estación
const correosEstaciones = {
  '10B': 'torreb10.estenf@clinicadeoccidente.com',
  '9B': 'torreb9.estenf@clinicadeoccidente.com',
  '8B': 'torreb8.estenf@clinicadeoccidente.com',
  '7B': 'torreb7.estenf@clinicadeoccidente.com',
  '5B': 'torreb5.estenfe@clinicadeoccidente.com',
  'UNIDAD AISLAMIENTO': 'unidad.aislamiento@clinicadeoccidente.com',
  '5A': 'estacion5a@clinicadeoccidente.com',
  '5 TORRE C': 'estacion5.tc@clinicadeoccidente.com',
  '4A': 'estacion4a@clinicadeoccidente.com',
  '4 TORRE C': 'estacion4.tc@clinicadeoccidente.com',
  '3F': 'estacion3f@clinicadeoccidente.com',
  '3E': 'estacion3e@clinicadeoccidente.com',
  '3E TORRE C': 'estacion3e.tc@clinicadeoccidente.com',
  '3D2': 'estacion3d2@clinicadeoccidente.com',
  '3C': 'estacion3c@clinicadeoccidente.com',
  '3B': 'estacion3b@clinicadeoccidente.com',
  '3A': 'estacion3a@clinicadeoccidente.com',
  '2B': 'estacion2b@clinicadeoccidente.com'
};

// Asignación de correoEstacion usando el objeto
const correoEstacion = correosEstaciones[estacion.toUpperCase()] || ""; // Valor predeterminado si no coincide
Logger.log(`Correo asignado para estación '${estacion}': ${correoEstacion}`);


// Definición de la función para obtener el correo de programación según la especialidad
function obtenerCorreoProgramacion(especialidad) {
  const especialidadesPorCorreo = {
    'programacion.cirugia1@clinicadeoccidente.com': [
      'Urología oncológica', 'Mastología', 'Cirugía plástica oncológica', 
      'Cirugía plástica', 'Cirugía gastro oncológica', 'Hematología', 'Cirugía de tórax'
    ],
    'programacion.cirugia2@clinicadeoccidente.com': [
      'Coloproctología', 'Ginecología oncológica', 'Ortopedia reconstructiva', 
      'Ortopedia oncológica', 'Cirugía de cabeza y cuello'
    ],
    'programacion.cirugia3@clinicadeoccidente.com': [
      'Urología', 'Ginecología (No obstetricia)', 'Cirugía vascular', 
      'Neumología', 'Anestesiología del dolor', 'Neurología'
    ],
    'programacion.cirugia4@clinicadeoccidente.com': [
      'Otorrinolaringología', 'Cirugía pediátrica', 'Cirugía general', 
      'Maxilofacial', 'Cirugía gastrointestinal: CPRE', 'Cirugía hepatobiliar', 'Cirugía bariátrica', 'Otras especialidades'
    ],
    'programacion.cirugia5@clinicadeoccidente.com': [
      'Ginecología obstétrica', 'Cirugía cardiovascular', 'Gastrointestinal: Ecoendoscopias', 
      'Neurocirugía', 'Ortopedia general', 'Cirugía de mano', 'Ortopedia infantil'
    ]
  };

  // Itera sobre cada clave (correo) y verifica si la especialidad está en el grupo correspondiente
  for (const [correo, especialidades] of Object.entries(especialidadesPorCorreo)) {
    if (especialidades.includes(especialidad)) {
      return correo;
    }
  }
  return ''; // Retorna vacío si la especialidad no coincide con ningún grupo
}

// Asignación de correoProgramacion usando la función
let correoProgramacion = obtenerCorreoProgramacion(especialidad);
  var correos2 = " "
  if(requiereHD =="Sí"){var correosHD = "yehiner.calderon@clinicadeoccidente.com, servicio.transfusional@clinicadeoccidente.com, "}else{var correosHD = correos2}

  var correosHospi = correoEstacion+", myriam.ruiz@clinicadeoccidente.com, alejandra.daza@clinicadeoccidente.com, "
  
  var correosQx= "claudia.nunez@clinicadeoccidente.com, cinthy.castillo@clinicadeoccidente.com, coord.enfermeriacx@clinicadeoccidente.com, niyi.ruiz@clinicadeoccidente.com, material.ots@clinicadeoccidente.com, johana.carreno@clinicadeoccidente.com.co, "
  var correosAdmin= "facturacion.urgencias@clinicadeoccidente.com.co, fernando.trivino@clinicadeoccidente.com, autorizacion.servicios@clinicadeoccidente.com, deivy.canaveral@clinicadeoccidente.com, admisiones@clinicadeoccidente.com, magda.ortiz@clinicadeoccidente.com, mini.meneses@clinicadeoccidente.com.co, yolanda.paz@clinicadeoccidente.com, yineth.canamejoy@clinicadeoccidente.com, olga.vargas@clinicadeoccidente.com, andres.puertas@clinicadeoccidente.com, "
  var correosAuditores= "juan.ballesteros@clinicadeoccidente.com, jimmy.folleco@clinicadeoccidente.com, "
  if ((regimensalud=="Planes Voluntarios")||(regimensalud== "ARL")||(regimensalud=="Especial")){var correosPV = "beatriz.vallecilla@clinicadeoccidente.com, jhonatan.arias@clinicadeoccidente.com, geraldin.delgado@clinicadeoccidente.com, "} else {var correosPV ="alexander.torres@clinicadeoccidente.com, "}
  var correosUrg = "jessica.parra@clinicadeoccidente.com, geovany.rolon@clinicadeoccidente.com, urgencias@clinicadeoccidente.com, medicos.urgencias@clinicadeoccidente.com, "
  if (ubicacionPaciente== "Urgencias"){
  var correos = correosQx+correosAdmin+correosAuditores+correosPV+correosUrg+correoProgramacion+correosHD
  }else if(ubicacionPaciente == "Hospitalización"){var correos = correosQx+correosAdmin+correosAuditores+correosPV+correosHospi+correoProgramacion}
  
  

/* VARIABLES PDF*/
  var carpeta=DriveApp.getFolderById(ID_CARPETA);
  var archivoPlantilla= DriveApp.getFileById(PLANTILLA);
  var copiaArchivoPlantilla=archivoPlantilla.makeCopy(documento);
  var idArchivoCopia=copiaArchivoPlantilla.getId();

  var doc=DocumentApp.openById(idArchivoCopia)
  var txt=doc.getBody();
  txt.replaceText("{{nombrePaciente}}",nombrePaciente)
  txt.replaceText("{{documento}}",documento)
  txt.replaceText("{{numCuenta}}",numCuenta)
  txt.replaceText("{{numIngreso}}",numIngreso)
  txt.replaceText("{{código}}",codigoHoja2)
  txt.replaceText("{{eps}}",eps)
  txt.replaceText("{{procedimiento}}",procedimiento)
  txt.replaceText("{{especialidad}}",especialidad)
  txt.replaceText("{{especialista}}",especialista)
  txt.replaceText("{{fechahoraProcedimientoqx}}",fechahoraProcedimientoqx)

  doc.saveAndClose();

  var blob=copiaArchivoPlantilla.getBlob();
  var pdf=carpeta.createFile(blob)
  copiaArchivoPlantilla.setTrashed(true)

  Logger.log(correos)

  
  GmailApp.sendEmail(
        correos+", alexander.torres@clinicadeoccidente.com",
        "Cirugía diferida "+codigoHoja2 +" - Paciente: " +nombrePaciente +" - "+documento, 
        "",
        { attachments:[pdf],
          name: "Cirugía diferida ",
          htmlBody: '<p>Paciente para protocolo de cirug&iacute;a de urgencia diferida:</p><table border="1" style="border-collapse: collapse; width: 100%; height: 108px;"><tbody><tr><td style="width: 50%;"><strong> Datos del paciente </strong></td><td style="width: 50%;"><em> &nbsp; </em></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Nombre del paciente </strong></td><td style="width: 50%; height: 18px;"><em> </em><div><div><em> '+ nombrePaciente+' </em></div></div></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Documento del paciente </strong></td><td style="width: 50%; height: 18px;"><em> </em><div><div><em> '+documento+' </em></div></div></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> EPS </strong></td><td style="width: 50%; height: 18px;"><em> '+eps+' '+regimensalud+' </em></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Teléfonos reportados </strong></td><td style="width: 50%; height: 18px;"><em> '+telefonos+' </em></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> N&uacute;mero de cuenta </strong></td><td style="width: 50%; height: 18px;"><em> </em><div><div><em> '+ numCuenta+' </em></div></div></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> N&uacute;mero de ingreso </strong></td><td style="width: 50%; height: 18px;"><em> </em><div><div><em> '+ numIngreso+' </em></div></div></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Fecha de ingreso del usuario </strong></td><td style="width: 50%; height: 18px;"><em> </em><div><div><em> '+ fechaIngreso+' </em></div></div></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Servicio que activa el protocolo </strong></td><td style="width: 50%; height: 18px;"><em> </em><div><div><em> '+ ubicacionPaciente+' </em></div></div></td></tr></tbody></table><p></p><table border="1" style="border-collapse: collapse; width: 100%;"><tbody><tr><td style="width: 50%;"><strong> Datos del procedimiento </strong></td><td style="width: 50%;"></td></tr><tr><td style="width: 50%;"><strong> Procedimiento a realizar </strong></td><td style="width: 50%;"><em> '+procedimiento+' </em></td></tr><tr><td style="width: 50%;"><strong> Especialista </strong></td><td style="width: 50%;"><em> '+especialista+' </em></td></tr><tr><td style="width: 50%;"><strong> Especialidad </strong></td><td style="width: 50%;"><em> '+especialidad+' </em></td></tr><tr><td style="width: 50%;"><strong> Requiere material de osteos&iacute;ntesis </strong></td><td style="width: 50%;"><em> '+requiereMaterial+' </em></td></tr><tr><td style="width: 50%;"><strong> Material solicitado </strong></td><td style="width: 50%;"><em> '+material+' </em></td></tr><tr><td style="width: 50%;"><strong> Requiere hemoderivados </strong></td><td style="width: 50%;"><em> '+requiereHD+' <br /></em></td></tr><tr><td style="width: 50%;"><strong> Requiere UCI POP </strong></td><td style="width: 50%;"><em> '+uciPOP+' <br /></em></td></tr><tr><td style="width: 50%;"><strong> Qué hemoderivados se requiere? </strong></td><td style="width: 50%;"><em> '+reserva+' <br /></em></td></tr><tr><td style="width: 50%;"><strong> Fecha de definici&oacute;n de procedimiento </strong></td><td style="width: 50%;"><em> '+fechaDefinicion+' </em></td></tr></tbody></table><p></p><table border="1" style="border-collapse: collapse; width: 100%; height: 93px;"><tbody><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Avales egreso </strong></td><td style="width: 50%; height: 18px;"></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Fecha de autorizaci&oacute;n </strong></td><td style="width: 50%; height: 18px;"><em> '+autorizacion+' </em></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Fecha de aval de anestesia </strong></td><td style="width: 50%; height: 18px;"><em> '+anestesia+' </em></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Fecha de programaci&oacute;n </strong></td><td style="width: 50%; height: 18px;"><em> '+fechahoraProcedimientoqx+' </em></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Fecha de egreso del paciente </strong></td><td style="width: 50%; height: 18px;"><em> '+fechaEgreso+' </em></td></tr></tbody></table><p>Agradezco responder a este correo con la confirmaci&oacute;n de recibido, y los avances correspondientes en lo pendiente. </p><p>Respecto al material de osteosíntesis el equipo de urgencias coloca en este correo con lo que se cuente en la historia clínica sin hacer verificaciones adicionales. </p><p>El estado actual de este caso y de los casos activos es posible consultarlo en el siguiente link</p><p><a href="https://docs.google.com/spreadsheets/d/1Hkv8CLofSZPFfTty058hJX0sZGU5r66P_oi7j1tPFMo/edit?resourcekey=&amp;gid=934238360#gid=934238360"> Base de datos cirug&iacute;a diferida </a></p>'
        })  
        actualizarFormulario();
        CorreoCreacion(datos);
}

