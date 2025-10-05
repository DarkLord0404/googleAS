function permisos(){
  DriveApp.getFolders();
  DocumentApp.openById(PLANTILLA)
  GmailApp.getDrafts()
}

function EnviarCorreo(datos) {
  const PLANTILLA="15Cm69vosJjmiS4iYLmC25fCBN6UrP2iEvTqvzfPiIKY"
  const ID_CARPETA="1yDMDZdm3KvOH25hymtvN0YrTdvg2ZiOH"

  var spreadSheet = SpreadsheetApp.openById("1Hkv8CLofSZPFfTty058hJX0sZGU5r66P_oi7j1tPFMo");
  var hoja = spreadSheet.getSheetByName("Respuestas");
  var filaUsar = hoja.getRange("a1:a").getLastRow();
  var codigo = hoja.getRange(filaUsar,1).getValue();
  
  var respuestas = datos.namedValues
  //var solicita = ["Dirección de correo electrónico"][0]
  
var programacionqx = respuestas["¿Paciente cuenta con programación quirúrgica?"]
                      ? respuestas["¿Paciente cuenta con programación quirúrgica?"][0]
                      : null;

var requiereFormato = respuestas["¿Requiere generar formato de egreso?"]
                      ? respuestas["¿Requiere generar formato de egreso?"][0]
                      : null;

// Verificación de condiciones
if (programacionqx === "No" && requiereFormato === "No") {
    Logger.log("No se requiere formulario ni programación quirúrgica. Finalizando ejecución.");
    return; // Detiene la ejecución si ambas condiciones son "No".
}

if (!programacionqx || programacionqx === "Sí") {
    Logger.log("Paciente con programación quirúrgica. Continúa la ejecución, ignorando requiereFormato.");
    // No se realiza un return aquí, continúa la ejecución.
} else if (!requiereFormato || requiereFormato !== "Sí") {
    Logger.log("No se requiere formulario o el valor no es válido. Paciente pendiente programar.");
    return; // Detiene la ejecución si requiereFormato no es válido y no hay programación quirúrgica.
}

  var nombrePaciente = respuestas["Nombre completo del paciente"] 
                     ? respuestas["Nombre completo del paciente"][0].trim()  
                     : null;
  if (!nombrePaciente) {
    Logger.log("Error: El campo 'Nombre completo del paciente' no está definido o está vacío.");
    return; // Finaliza la función como salida si el formulario es diferente
  }
  var estacion = respuestas["¿En qué estación se encuentra el paciente?"] [0]
  var documento = respuestas["Documento del paciente"] 
                ? respuestas["Documento del paciente"][0].trim() 
                : "";
  var telefonos = respuestas["Teléfonos de contacto"] [0]
  var numCuenta = respuestas["Número de cuenta de la atención"][0]
  var numIngreso = respuestas["Número de ingreso del paciente"][0]
  var fechaIngreso = respuestas["Fecha de ingreso"][0] 
  var ubicacionPaciente = respuestas["Ubicación del paciente"][0]
  var eps = respuestas["EPS"][0]
  var regimensalud = respuestas["Régimen de salud"][0]

  if (!(["Sura", "Sanitas", "Coosalud", "Compensar"].includes(eps))) {
  numCuenta = "Requiere nueva cuenta";
  numIngreso = "Requiere nuevo ingreso";
}

  var especialidad =respuestas["Especialidad que programa el procedimiento"][0]
  var especialista = respuestas["Especialista que programa"][0]
  var procedimiento = respuestas["Procedimiento a realizar"][0]
  var fechaDefinicion = respuestas["Fecha de definición del procedimiento"][0]
  var requiereMaterial = respuestas["Requiere material de osteosíntesis"][0]
  if(requiereMaterial=="Sí"){var material = respuestas["¿Qué material de osteosíntesis se requiere?"][0]} else {var material="No aplica"}
    
  var uciPOP = respuestas["Requiere UCI posquirúrgica?"][0]

  var requiereHD = respuestas["¿Requiere reserva de hemoderivados?"][0]
  if(requiereHD =="Sí"){var reserva = respuestas["¿Qué se requiere reservar?"][0]} else {var reserva="No requiere reserva de hemocomponentes"}

  var fechaInclusion = respuestas["Marca temporal"][0]
  var fechaProgramacionqx = respuestas["Fecha programada del procedimiento"][0]
  var horaProcedimientoqx = respuestas["Hora programada del procedimiento"][0]
  if (programacionqx == "No") {
    var formularioLink = 'https://docs.google.com/forms/d/e/1FAIpQLSc2cMsgk0-42gBJRbUXFtul6i9uVEqyBTEEOjTHrdcFyJqTpg/viewform';
    var fechahoraProcedimientoqxLink = '<a href="' + formularioLink + '" target="_blank">Pendiente programar, programe aquí</a>';
    var fechahoraProcedimientoqx = "Pendiente Programar"
  } else {
  var fechahoraProcedimientoqx = fechaProgramacionqx + " a las " + horaProcedimientoqx;
  var fechahoraProcedimientoqxLink = fechaProgramacionqx + " a las " + horaProcedimientoqx; 
  }
  var fechaAnestesia = respuestas["Fecha del aval de anestesiología"][0]
  var fechaAutorizacion = respuestas["Fecha de la autorización quirúrgica en el ámbito hospitalario"][0]
  var fechaEgreso = respuestas["Fecha de egreso del paciente"][0]
  
  var anestesia = fechaAnestesia || "Pendiente valoración preanestésica";
  var autorizacion = fechaAutorizacion || "Pendiente autorización";
 
  
// Definición del objeto con las estaciones y sus correos
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
let correoEstacion = correosEstaciones[estacion.toUpperCase()] || ''; // Valor por defecto '' si no coincide

// Diccionario que asocia correos con especialidades y el número de programadora como primer elemento
const especialidadesPorCorreo = {
  'programacion.cirugia1@clinicadeoccidente.com': [
    '1', // Número de la programadora
    'Urología oncológica', 'Mastología', 'Cirugía plástica oncológica', 
    'Cirugía plástica', 'Cirugía gastro oncológica', 'Hematología', 'Cirugía de tórax'
  ],
  'programacion.cirugia2@clinicadeoccidente.com': [
    '2',
    'Coloproctología', 'Ginecología oncológica', 'Ortopedia reconstructiva', 
    'Ortopedia oncológica', 'Cirugía de cabeza y cuello'
  ],
  'programacion.cirugia3@clinicadeoccidente.com': [
    '3',
    'Urología', 'Ginecología (No obstetricia)', 'Cirugía vascular', 
    'Neumología', 'Anestesiología del dolor', 'Neurología'
  ],
  'programacion.cirugia4@clinicadeoccidente.com': [
    '4',
    'Otorrinolaringología', 'Cirugía pediátrica', 'Cirugía general', 
    'Maxilofacial', 'Cirugía gastrointestinal: CPRE', 'Cirugía hepatobiliar', 'Cirugía bariátrica', 'Ortopedia general', 'Otras especialidades'
  ],
  'programacion.cirugia5@clinicadeoccidente.com': [
    '5',
    'Ginecología obstétrica', 'Cirugía cardiovascular', 'Gastrointestinal: Ecoendoscopias', 
    'Neurocirugía', 'Cirugía de mano', 'Ortopedia infantil'
  ]
};

// Diccionario que asocia el número con el teléfono
const programadorasInfo = {
  '1': '@+573117433319',
  '2': '@+573205529215',
  '3': '@+573128169706',
  '4': '@+573147427657',
  '5': '@+573127575741'
};

// Determinación directa de las variables
let correoProgramacion = '';
let numeroProgramadora = '';

// Iterar sobre el diccionario para asignar valores a las variables
for (const [correo, especialidades] of Object.entries(especialidadesPorCorreo)) {
  if (especialidades.includes(especialidad)) {
    correoProgramacion = correo;
    numeroProgramadora = programadorasInfo[especialidades[0]]; // El primer elemento es el número
    break; // Detener el ciclo una vez que se encuentre la coincidencia
  }
}

// Resultado en variables independientes
console.log(`Correo Programación: ${correoProgramacion}`);
console.log(`Número Programadora: ${numeroProgramadora}`);



// Asignación de correoProgramacion usando la función

  var correos2 = " "
  if(requiereHD == "Sí"){var correosHD = ", servicio.transfusional@clinicadeoccidente.com,autorizacion.ambulatoria@clinicadeoccidente.com, "}else{var correosHD = correos2}

  var correosHospi = correoEstacion+", myriam.ruiz@clinicadeoccidente.com, alejandra.daza@clinicadeoccidente.com, "
  
  var correosQx= "claudia.nunez@clinicadeoccidente.com, cinthy.castillo@clinicadeoccidente.com, coord.enfermeriacx@clinicadeoccidente.com, niyi.ruiz@clinicadeoccidente.com, material.ots@clinicadeoccidente.com, johana.carreno@clinicadeoccidente.com.co, "
  var correosAdmin= "facturacion.urgencias@clinicadeoccidente.com.co, fernando.trivino@clinicadeoccidente.com, autorizacion.servicios@clinicadeoccidente.com, deivy.canaveral@clinicadeoccidente.com, admisiones@clinicadeoccidente.com, magda.ortiz@clinicadeoccidente.com, mini.meneses@clinicadeoccidente.com.co, yolanda.paz@clinicadeoccidente.com, yineth.canamejoy@clinicadeoccidente.com, olga.vargas@clinicadeoccidente.com, andres.puertas@clinicadeoccidente.com, "
  var correosAuditores= "juan.ballesteros@clinicadeoccidente.com, jimmy.folleco@clinicadeoccidente.com, "
  if ((regimensalud=="Planes Voluntarios")||(regimensalud== "ARL")||(regimensalud=="Especial")){var correosPV = "beatriz.vallecilla@clinicadeoccidente.com, jhonatan.arias@clinicadeoccidente.com, geraldin.delgado@clinicadeoccidente.com, "} else {var correosPV ="alexander.torres@clinicadeoccidente.com, "}
  var correosUrg = "jessica.parra@clinicadeoccidente.com, geovany.rolon@clinicadeoccidente.com, urgencias@clinicadeoccidente.com, medicos.urgencias@clinicadeoccidente.com, "
  if (ubicacionPaciente== "Urgencias"){
  var correos = correosQx+correosAdmin+correosAuditores+correosPV+correosUrg+correoProgramacion+correosHD
  }else if(ubicacionPaciente == "Hospitalización"){var correos = correosQx+correosAdmin+correosAuditores+correosPV+correosHospi+correoProgramacion}
  
Logger.log(correosHD)
Logger.log(correoProgramacion)

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
  txt.replaceText("{{código}}",codigo)
  txt.replaceText("{{eps}}",eps)
  txt.replaceText("{{procedimiento}}",procedimiento)
  txt.replaceText("{{especialidad}}",especialidad)
  txt.replaceText("{{especialista}}",especialista)
  txt.replaceText("{{fechahoraProcedimientoqx}}",fechahoraProcedimientoqx)

  doc.saveAndClose();

  var blob=copiaArchivoPlantilla.getBlob();
  var pdf=carpeta.createFile(blob)
  copiaArchivoPlantilla.setTrashed(true)

  
  GmailApp.sendEmail(
        correos+", alexander.torres@clinicadeoccidente.com",
        "Cirugía diferida "+codigo +" - " +nombrePaciente +" - "+documento, 
        "",
        { attachments:[pdf],
          name: "Cirugía diferida ",
          htmlBody: '<p>Paciente para protocolo de cirug&iacute;a de urgencia diferida:</p><table border="1" style="border-collapse: collapse; width: 100%; height: 108px;"><tbody><tr><td style="width: 50%;"><strong> Datos del paciente </strong></td><td style="width: 50%;"><em> &nbsp; </em></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Nombre del paciente </strong></td><td style="width: 50%; height: 18px;"><em> </em><div><div><em> '+ nombrePaciente+' </em></div></div></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Documento del paciente </strong></td><td style="width: 50%; height: 18px;"><em> </em><div><div><em> '+documento+' </em></div></div></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> EPS </strong></td><td style="width: 50%; height: 18px;"><em> '+eps+' '+regimensalud+' </em></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Teléfonos reportados </strong></td><td style="width: 50%; height: 18px;"><em> '+telefonos+' </em></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> N&uacute;mero de cuenta </strong></td><td style="width: 50%; height: 18px;"><em> </em><div><div><em> '+ numCuenta+' </em></div></div></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> N&uacute;mero de ingreso </strong></td><td style="width: 50%; height: 18px;"><em> </em><div><div><em> '+ numIngreso+' </em></div></div></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Fecha de ingreso del usuario </strong></td><td style="width: 50%; height: 18px;"><em> </em><div><div><em> '+ fechaIngreso+' </em></div></div></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Servicio que activa el protocolo </strong></td><td style="width: 50%; height: 18px;"><em> </em><div><div><em> '+ ubicacionPaciente+' </em></div></div></td></tr></tbody></table><p></p><table border="1" style="border-collapse: collapse; width: 100%;"><tbody><tr><td style="width: 50%;"><strong> Datos del procedimiento </strong></td><td style="width: 50%;"></td></tr><tr><td style="width: 50%;"><strong> Procedimiento a realizar </strong></td><td style="width: 50%;"><em> '+procedimiento+' </em></td></tr><tr><td style="width: 50%;"><strong> Especialista </strong></td><td style="width: 50%;"><em> '+especialista+' </em></td></tr><tr><td style="width: 50%;"><strong> Especialidad </strong></td><td style="width: 50%;"><em> '+especialidad+' </em></td></tr><tr><td style="width: 50%;"><strong> Requiere material de osteos&iacute;ntesis </strong></td><td style="width: 50%;"><em> '+requiereMaterial+' </em></td></tr><tr><td style="width: 50%;"><strong> Material solicitado </strong></td><td style="width: 50%;"><em> '+material+' </em></td></tr><tr><td style="width: 50%;"><strong> Requiere hemoderivados </strong></td><td style="width: 50%;"><em> '+requiereHD+' '+reserva+'<br /></em></td></tr><tr><td style="width: 50%;"><strong> Requiere UCI POP </strong></td><td style="width: 50%;"><em> '+uciPOP+' <br /></em></td></tr><tr><td style="width: 50%;"><strong> Fecha de definici&oacute;n de procedimiento </strong></td><td style="width: 50%;"><em> '+fechaDefinicion+' </em></td></tr></tbody></table><p></p><table border="1" style="border-collapse: collapse; width: 100%; height: 93px;"><tbody><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Avales egreso </strong></td><td style="width: 50%; height: 18px;"></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Fecha de autorizaci&oacute;n </strong></td><td style="width: 50%; height: 18px;"><em> '+autorizacion+' </em></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Fecha de aval de anestesia </strong></td><td style="width: 50%; height: 18px;"><em> '+anestesia+' </em></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Fecha de programaci&oacute;n </strong></td><td style="width: 50%; height: 18px;"><em> '+fechahoraProcedimientoqxLink+' </em></td></tr><tr style="height: 18px;"><td style="width: 50%; height: 18px;"><strong> Fecha de egreso del paciente </strong></td><td style="width: 50%; height: 18px;"><em> '+fechaEgreso+' </em></td></tr></tbody></table><p>Agradezco responder a este correo con la confirmaci&oacute;n de recibido, y los avances correspondientes en lo pendiente. </p><p>Respecto al material de osteosíntesis el equipo de urgencias coloca en este correo con lo que se cuente en la historia clínica sin hacer verificaciones adicionales. </p><p>En el caso en el que el paciente requiera de reserva de hemoderivados, si su procedimiento está programado más allá de 72 horas desde el egreso, está reserva deberá ser gestionadas en el ámbito ambulatorio, el servicio a cargo del paciente entregará orden ambulatoria, la central de autorizaciones ambulatoria gestionará la autorización de la misma y el paciente deberá asistir 72 horas antes del procedimiento al laboratorio clínico para toma de la muestra a la sede Sheraton. En el caso de que el procedimiento sea programado antes de 72 horas de su egreso, la muestra debe tomarse de manera intrahospitalaria y se debe esperar el resultado para saber si requiere nueva muestra.</p><p>El estado actual de este caso y de los casos activos es posible consultarlo en el siguiente link</p><p><a href="https://docs.google.com/spreadsheets/d/1Hkv8CLofSZPFfTty058hJX0sZGU5r66P_oi7j1tPFMo/edit?resourcekey=&amp;gid=934238360#gid=934238360"> Base de datos cirug&iacute;a diferida </a></p>'
        })  

actualizarFormulario();
CorreoCreacion(datos);

}