function permisos(){
  DriveApp.getFolders();
  GmailApp.getDrafts()
}

function CorreoCreacion(datos) {
   var spreadSheet = SpreadsheetApp.openById("1Hkv8CLofSZPFfTty058hJX0sZGU5r66P_oi7j1tPFMo");
  var hoja = spreadSheet.getSheetByName("Respuestas");
  var filaUsar = hoja.getRange("a1:a").getLastRow();
  var codigo = hoja.getRange(filaUsar,1).getValue();
  
  var respuestas = datos.namedValues
  var nombrePaciente = respuestas["Nombre completo del paciente"] 
                     ? respuestas["Nombre completo del paciente"][0] 
                     : null;

  if (!nombrePaciente) {
    Logger.log("Error: El campo 'Nombre completo del paciente' no está definido o está vacío.");
    return; // Finaliza la función como salida si el formulario es diferente
  }

  var requiereFormato = respuestas["¿Requiere generar formato de egreso?"]
                      ?respuestas["¿Requiere generar formato de egreso?"][0]
                      : null;
  /*if (!requiereFormato || requiereFormato == "Sí") {
    Logger.log("No se requiere formulario o el valor no es válido. Paciente pendiente programar.");
    return; // Finaliza la ejecución de este bloque de código.
  }*/
  
  var programacionqx = respuestas["¿Paciente cuenta con programación quirúrgica?"]
                      ?respuestas["¿Paciente cuenta con programación quirúrgica?"][0]
                      : null;
  if (!programacionqx || programacionqx == "Sí") {
    Logger.log("No se requiere preingreso, paciente con programación.");
    return; // Finaliza la ejecución de este bloque de código.
  }


  var estacion = respuestas["¿En qué estación se encuentra el paciente?"] [0]
  var documento = respuestas["Documento del paciente"] 
                ? respuestas["Documento del paciente"][0].trim() 
                : "";
  var telefonos = respuestas["Teléfonos de contacto"] [0]
  var ubicacionPaciente = respuestas["Ubicación del paciente"][0]
  var eps = respuestas["EPS"][0]
  var regimensalud = respuestas["Régimen de salud"][0]
  
  var especialidad =respuestas["Especialidad que programa el procedimiento"][0]
  var especialista = respuestas["Especialista que programa"][0]
  var procedimiento = respuestas["Procedimiento a realizar"][0]
  var fechaDefinicion = respuestas["Fecha de definición del procedimiento"][0]
  var requiereMaterial = respuestas["Requiere material de osteosíntesis"][0]
  if(requiereMaterial=="Sí"){var material = respuestas["¿Qué material de osteosíntesis se requiere?"][0]} else {var material="No aplica"}
    
  var uciPOP = respuestas["Requiere UCI posquirúrgica?"][0]

  var requiereHD = respuestas["¿Requiere reserva de hemoderivados?"][0]
  if(requiereHD =="Sí"){var reserva = respuestas["¿Qué se requiere reservar?"][0]} else {var reserva="No requiere reserva de hemocomponentes"}

  var fechaAnestesia = respuestas["Fecha del aval de anestesiología"][0]
  var fechaAutorizacion = respuestas["Fecha de la autorización quirúrgica en el ámbito hospitalario"][0]
  
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
  '1': '@+57 311 7433319',
  '2': '@+57 320 5529215',
  '3': '@+57 312 8169706',
  '4': '@+57 314 7427657',
  '5': '@+57 312 7575741'
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
  var correosHospi2 = "alexander.torres@clinicadeoccidente.com"
  var correosHospi = correoEstacion+", myriam.ruiz@clinicadeoccidente.com, alejandra.daza@clinicadeoccidente.com"
  var correosUrg = "jessica.parra@clinicadeoccidente.com, geovany.rolon@clinicadeoccidente.com, urgencias@clinicadeoccidente.com"
  var formularioLink = 'https://docs.google.com/forms/d/e/1FAIpQLSc2cMsgk0-42gBJRbUXFtul6i9uVEqyBTEEOjTHrdcFyJqTpg/viewform';
  var fechahoraProcedimientoqxLink = '<a href="' + formularioLink + '" target="_blank">Programe aquí</a>';
  if (ubicacionPaciente== "Urgencias"){
  var correos = correosUrg//+correoProgramacion
  }else if(ubicacionPaciente == "Hospitalización"){var correos = correosHospi}
  
  
  

  GmailApp.sendEmail(
        correos +", alexander.torres@clinicadeoccidente.com",
        "Preingreso cirugía diferida "+codigo +" - " +nombrePaciente +" - "+documento, 
        "",
        { name: "Preingreso a cirugía diferida ",
          htmlBody: '<p><strong>*CASO '+codigo+'*</strong></p><p>Buenos d&iacute;as, a continuaci&oacute;n se presenta paciente para el protocolo de cirug&iacute;a diferida, solicitamos su apoyo para la asignaci&oacute;n de FECHA y HORA de procedimiento quirurgico</p><p>*Datos del paciente:*<br />Nombre del paciente:&nbsp;'+ nombrePaciente+'<br />Documento del paciente:&nbsp;'+documento+'<br />Telefono de contacto:&nbsp;'+ telefonos+'<br />EPS:&nbsp;'+eps+' '+regimensalud+'<br />Servicio que activa el protocolo: '+ ubicacionPaciente+'</p><p>*Datos del procedimiento*<br />Procedimiento a realizar: <em>'+procedimiento+'</em><br />Especialista: '+especialista+'<br />Especialidad:'+especialidad+'<br />Requiere material de osteos&iacute;ntesis: '+requiereMaterial+'<br />Material solicitado:&nbsp; '+material+'<br />Requiere hemoderivados:&nbsp;'+requiereHD+'<br />Requiere UCI POP:&nbsp; '+uciPOP+'<br />Fecha de definici&oacute;n de procedimiento: '+fechaDefinicion+'</p><p>*Avales egreso*<br />Fecha de autorizaci&oacute;n: '+autorizacion+'<br />Fecha de aval de anestesia: '+anestesia+'</p><p>Agradecemos a '+numeroProgramadora+' responder a este mensaje con la fecha y hora de programaci&oacute;n</p><p>_NOTA: Respecto al material de osteos&iacute;ntesis el equipo de urgencias coloca en este correo con lo que se cuente en la historia cl&iacute;nica sin hacer verificaciones adicionales._</p> <p><table border="1" style="border-collapse: collapse; width: 100%;"<tr><td style="width: 50%; height: 18px;"><em> '+fechahoraProcedimientoqxLink+' </em></td></tr></table></p>'
        })  

actualizarFormulario();
}
