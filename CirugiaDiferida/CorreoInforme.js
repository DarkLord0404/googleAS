function correoInforme() {
  var spreadSheet = SpreadsheetApp.openById("1Hkv8CLofSZPFfTty058hJX0sZGU5r66P_oi7j1tPFMo");
  var hoja = spreadSheet.getSheetByName("Correo");
  if (!hoja) return;

  var hoy = new Date();
  hoy.setDate(hoy.getDate() - 3);
  var año = hoy.getFullYear();
  var mes = hoy.getMonth() + 1;
  var nombreMes = hoy.toLocaleString('es-ES', { month: 'long' });
  var mesFormateado = mes < 10 ? '0' + mes : mes.toString();
  var valorBuscado = año + mesFormateado + " " + nombreMes.charAt(0).toUpperCase() + nombreMes.slice(1);

  var resultadoBusqueda = ejecutarBusquedaPrimeraColumna(hoja, "A3:F24", valorBuscado);
  var tiemposPromedio = resultadoBusqueda.tiemposPromedio;
  var cantUrgencias = resultadoBusqueda.cantUrgencias;
  var cantHospi = resultadoBusqueda.cantHospi;
  var resultadosFuncion2 = ejecutarBusquedaPrimeraFila(hoja, "A27:y55", valorBuscado) || [];
  var cantidadCasos = resultadosFuncion2.reduce((sum, item) => sum + item[1], 0);
  var resultadosFuncion3 = ejecutarBusquedaPrimeraFila(hoja, "A58:y85", valorBuscado) || [];
  var estanciaInactiva = ejecutarBusquedaEstanciaInactiva(hoja, "A87:B120", valorBuscado);
  var casosActivos = hoja.getRange("k4").getValue();

  resultadosFuncion2.sort((a, b) => b[1] - a[1]);
  resultadosFuncion3.sort((a, b) => b[1] - a[1]);

  // Gráficas de torta (EAPB y Especialidad) construidas en hoja temporal
  var graficoEapb = crearYExportarGrafico(hoja, resultadosFuncion2, "Casos por EAPB", { row: 102, col: 1 });
  var graficoEspecialidad = crearYExportarGrafico(hoja, resultadosFuncion3, "Casos por Especialidad", { row: 120, col: 1 });

  // Datos y gráfica de barras de casos completados por mes (año en curso) desde "Base de datos"
  var datosMensuales = obtenerCasosCompletadosPorMes(spreadSheet, año); // [["Mes","Cantidad"], ...]
  var graficoMensual = crearYExportarGraficoBarrasTemporal(spreadSheet, datosMensuales, "Casos completados por mes " + año);
  var datosCumple = obtenerPorcentajeCumplimientoPorMes(spreadSheet, año);
  var graficoCumple = crearYExportarGraficoBarrasCumplimiento(spreadSheet, datosCumple, "Cumplimiento de programación " + año + " (%)");

  enviarInformeCorreo(
    cantidadCasos,
    tiemposPromedio,
    estanciaInactiva,
    nombreMes,
    graficoEapb,
    graficoEspecialidad,
    casosActivos,
    cantUrgencias,
    cantHospi,
    graficoMensual,
    graficoCumple // NUEVO
  );
}

function ejecutarBusquedaPrimeraColumna(hoja, rango, valorBuscado) {
  var valores = hoja.getRange(rango).getValues();
  var tiemposPromedio = [];
  var cantUrgencias = null;
  var cantHospi = null;

  for (var i = 0; i < valores.length; i++) {
    if (valores[i][0] == valorBuscado) {
      tiemposPromedio.push(["Ingreso a CDO - Definición de cirugía", parseFloat(valores[i][1]).toFixed(1)]);
      tiemposPromedio.push(["Definición de cirugía - Egreso de CDO", parseFloat(valores[i][2]).toFixed(1)]);
      tiemposPromedio.push(["Egreso de CDO - Procedimiento quirúrgico", parseFloat(valores[i][3]).toFixed(1)]);
      cantUrgencias = parseInt(valores[i][5]);
      cantHospi = parseInt(valores[i][4]);
      break;
    }
  }
  return {tiemposPromedio, cantUrgencias, cantHospi};
}

function ejecutarBusquedaPrimeraFila(hoja, rango, valorBuscado) {
  var valores = hoja.getRange(rango).getValues();
  var datosGrafico = [];
  var columnaEncontrada = valores[0].findIndex(val => val === valorBuscado);

  if (columnaEncontrada >= 0) {
    for (var i = 1; i < valores.length; i++) {
      var nombreVariable = valores[i][0];
      if (nombreVariable) {
        var valorVariable = parseFloat(valores[i][columnaEncontrada]);
        if (!isNaN(valorVariable)) {
          datosGrafico.push([nombreVariable, valorVariable]);
        }
      }
    }
    return datosGrafico;
  }
  return [];
}

function ejecutarBusquedaEstanciaInactiva(hoja, rango, valorBuscado) {
  var valores = hoja.getRange(rango).getValues();
  for (var i = 0; i < valores.length; i++) {
    if (valores[i][0] == valorBuscado) {
      return parseFloat(valores[i][1]).toFixed(1);
    }
  }
  return null;
}

/**
 * Construye un gráfico PIE en HOJA TEMPORAL y devuelve el blob (no toca "Correo").
 * Mantiene firma original pero ignora rangoInicio.
 * Aplica estilos: fuente Roboto, color de texto #0c1c3a, título Roboto 24.
 */
function crearYExportarGrafico(hoja, datos, titulo, rangoInicio) {
  if (!datos || datos.length === 0) return null;

  var ss = hoja.getParent();
  var tmpName = "__tmp_email_charts_" + Utilities.getUuid().slice(0, 8);
  var tmpSheet;
  var blob = null;

  try {
    tmpSheet = ss.insertSheet(tmpName);

    var values = [["Categoría", "Valor"]];
    datos.forEach(function (fila) {
      values.push([String(fila[0]), Number(fila[1])]);
    });
    tmpSheet.getRange(1, 1, values.length, 2).setValues(values);

    var chart = tmpSheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(tmpSheet.getRange(1, 1, values.length, 2))
      .setOption("title", titulo)
      .setOption("titleTextStyle", { fontName: "Roboto", fontSize: 24, color: "#0c1c3a" })
      .setOption("legend", { textStyle: { fontName: "Roboto", fontSize: 18, color: "#0c1c3a" } })
      .setOption("pieSliceTextStyle", { fontName: "Roboto", fontSize: 18, color: "#0c1c3a" })
      .setPosition(1, 4, 0, 0)
      .build();

    tmpSheet.insertChart(chart);
    blob = chart.getBlob();
  } catch (e) {
    console.error("Error creando gráfico temporal (PIE):", e);
  } finally {
    if (tmpSheet) {
      ss.deleteSheet(tmpSheet);
    }
  }

  return blob;
}

/**
 * Lee "Base de datos" y devuelve matriz [["Mes","Cantidad"], ...] para el año dado.
 * - Col B: Grupo_Estado (filtrar "Completado")
 * - Col C: Mes ("YYYYMM NombreMes", p.ej. "202508 Agosto")
 * - Solo año en curso (prefijo YYYY)
 */
function obtenerCasosCompletadosPorMes(spreadSheet, año) {
  var hojaBD = spreadSheet.getSheetByName("Base de datos");
  if (!hojaBD) return [["Mes","Cantidad"]];

  var lastRow = hojaBD.getLastRow();
  if (lastRow < 2) return [["Mes","Cantidad"]];

  // Tomamos columnas B y C desde fila 2
  var valores = hojaBD.getRange(2, 2, lastRow - 1, 2).getValues(); // [ [Grupo_Estado, Mes], ... ]

  // Inicializar conteo por mes 1..12
  var conteo = new Array(12).fill(0);

  valores.forEach(function (fila) {
    var grupo = fila[0];
    var mesStr = (fila[1] || "").toString().trim(); // "YYYYMM NombreMes"
    if (grupo === "Completado" && mesStr.startsWith(String(año)) && mesStr.length >= 6) {
      var mm = parseInt(mesStr.slice(4, 6), 10); // 01..12
      if (!isNaN(mm) && mm >= 1 && mm <= 12) {
        conteo[mm - 1] += 1;
      }
    }
  });

  var mesesNombres = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
  var datos = [["Mes","Cantidad"]];
  for (var i = 0; i < 12; i++) {
    datos.push([mesesNombres[i], conteo[i]]);
  }
  return datos;
}

/**
 * PIE chart en hoja temporal, con estilos unificados.
 */
function crearYExportarGrafico(hoja, datos, titulo, rangoInicio) {
  if (!datos || datos.length === 0) return null;

  var ss = hoja.getParent();
  var tmpName = "__tmp_email_charts_" + Utilities.getUuid().slice(0, 8);
  var tmpSheet;
  var blob = null;

  try {
    tmpSheet = ss.insertSheet(tmpName);

    var values = [["Categoría", "Valor"]];
    datos.forEach(function (fila) {
      values.push([String(fila[0]), Number(fila[1])]);
    });
    tmpSheet.getRange(1, 1, values.length, 2).setValues(values);

    var chart = tmpSheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(tmpSheet.getRange(1, 1, values.length, 2))
      .setOption("title", titulo)
      .setOption("titleTextStyle", { fontName: "Roboto", fontSize: 18, color: "#0c1c3a" })
      .setOption("legend", { textStyle: { fontName: "Roboto", fontSize: 10, color: "#0c1c3a" } })
      .setOption("pieSliceTextStyle", { fontName: "Roboto", fontSize: 10, color: "#0c1c3a" })
      .setPosition(1, 4, 0, 0)
      .build();

    tmpSheet.insertChart(chart);
    blob = chart.getBlob();
  } catch (e) {
    console.error("Error creando gráfico temporal (PIE):", e);
  } finally {
    if (tmpSheet) {
      ss.deleteSheet(tmpSheet);
    }
  }

  return blob;
}

// NUEVO: % de "Cumple programación" por mes (año en curso), ignorando vacíos.
function obtenerPorcentajeCumplimientoPorMes(spreadSheet, año) {
  var hojaBD = spreadSheet.getSheetByName("Base de datos");
  if (!hojaBD) return [["Mes","% Cumplimiento"]];

  var lastRow = hojaBD.getLastRow();
  var lastCol = hojaBD.getLastColumn();
  if (lastRow < 2) return [["Mes","% Cumplimiento"]];

  var colMes = hojaBD.getRange(2, 3, lastRow - 1, 1).getValues();
  var colCumple = hojaBD.getRange(2, lastCol, lastRow - 1, 1).getValues();

  var tot = new Array(12).fill(0);
  var ok  = new Array(12).fill(0);

  for (var i = 0; i < colMes.length; i++) {
    var m = (colMes[i][0] || "").toString().trim();
    var v = colCumple[i][0];
    if (!m || m.length < 6) continue;
    if (!m.startsWith(String(año))) continue;

    var mm = parseInt(m.slice(4, 6), 10);
    if (isNaN(mm) || mm < 1 || mm > 12) continue;

    if (v === "" || v === null) continue;

    tot[mm - 1] += 1;

    var s = (typeof v === "boolean") ? v : v.toString().trim().toLowerCase();
    var esSi = (s === true) || (s === "sí") || (s === "si") || (s === "yes");
    if (esSi) ok[mm - 1] += 1;
  }

  var mesesNombres = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
  var datos = [["Mes","% Cumplimiento","Etiqueta"]]; // añadimos columna para la etiqueta

  for (var i = 0; i < 12; i++) {
    var pct = (tot[i] > 0) ? (ok[i] * 100 / tot[i]) : 0;
    var pctRedondeado = Math.round(pct); // sin decimales
    datos.push([mesesNombres[i], pct, pctRedondeado + "%"]); // numérico + texto para etiqueta
  }
  return datos;
}

function crearYExportarGraficoBarrasCumplimiento(spreadSheet, datos, titulo) {
  if (!datos || datos.length <= 1) return null;

  // Tomamos solo las dos primeras columnas (Mes y %)
  var data = datos.map(function (r) { return [r[0], r[1]]; });

  // Quitar título del eje horizontal
  data[0][0] = "";

  // Redondear a enteros
  for (var i = 1; i < data.length; i++) {
    data[i][1] = Math.round(Number(data[i][1]) || 0);
  }

  var tmpName = "__tmp_email_bar_pct_" + Utilities.getUuid().slice(0, 8);
  var tmpSheet, blob = null;

  try {
    tmpSheet = spreadSheet.insertSheet(tmpName);
    // Escribimos exactamente 2 columnas: Mes | %
    tmpSheet.getRange(1, 1, data.length, 2).setValues(data);

    // Formato de número de la serie: 0% literal sin escalar (no convierte 75 en 75/100)
    // En Apps Script, el % literal puede ponerse como 0"%" (más estable).
    tmpSheet.getRange(2, 2, data.length - 1, 1).setNumberFormat('0"%"');

    var chart = tmpSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(tmpSheet.getRange(1, 1, data.length, 2)) // solo 2 columnas (categoría + serie)
      .setOption("title", titulo)
      .setOption("titleTextStyle", { fontName: "Roboto", fontSize: 18, color: "#0c1c3a" })
      .setOption("legend", { position: "none", textStyle: { fontName: "Roboto", fontSize: 10, color: "#0c1c3a" } })
      .setOption("hAxis", { slantedText: true, textStyle: { fontName: "Roboto", fontSize: 10, color: "#0c1c3a" } })
      .setOption("vAxis", { minValue: 0, viewWindow: { max: 100 }, textStyle: { fontName: "Roboto", fontSize: 10, color: "#0c1c3a" } })
      .setOption("series", { 0: { color: "#0c1c3a", dataLabel: "value" } }) // etiquetas = valor de la serie
      .setPosition(1, 4, 0, 0)
      .build();

    tmpSheet.insertChart(chart);
    blob = chart.getBlob();
  } catch (e) {
    console.error("Error creando gráfico cumplimiento (%):", e);
  } finally {
    if (tmpSheet) {
      spreadSheet.deleteSheet(tmpSheet);
    }
  }

  return blob;
}




/**
 * COLUMN chart en hoja temporal, con estilos unificados y etiquetas visibles.
 * (Quita el título del eje horizontal dejando vacío el encabezado "Mes")
 */
function crearYExportarGraficoBarrasTemporal(spreadSheet, datos, titulo) {
  if (!datos || datos.length <= 1) return null; // encabezado + al menos un dato

  var tmpName = "__tmp_email_bar_" + Utilities.getUuid().slice(0, 8);
  var tmpSheet, blob = null;

  try {
    tmpSheet = spreadSheet.insertSheet(tmpName);

    // Clonamos datos y quitamos el encabezado del eje X ("Mes") para evitar el espacio
    var data = datos.map(function (r) { return r.slice(); });
    data[0][0] = ""; // sin título de eje horizontal

    // Escribimos 2 columnas: ["", "Cantidad"]
    var rng = tmpSheet.getRange(1, 1, data.length, 2);
    rng.setValues(data);

    var chart = tmpSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(tmpSheet.getRange(1, 1, data.length, 2))
      .setOption("title", titulo)
      .setOption("titleTextStyle", { fontName: "Roboto", fontSize: 18, color: "#0c1c3a" })
      .setOption("legend", { position: "none", textStyle: { fontName: "Roboto", fontSize: 10, color: "#0c1c3a" } })
      .setOption("hAxis", { slantedText: true, textStyle: { fontName: "Roboto", fontSize: 10, color: "#0c1c3a" } })
      .setOption("vAxis", { minValue: 0, textStyle: { fontName: "Roboto", fontSize: 10, color: "#0c1c3a" } })
      .setOption("series", { 0: { color: "#60605f", dataLabel: "value" } })
      .setOption("annotations", { alwaysOutside: true, textStyle: { fontName: "Roboto", fontSize: 10, color: "#0c1c3a" } })
      .setPosition(1, 4, 0, 0)
      .build();

    tmpSheet.insertChart(chart);
    blob = chart.getBlob();
  } catch (e) {
    console.error("Error creando gráfico temporal (COLUMN con estilos):", e);
  } finally {
    if (tmpSheet) {
      spreadSheet.deleteSheet(tmpSheet);
    }
  }

  return blob;
}





function enviarInformeCorreo(cantidadCasos, tiemposPromedio, estanciaInactiva, nombreMes, blobGrafico1, blobGrafico2, casosActivos, cantUrgencias, cantHospi, blobGraficoMensual, blobGraficoCumple) {
  var destinatario1 = "alexander.torres@clinicadeoccidente.com";
  var destinatario = "alexander.torres@clinicadeoccidente.com, jessica.parra@clinicadeoccidente.com, claudia.nunez@clinicadeoccidente.com, cinthy.castillo@clinicadeoccidente.com, coord.enfermeriacx@clinicadeoccidente.com, niyi.ruiz@clinicadeoccidente.com, johana.carreno@clinicadeoccidente.com, fernando.trivino@clinicadeoccidente.com, deivy.canaveral@clinicadeoccidente.com, juan.ballesteros@clinicadeoccidente.com, jimmy.folleco@clinicadeoccidente.com, beatriz.vallecilla@clinicadeoccidente.com, jhonatan.arias@clinicadeoccidente.com,geovany.rolon@clinicadeoccidente.com, piedad.gonzalez@clinicadeoccidente.com, khaterine.arteaga@clinicadeoccidente.com, jhonn.pena@clinicadeoccidente.com, myriam.ruiz@clinicadeoccidente.com, alejandra.daza@clinicadeoccidente.com, janeth.vasquez@clinicadeoccidente.com, maria.orozco@clinicadeoccidente.com, melissa.arcos@clinicadeoccidente.com, javier.tovar@clinicadeoccidente.com,andres.puertas@clinicadeoccidente.com";
  var asunto = "Informe cirugía ambulatoria priorizada de " + nombreMes;

  var tablaTiempos = "";
  tiemposPromedio.forEach(fila => {
    tablaTiempos += `<tr><td>${fila[0]}</td><td>${fila[1]}</td></tr>`;
  });

  var bloqueGraficas = `
<p style="font-family: Arial, sans-serif; font-size: 14px;">
A continuación, se presentan las gráficas con la distribución de casos por <strong>EAPB</strong>, por <strong>especialidad quirúrgica</strong>, la <strong>cantidad de casos mensual</strong> y el cumplimiento a la programación quirurgicia de los casos completados en el año:</p>
${blobGrafico1 ? '<br><img src="cid:graficoEapb">' : ''}
${blobGrafico2 ? '<br><img src="cid:graficoEspecialidad">' : ''}
${blobGraficoMensual ? '<br><img src="cid:graficoMensual">' : ''}
${blobGraficoCumple ? '<br><img src="cid:graficoCumple">' : ''}
`;

  var mensaje = `
<p style="font-family: Arial, sans-serif; font-size: 14px;">Buenos días,</p>

<p style="font-family: Arial, sans-serif; font-size: 14px;">
Con gusto presento el <strong>informe mensual de resultados</strong> correspondiente al mes de <strong style="color:#2E86C1;">${nombreMes}</strong> del programa de <strong>cirugía ambulatoria priorizada (Cirugía diferida)</strong>:</p>

<ul style="font-family: Arial, sans-serif; font-size: 14px;">
  <li>Se gestionaron un total de <strong style="color:#B03A2E;">${cantidadCasos}</strong> casos durante el mes, ${cantUrgencias} en el servicio de hospitalización y ${cantHospi} en urgencias.</li>
  <li>Se evitaron <strong style="color:#B03A2E;">${estanciaInactiva}</strong> días de estancia inactiva (días entre el egreso del CDO y la realización del procedimiento quirúrgico).</li>
</ul>

<p style="font-family: Arial, sans-serif; font-size: 14px;"><strong style="color:#2471A3;">Tiempos promedio por actividad:</strong></p>

<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 14px;">
  <tr style="background-color: #D6EAF8;">
    <th>Etapa del proceso</th>
    <th>Tiempo promedio (días)</th>
  </tr>
  ${tablaTiempos}
</table>

<br>
${bloqueGraficas}

<br><br>
<p style="font-family: Arial, sans-serif; font-size: 14px;">
<em>Este informe corresponde al corte consolidado al cierre del mes.</em><br>
No obstante, es importante tener en cuenta que <strong>aún existen ${casosActivos} casos activos</strong>, es decir, que fueron egresados y están a la espera de cirugía, que contarán en la estadística en el momento que se lleve a cabo el procedimiento.</p>

<p style="font-family: Arial, sans-serif; font-size: 14px;">
Para visualizar los <strong>indicadores en tiempo real</strong>, puede acceder al siguiente enlace:<br>
<a href="https://docs.google.com/spreadsheets/d/1Hkv8CLofSZPFfTty058hJX0sZGU5r66P_oi7j1tPFMo/edit?gid=1972684983" style="color:#2874A6;" target="_blank">Ver tablero de seguimiento en tiempo real</a></p>

<p style="font-family: Arial, sans-serif; font-size: 14px;">Gracias por su atención.</p>`;

  let inlineImages = {};
  if (blobGrafico1) inlineImages.graficoEapb = blobGrafico1;
  if (blobGrafico2) inlineImages.graficoEspecialidad = blobGrafico2;
  if (blobGraficoMensual) inlineImages.graficoMensual = blobGraficoMensual;
  if (blobGraficoCumple) inlineImages.graficoCumple = blobGraficoCumple;

  MailApp.sendEmail({
    to: destinatario,
    subject: asunto,
    name: "Informes cirugía diferida",
    htmlBody: mensaje,
    inlineImages: inlineImages
  });
}
