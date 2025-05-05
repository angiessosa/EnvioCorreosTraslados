//GOOGLE APPS SCRIPT


function enviarCorreosTraslados() {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TRASLADOS CORREOS");
    const datos = hoja.getDataRange().getValues();
    const encabezados = datos[0];
  
    // INDICES DE COLUMNAS
    const iEmpresa = encabezados.indexOf("EMPRESA");
    const iNit = encabezados.indexOf("NIT");
    const iPoliza = encabezados.indexOf("POLIZA");
    const iNumCaso = encabezados.indexOf("NUMERO CASO");
    const iFechaRadicado = encabezados.indexOf("FECHA RADICADO");
    const iFechaSolicitada = encabezados.indexOf("FECHA SOLICITADA");
    const iLocalidad = encabezados.indexOf("LOCALIDAD");
    const iCanal = encabezados.indexOf("CANAL");
    const iCanalComercial = encabezados.indexOf("CANAL COMERCIAL");
    const iAsesor = encabezados.indexOf("ASESOR");
    const iPygGlobal = encabezados.indexOf("P&G GLOBAL");
    const iUltAporte = encabezados.indexOf("ULTIMO APORTE");
    const iCorreoEnv = encabezados.indexOf("CORREO ENVIADO");
  
    const destinatarios = ["angie.sosa@segurosbolivar.com"];
  
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const correoYaEnviado = fila[iCorreoEnv].toString().toUpperCase() === "SI";
  
      if (!correoYaEnviado) {
        const empresa = fila[iEmpresa];
        const nit = fila[iNit];
        const poliza = fila[iPoliza];
        const numCaso = fila[iNumCaso];
        const fechaRad = formatearFecha(fila[iFechaRadicado]);
        const fechaSoli = formatearFecha(fila[iFechaSolicitada]);
        const localidad = fila[iLocalidad];
        const canal = fila[iCanal];
        const canalComercial = fila[iCanalComercial];
        const asesor = fila[iAsesor];
        const pygGlobal = formatearPorcentaje(fila[iPygGlobal]);
        const ultAporte = formatearMoneda(fila[iUltAporte]);
  
        const asunto = `Traslado ARL Póliza ${poliza} - NIT ${nit} - ${empresa} - CASO ${numCaso}`;
        const mensaje = `Nos permitimos informar que el día ${fechaRad} fue radicada la solicitud de traslado de la empresa ${empresa} con NIT ${nit} y Póliza ${poliza} para surtir efecto a partir del ${fechaSoli}.
  
  Localidad: ${localidad}
  Canal: ${canal}
  Canal Comercial: ${canalComercial}
  Asesor A y C: ${asesor}
  P&G Global: ${pygGlobal}
  Último Aporte: ${ultAporte}
  `;
  
        // ENVIAR CORREO
        MailApp.sendEmail({
          to: destinatarios.join(","),
          subject: asunto,
          body: mensaje
        });
  
        // MARCAR COMO ENVIADO
        hoja.getRange(i + 1, iCorreoEnv + 1).setValue("SI");
      }
    }
  }
  
  //Función para formatear fecha como "dd/MM/yyyy"
  function formatearFecha(valor) {
    const fecha = new Date(valor);
    if (!isNaN(fecha)) {
      // Ajusta la fecha a la zona horaria 'America/Bogota' para evitar desajustes
      const utcOffset = fecha.getTimezoneOffset() * 60000; // Convertir offset a milisegundos
      const fechaLocal = new Date(fecha.getTime() + utcOffset); // Ajustar a UTC
      return Utilities.formatDate(fechaLocal, "America/Bogota", "dd/MM/yyyy");
    }
    return valor;
  }
  
  function formatearPorcentaje(valor) {
    const porcentaje = parseFloat(valor);
    if (!isNaN(porcentaje)) {
      return (porcentaje * 100).toFixed(2).replace('.', ',') + '%';
    }
    return valor;
  }
  
  function formatearMoneda(valor) {
    const numero = parseFloat(valor);
    if (!isNaN(numero)) {
      return '$' + numero.toLocaleString('es-CO');
    }
    return valor;
  }