// Función para ejecutar cuando se envía el formulario
function onFormSubmit(e) {
    // Obtener la respuesta del formulario
    var respuesta = e.response.getItemResponses();
  
    // Abrir la hoja de cálculo donde se almacenarán las respuestas
    var libroDeCalculo = SpreadsheetApp.openById('ID Hoja de calculo');
    var hojaRespuestas = libroDeCalculo.getSheetByName('Respuestas formulario');

    // Abrir el documento donde se encuentra la plantilla
    var templateFile = DriveApp.getFileById('ID Documento de plantilla');

    // Hacer una copia de la plantilla para generar el nuevo documento
    var tempFile = templateFile.makeCopy('TempCopy');
    
    // Obtener el documento recién creado
    var doc = DocumentApp.openById(tempFile.getId());
    var cuerpo = doc.getBody();
  
    // Reemplazar los marcadores de posición en la plantilla con las respuestas del formulario
    cuerpo.replaceText('{Nombre}', respuesta[0].getResponse());
    cuerpo.replaceText('{ApellidoPaterno}', respuesta[1].getResponse());
    cuerpo.replaceText('{ApellidoMaterno}', respuesta[2].getResponse());
    cuerpo.replaceText('{Matricula}', respuesta[3].getResponse());
    cuerpo.replaceText('{ProgramaEducativo}', respuesta[4].getResponse());
    cuerpo.replaceText('{NSS}', respuesta[5].getResponse());
    
    if( respuesta[7].getResponse() == "Programa de prácticas profesionales (PPP)." ){
      cuerpo.replaceText('{PPP/PVVC}', "PPP");
      cuerpo.replaceText('{PPP/PVVC2}', "Programa de prácticas profesionales (PPP)");
    }else{
      cuerpo.replaceText('{PPP/PVVC}', "PVVC");
      cuerpo.replaceText('{PPP/PVVC2}', "Proyecto de vinculación con valor en créditos (PVVC)");
    }

    cuerpo.replaceText('{NombreUR}', respuesta[8].getResponse());
    cuerpo.replaceText('{NombreDestinatario}', respuesta[9].getResponse());
    cuerpo.replaceText('{ProfesionDestinatario}', respuesta[10].getResponse());
    cuerpo.replaceText('{PuestoDestinatario}', respuesta[11].getResponse());

    cuerpo.replaceText('{FechaInicio}', respuesta[12].getResponse());
    cuerpo.replaceText('{FechaTermino}', respuesta[13].getResponse());

    cuerpo.replaceText('{HoraInicio}', respuesta[14].getResponse());
    cuerpo.replaceText('{HoraSalida}', respuesta[15].getResponse());
    //16 respuestas (0-15)

    cuerpo.replaceText('{Fecha}', obtenerFechaActual() );

    // Obtener el valor actual del contador
    // Reemplazar el marcador de posición en la plantilla con el número actual del contador
    var counterSheet = libroDeCalculo.getSheetByName('Contador');
    
    if( respuesta[7].getResponse() == "Programa de prácticas profesionales (PPP)." ){
      var counter = counterSheet.getRange("B1").getValue();
      counter++;
      if( counter < 100 ){
        var counterValue = '0' + counter;
        cuerpo.replaceText("{Contador}", counterValue);
      }else{
        cuerpo.replaceText("{Contador}", counter);
      }
      counterSheet.getRange("B1").setValue(counter);
    }else{
      var counter = counterSheet.getRange("B2").getValue();
      counter++;
      if( counter < 100 ){
        var counterValue = '0' + counter;
        cuerpo.replaceText("{Contador}", counterValue);
      }else{
        cuerpo.replaceText("{Contador}", counter);
      }
      counterSheet.getRange("B2").setValue(counter);
    }

    doc.saveAndClose();
    var fechaActual = Utilities.formatDate(new Date(), "GMT-7", "dd-MM-yyyy");

    // Guardar una copia del documento como archivo PDF en Google Drive
    var nombreArchivo = fechaActual + ' Carta presentación - ' + respuesta[0].getResponse() + ' ' + respuesta[1].getResponse() + ' - ' + respuesta[3].getResponse() +'.pdf';
    var carpeta = DriveApp.getFolderById('ID Carpeta Drive');
    var urlPDF = carpeta.createFile(tempFile.getAs('application/pdf')).setName(nombreArchivo).getUrl();

    // Insertar las respuestas en la hoja de cálculo
    var fila = [];
    fila.push( fechaActual );

    for (var i = 0; i < respuesta.length; i++) {
        if (i == 6){

            var fileId = respuesta[6].getResponse();
            var file = DriveApp.getFileById(fileId);
            var pdfFileUrl = file.getUrl();
            fila.push( pdfFileUrl );

        }else{
            fila.push(respuesta[i].getResponse());
        }
    }
    fila.push(urlPDF);
    hojaRespuestas.appendRow(fila);

    // Eliminar el archivo temporal de Google Docs
    //DriveApp.removeFile(tempFile);
    DriveApp.getFileById(tempFile.getId()).setTrashed(true);

    //--------------------------------------------------------------------------------------------------

    var urlLibroDeCalculo = libroDeCalculo.getUrl();
    
    var correoDestino = "correo@mail.com"; 
    var asunto = "Solicitud de carta de presentación" + ' - ' + respuesta[1].getResponse() + ' ' + respuesta[2].getResponse() + ' ' + respuesta[0].getResponse();
    var mensaje = respuesta[0].getResponse() + ' ' + respuesta[1].getResponse() + ' ' + respuesta[2].getResponse() + " ha solicitado una carta de presentación.\n\nVer datos de solicitud:\n" + urlLibroDeCalculo + "\n\nArchivo de vigencia:\n" + pdfFileUrl + "\n\nRevisar carta creada:\n" + urlPDF;

    MailApp.sendEmail(correoDestino, asunto, mensaje);

  //--------------------------------------------------------------------------------------------------
    
    // Mostrar la URL del archivo PDF en el registro
    //Logger.log('El archivo PDF ha sido guardado en: ' + urlPDF);
  }
  
/*Reemplazar las ID y nombres de las hojas de cálculo, documentos y carpetas con los 
tuyos propios. Además, debes asegurarte de que el script tenga los permisos necesarios para acceder a los archivos 
de Google Drive, Google Docs y Google Sheets.*/

function obtenerFechaActual() {
    var fechaActual = new Date();
    var dia = fechaActual.getDate();
    var mes = fechaActual.toLocaleString('es', { month: 'long' }); //obtiene el mes en formato texto
    var anio = fechaActual.getFullYear();
    var fechaConFormato = dia + ' de ' + mes + ' del ' + anio;
    return fechaConFormato;
}


