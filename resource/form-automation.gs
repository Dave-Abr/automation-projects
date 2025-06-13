function validarCorreo1(){

  // Acceder al Google Sheets del formulario

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Respuestas de formulario 1'); 
  const sheetquery = ss.getSheetByName('Query'); 
  const lastRow = sheet.getLastRow();
  const range = sheetquery.getRange(lastRow, 1, 1, sheetquery.getLastColumn());
  const values = range.getValues();

  // Asignar las carpetas

  const CarpetaPDF      ="1oZZyRBjHp_Uler8Tz_mMtfW5TGIk9Y1h"; 
  const Editable ="1ygJl6jLcbs0Llr4oE-KdBTChDnOTIpbx"; 
  var CarpetaEditable = DriveApp.getFolderById(Editable);

  // Asignar la variables

  const PlantillaDoc = values[0][70];
  const NombreDoc = values[0][71];
  const PlandePagosHTML = values[0][70];
  const ProductosHTML = values[0][71];

  const Nombre = values[0][73];
  const Direccion = values[0][74];
  const Correo = values[0][7];
  const FechaCreacion = values[0][75];
  const FechaVencimiento = values[0][76];
  const Clima = values[0][78];
  const Estiercol = values[0][79];

  const R1 = values[0][81];
  const R2 = values[0][82];
  const R3 = values[0][83];
  const R4 = values[0][84];
  const R5 = values[0][82];
  const R6 = values[0][85];

  const Tamano = values[0][87];

  const P1 = values[0][87];
  const C1 = values[0][88];
  const V1 = values[0][89];
  const T1 = values[0][90];

  const P2 = values[0][91];
  const C2 = values[0][92];
  const V2 = values[0][93];
  const T2 = values[0][94];

  const P3 = values[0][95];
  const C3 = values[0][96];
  const V3 = values[0][97];
  const T3 = values[0][98];

  const Total = values[0][99];

  const Anticipo = values[0][101];
  const AnticipoP = values[0][102];
  const Instalacion = values[0][103];
  const InstalacionP = values[0][104];
  const Credito = values[0][105];
  const CreditoP = values[0][106];
  const NumeroCuotas = values[0][107];
  const Cuota = values[0][108];
  const CuotaP = values[0][119];

  const ActivoVendedor = values[0][65];
  const NombreVendedor = values[0][66];
  const CargoVendedor = values[0][67];
  const CelularVendedor = values[0][68];
  const CorreoVendedor = values[0][1];
  const EnlaceWhatsApp = values[0][69];

  const Comentario = values[0][112];
  const NumeroenLetras = values[0][113];
  const htmlproductos = values[0][117];
  const htmlpagos = values[0][118];
  const EnviaralCliente = values[0][119];
  const editable = values[0][120];

  // Validad si es vendedor está activo

  if (ActivoVendedor == "Activo") {
  
    var carpeta= DriveApp.getFolderById(CarpetaPDF);
    var archivoPlantilla=DriveApp.getFileById(PlantillaDoc);
    var copiaArchivo=archivoPlantilla.makeCopy();
    var idArchivoCopia= copiaArchivo.getId();
    var doc=DocumentApp.openById(idArchivoCopia);
    var texto= doc.getBody();

    texto = texto.replaceText("{{nombre}}", Nombre);
    texto = texto.replaceText("{{direccion}}", Direccion);
    texto = texto.replaceText("{{fechacreacion}}", FechaCreacion);
    texto = texto.replaceText("{{estiercol}}", Estiercol);
    texto = texto.replaceText("{{clima}}", Clima );
    texto = texto.replaceText("{{r1}}", R1);
    texto = texto.replaceText("{{r2}}", R2);
    texto = texto.replaceText("{{r3}}", R3);
    texto = texto.replaceText("{{r4}}", R4);
    texto = texto.replaceText("{{r5}}", R5);
    texto = texto.replaceText("{{r6}}", R6);
    texto = texto.replaceText("{{tamano}}", Tamano);
    texto = texto.replaceText("{{P1}}", P1);
    texto = texto.replaceText("{{C1}}", C1);
    texto = texto.replaceText("{{V1}}", V1);
    texto = texto.replaceText("{{T1}}", T1);
    texto = texto.replaceText("{{P2}}", P2);
    texto = texto.replaceText("{{C2}}", C2);
    texto = texto.replaceText("{{V2}}", V2);
    texto = texto.replaceText("{{T2}}", T2);
    texto = texto.replaceText("{{P3}}", P3);
    texto = texto.replaceText("{{C3}}", C3);
    texto = texto.replaceText("{{V3}}", V3);
    texto = texto.replaceText("{{T3}}", T3);
    texto = texto.replaceText("{{total}}", Total);
    texto = texto.replaceText("{{fechavencimiento}}", FechaVencimiento);
    texto = texto.replaceText("{{anticipo}}", Anticipo);
    texto = texto.replaceText("{{antp}}", AnticipoP);
    texto = texto.replaceText("{{instalacion}}", Instalacion);
    texto = texto.replaceText("{{insp}}", InstalacionP);
    texto = texto.replaceText("{{creditop}}", CreditoP);
    texto = texto.replaceText("{{credito}}", Credito);
    texto = texto.replaceText("{{cu1}}", Cuota);
    texto = texto.replaceText("{{ncuotas}}", NumeroCuotas);
    texto = texto.replaceText("{{nombrevendedor}}", NombreVendedor);
    texto = texto.replaceText("{{cargovended}}", CargoVendedor);
    texto = texto.replaceText("{{celularvendedor}}", CelularVendedor);
    texto = texto.replaceText("{{correovendedor}}", CorreoVendedor);
    texto = texto.replaceText("{{numeroenletras}}", NumeroenLetras);
    texto = texto.replaceText("{{comendatarios}}", Comentario);


    MexicoLargaBiodigestor= "1DntY42yAqzuDswXbU7pLB448_Wykx-j9UJE55gDqAEU";
    MexicoLargaAccesorios = "19OQvwWjOl_I4xJ_-QouSX0CyixXE10rVmXF1s2Xotck";
    MexicoCortaBiodigestor = "1iytfDkLKt0VTC7G8sN2i8ywHinXnkOWOJpXBB1Ri7l4";
    MexicoCortaAccesorios= "1iytfDkLKt0VTC7G8sN2i8ywHinXnkOWOJpXBB1Ri7l4";
    ColombiaLargaBiodigestor= "1DntY42yAqzuDswXbU7pLB448_Wykx-j9UJE55gDqAEU";
    ColombiaLargaAccesorio= "";
    ColombiaCortaBiodigestor= "12376AcxV11YGy1DC9yo_Jpkt-7W83roGb59lfcl_vFk";
    ColombiaCortaAccesorio= "12376AcxV11YGy1DC9yo_Jpkt-7W83roGb59lfcl_vFk";

    // Validad la platilla que se utilizará

    if (PlantillaDoc == MexicoLargaBiodigestor || PlantillaDoc == ColombiaLargaBiodigestor){
      
      var tablas = texto.getTables();
      eliminarTablasPorProducto(P1, tablas);
      

    }

    doc.saveAndClose();
    
    var pdfFile = generarPDF(idArchivoCopia, carpeta, NombreDoc);
    var PDFurl = pdfFile.getUrl();

    sheetquery.getRange(lastRow, 122).setValue(PDFurl);


    var textoHtml=HtmlService.createHtmlOutputFromFile("HTML Vendedor").getContent();

    textoHtml=textoHtml.replace("{{nombre}}",Nombre);
    textoHtml=textoHtml.replace("{{direccion}}",Direccion);
    textoHtml=textoHtml.replace("{{total}}",Total);
    textoHtml=textoHtml.replace("{{productoshtml}}",htmlproductos);
    textoHtml=textoHtml.replace("{{plandepagoshtml}}",htmlpagos);
    textoHtml=textoHtml.replace("{{C1}}",C1);
    textoHtml=textoHtml.replace("{{C2}}",C2);
    textoHtml=textoHtml.replace("{{C3}}",C3);
    textoHtml=textoHtml.replace("{{P1}}",P1);
    textoHtml=textoHtml.replace("{{P2}}",P2);
    textoHtml=textoHtml.replace("{{P3}}",P3);
    textoHtml=textoHtml.replace("{{anticipo}}",Anticipo);
    textoHtml=textoHtml.replace("{{anticipop}}",AnticipoP);
    textoHtml=textoHtml.replace("{{instalacion}}",Instalacion);
    textoHtml=textoHtml.replace("{{instalacionp}}",InstalacionP);
    textoHtml=textoHtml.replace("{{credito}}",Credito);
    textoHtml=textoHtml.replace("{{creditop}}",CreditoP);
    textoHtml=textoHtml.replace("{{cuotas}}",NumeroCuotas);
    textoHtml=textoHtml.replace("{{plandepagoshtml}}",PlandePagosHTML);

    if (editable == "Editable"){

      var copiaTemporal = DriveApp.getFileById(idArchivoCopia).makeCopy(NombreDoc, CarpetaEditable); // Crear copia en la carpeta correcta
      var idCopiaEditable = copiaTemporal.getId();
      DriveApp.getFileById(idArchivoCopia).setTrashed(true); 
      var archivoTemporal = DriveApp.getFileById(idCopiaEditable);
      archivoTemporal.setName(NombreDoc);
      DriveApp.getFileById(idCopiaEditable).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
      var DOCURL = archivoTemporal.getUrl();

      textoHtml=textoHtml.replace("{{1}}","<br>Enlace del documento editable:<br>");
      textoHtml=textoHtml.replace("{{2}}","<u><a>{{enlace}}</a></u><br><br>");
      textoHtml=textoHtml.replace("{{enlace}}",DOCURL);

      
      enviarCorreo(CorreoVendedor, NombreDoc, textoHtml, pdfFile);

    } else {
      textoHtml=textoHtml.replace("{{1}}","");
      textoHtml=textoHtml.replace("{{2}}","");
      enviarCorreo(CorreoVendedor, NombreDoc, textoHtml, pdfFile);
      if (EnviaralCliente == "Enviar al Cliente") {
        
        var textoHtml2=HtmlService.createHtmlOutputFromFile("HTML Cliente").getContent();
        textoHtml2=textoHtml2.replace("{{nombre}}",Nombre);
        textoHtml2=textoHtml2.replace("{{enlacewhatsapp}}",EnlaceWhatsApp);
        textoHtml2=textoHtml2.replace("{{fotowhatsapp}}",FotoWhatApp);
        var message = '¡Cotización Biodigestores!';
        const Correoventas = "ventas@sistema.bio";
        const nombreRemitente="Sistema.bio"
        const copiaOculta = "davida@sistema.bio";
        const asuntoo="Cotización de Biodigestores"
        
        GmailApp.sendEmail(Correo, asuntoo, message, {
          from: Correoventas,
          name: nombreRemitente,
          cc: CorreoVendedor,                 // dirección de correo para la copia
          attachments: [pdfFile.getAs(MimeType.PDF)],
          //bcc: copiaOculta,
          htmlBody:textoHtml2
        });
      }
      return;

    } 

   //Si el vendedor no esta activo

  } else if (ActivoVendedor == "Create Contact"){
    GmailApp.sendEmail(CorreoVendedor, "No estas registrado en este formulario", "Solicita que te registren a este formulario. Responde a este correo.", ); 
  }else if (ActivoVendedor == "Inactivo"){
    GmailApp.sendEmail(CorreoVendedor, "Inactivo en el formulario", "Solicita que te activen a este formulario. Responde a este correo.", ); 
  }
}




function eliminarTablaPorOrden(tablas,numeroTabla) {
  if (tablas.length > numeroTabla) {
        tablas[numeroTabla].removeFromParent();
        Logger.log("Tabla eliminada:"+ numeroTabla);
      }   else {
        Logger.log("No hay tablas en el documento:",numeroTabla);
      }
}

function generarPDF(idArchivoCopia, carpeta, NombreDoc) {
  var pdf = DriveApp.getFileById(idArchivoCopia).getAs("application/pdf");
  var pdfFile = carpeta.createFile(pdf).setName(NombreDoc);
  return pdfFile; // Retorna el archivo PDF para su uso
}

function eliminarTemporal(idArchivoCopia) {
  DriveApp.getFileById(idArchivoCopia).setTrashed(true); // Mueve el archivo a la papelera
  Logger.log("Archivo eliminado: " + idArchivoCopia); // Opcional: Log para depuración
}

function enviarCorreo(destinatarios, asunto, cuerpoHtml, pdfFile) {
  GmailApp.sendEmail(destinatarios, asunto, "Aquí está tu PDF", {
    attachments: [pdfFile.getAs(MimeType.PDF)], // Adjuntar el PDF
    htmlBody: cuerpoHtml // Cuerpo en HTML
  });
}

function validarVendedorActivo(values) {
  const ActivoVendedor = values[0][65];
  if (ActivoVendedor !== "Activo") {
    Logger.log("El vendedor no está activo. Se detiene la ejecución.");
    return false;
  }
  return true;
}


function eliminarTablasPorProducto(P1, tablas) {
    const tablasPorProducto = {
        "Biodigestor Sistema 8":  [11, 10, 9, 8, 7, 6, 5, 4, 3, 2],
        "Biodigestor Sistema 12": [11, 10, 9, 8, 7, 6, 5, 4, 3, 1],
        "Biodigestor Sistema 16": [11, 10, 9, 8, 7, 6, 5, 4, 2, 1],
        "Biodigestor Sistema 20": [11, 10, 9, 8, 7, 6, 5, 3, 2, 1],
        "Biodigestor Sistema 30": [11, 10, 9, 8, 7, 6, 4, 3, 2, 1],
        "Biodigestor Sistema 40": [11, 10, 9, 8, 7, 5, 4, 3, 2, 1],
        "Biodigestor Sistema 80": [11, 10, 9, 8, 6, 5, 4, 3, 2, 1],
        "Biodigestor Sistema 120":[11, 10, 9, 7, 6, 5, 4, 3, 2, 1],
        "Biodigestor Sistema 160":[11, 10, 8, 7, 6, 5, 4, 3, 2, 1],
        "Biodigestor Sistema 200":[11, 9, 8, 7, 6, 5, 4, 3, 2, 1],
        "Biodigestor Sistema 400":[10, 9, 8, 7, 6, 5, 4, 3, 2, 1]
    };

    if (P1 in tablasPorProducto) {
        tablasPorProducto[P1].forEach(index => eliminarTablaPorOrden(tablas, index));
    }
}
