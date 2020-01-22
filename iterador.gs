var sapp,ss,dapp,rev,mpp,sheet,params,results,results_sheet,properties,continuationToken,headers,title;
    dapp = DriveApp;
    sapp = SpreadsheetApp;
    ss = sapp.getActiveSpreadsheet();
    sheet = ss.getSheetByName('Operaciones');
    params = sheet.getRange('C3:E3').getValues();
    oferta = 'DISTANCIA';
    folderId = 'folderID';
    ope = params[0][0];
    results = [];
    results_sheet = ss.getSheetByName('Resultados');
    properties = PropertiesService.getUserProperties(); //Lo uso para guardar variables globales
    continuationToken = properties.getProperty('CONTINUATION_TOKEN');       
var date = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss");
var start_time,runtime,safe_runtime; 
    start_time = new Date();
    runtime = 0;
    safe_runtime = 150; //2.5 minuticos
var checkToken,button1,button2;   
  
function startFI() {  
  var template = HtmlService.createTemplateFromFile('sideBar').getRawContent();
  var contToken = PropertiesService.getUserProperties().getProperty('CONTINUATION_TOKEN');
  if (contToken === null) {checkToken = ''; button1 = 'Ejecutar'; button2 = ''}
  else {checkToken = 'Operación anterior incompleta, token pendiente.'; button1 = 'Continuar'; button2 = '<button id="Botón" type="reset" class="green" onclick="reset()">Reset</button>'}  
  var ui = HtmlService.createHtmlOutput(template.replace("{disp}",checkToken).replace("{main_button}",button1).replace("{reset_button}",button2)).setTitle('Operador');
  sapp.getUi().showSidebar(ui);
};

//-------------------------------------------------------------------ITERADOR DE ARCHIVOS, INICIO-------------------------------------------------------------------
   
function fileIterator(continuationToken){

  if(continuationToken === undefined){ //Primer Batch
    var files = dapp.getFolderById(folderId).getFiles();
    properties.setProperty('Counter', 0); //defino como variable global un contador
    properties.setProperty('ERROR_COUNTER', 0); //contador de errores
     switch (ope){ //acciones previas a la operacion
      case '1 -': //Cerrar
       headers = ['Nombre','URL','Cerrada "ESTUDIANTES" y Abierta "Fuera de cambios"'];break;
      case '2 -': //Abrir
       headers = ['Nombre','URL','Abierta "ESTUDIANTES" y Cerrada "Fuera de cambios"'];break;
      case '3 -': //Cambios
       headers = ['Nombre','URL','Incluida en "Informe fuera de cambios"']; generarPlanillaInformeDeCambios();break;
      case '4 -': //CSVs moodle
       headers = ['Nombre','URL','CSV generado']; generarPlanillaMoodleCSV(); vaciarCarpetaCSV();break;
      case '5 -': //Cruce Drive Vs Moodle
       headers = ['Nombre','URL','Referencia generada']; generarPlanillaReferenciaCruce(); showPicker();
     }
     results_sheet.clearContents(); results.push(headers);
  } 
  else{ //Batchs siguientes
    var files = dapp.continueFileIterator(continuationToken);
     switch (ope){ //acciones previas a cada batch subsiguiente
      case '3 -': //Cambios
      abrirPlanillaInformeDeCambios();break;
      case '4 -': //CSVs moodle
      abrirPlanillaMoodleCSV();break;
      case '5 -': //Cruce Drive Vs Moodle
      abrirPlanillaReferenciaCruce();
     }
  }
  
 var it = 0;
  while (files.hasNext()){
    var file = files.next(); 
//----------------la comida va aqui----------------
  switch (ope){
     case '1 -': //Cerrar
      results.push(cerrarPlanillas(file));break;
     case '2 -': //Abrir
      results.push(abrirPlanillas(file));break;
     case '3 -': //Cambios
      results.push(informeDeCambios(file));break;
     case '4 -': //CSVs moodle
      results.push(moodleCSV(file));break;
     case '5 -': //Cruce Drive Vs Moodle 
      results.push(generarReferenciaCruce(file));
    }  
//-----------------------------------------------
    runtime = ((new Date()) - start_time)/1000;
    it++;   
    if(runtime > safe_runtime){
     if(files.hasNext()){ //Si hay mas archivos, genera token para el proximo bacth
       continuationToken = files.getContinuationToken();
       PropertiesService.getUserProperties().setProperty('CONTINUATION_TOKEN', continuationToken);
       PropertiesService.getUserProperties().setProperty('Counter', parseInt((PropertiesService.getUserProperties().getProperty('Counter'))) + parseInt(it)); //sumo el contador
       var last = results_sheet.getLastRow();
       results_sheet.getRange(parseInt(last) + 1, 1, results.length, results[0].length).setValues(results);
       return continuationToken;
     } 
    }
  } //FIN DEL WHILE
  if (files.hasNext() == false) {  //Si NO hay mas archivos, condicional redundante (si llego aca no habia mas archivos)
    switch (ope){ //para operaciones adicionales posteriores a la iteracion sobre la carpeta
      case '5 -': continuationToken = 'OPERACION_ADICIONAL';break;
      default:    continuationToken = undefined; 
      }
    PropertiesService.getUserProperties().deleteProperty('CONTINUATION_TOKEN');  
    PropertiesService.getUserProperties().setProperty('Counter', parseInt((PropertiesService.getUserProperties().getProperty('Counter'))) + parseInt(it)); //sumo el contador
    var last = results_sheet.getLastRow();
    results_sheet.getRange(parseInt(last) + 1, 1, results.length, results[0].length).setValues(results);
    sheet.getRange(parseInt(params[0][2]), 6).setValue(new Date());
    sheet.getRange(parseInt(params[0][2]), 8).setValue(parseInt(PropertiesService.getUserProperties().getProperty('ERROR_COUNTER')));
    return continuationToken;
  }     
};
//-------------------------------------------------------------------ITERADOR DE ARCHIVOS, FIN-------------------------------------------------------------------

//----------------------------------------------FUNCIONES AUXILIARES----------------------------------------------

function writeSomething() { return properties.getProperty('Counter');} 

function hard_reset() { 
 PropertiesService.getUserProperties().deleteProperty('CONTINUATION_TOKEN');
 PropertiesService.getUserProperties().deleteProperty('Counter');
 PropertiesService.getUserProperties().deleteProperty('GLOBAL_AUX');
 PropertiesService.getUserProperties().deleteProperty('ERROR_COUNTER');
 results_sheet.clearContents();
}

function reset_token() { 
 PropertiesService.getUserProperties().deleteProperty('CONTINUATION_TOKEN');
}
