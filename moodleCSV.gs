var headers = ['username','password','firstname','lastname','email','course1','role1','course2','role2','course3','role3'];
var apellidoCol,nombreCol,dniCol,resolucionCol,soporteCol;
  apellidoCol = 1;
  nombreCol = 2;
  dniCol = 3;
  mailCol = 6  
  resolucionCol = 11;
  soporteCol = 12;
var parentMoodleCSV,folderMoodleCSV,sheetMoodleCSV;
  
//-------------------------------RUTINA PRINCIPAL-------------------------------
function moodleCSV(file) {
try {
 //primero genereo un "consolidado" con todos los estudiantes con la estructura de los csv y los inicios correspondiente
 var fileId = file.getId();
 var open = SpreadsheetApp.openById(fileId);
 var planilla = open.getSheetByName('Resumen').getRange('A9:AS').getValues();
 var sec = open.getSheetByName('ESTUDIANTES').getRange('B15').getValue();
 var url = file.getUrl();
 var array = [];
 
 //iteraciones por planilla
 for (var n in planilla){
  var rw = planilla[n]; //entrada individual de estudiante
  
   if (rw[0] === '') {continue} //line en blanco, saltear
  //resolucion,soporte
   var resolucion = rw[resolucionCol]; var res = resolucion.toString().replace(/\D/g,'').substring(0,3);
   var soporte = rw[soporteCol].toString();
  //conviero cuil/dni a dni y verifico
   var cuilDni = (rw[dniCol].toString()).replace(/\D/g,'');    //solo valores numericos;
   var first = cuilDni.substring(0,1); var med = cuilDni.substring(2,3); var len = cuilDni.length;
    if (len === 8 && (first <= 4 || first == 9))  {var dni = cuilDni}
    else if (len === 7 && first >= 4)             {var dni = cuilDni}
    else if (len === 11 && (med <= 4 || med == 9)){var dni = cuilDni.substring(2,10)}
    else if (len === 10 && med >= 4)              {var dni = cuilDni.substring(2,9)}
    else                                          {var dni = 'invalid'}
  //apellido,nombre,email,regexp(para validar el mail),rol
   var apellido = rw[apellidoCol];
   var nombre = rw[nombreCol];
   var email = rw[mailCol];
   var emailValidation = /\S+@\S+\.\S+/;
   var rol = 'student';
   
  //SI EL DNI ES VALIDO, EL MAIL ES VALIDO, LA RESOLUCION ES 106 Y EL SOPORTE ES VIRTUAL: LE ASIGNO LOS INICIOS
   if (dni !== 'invalid' && emailValidation.test(email) === true && res == 106 && soporte == 'VIRTUAL') { var check = true;
    var mods = []; for (var col=13;col<45;col++) { var mod = rw[col];if (mod.substring(0,6) == 'INICIO') {mods.push(campModules[col-13] + ' - S' +sec)}}
   array.push([check,url,sec,resolucion,soporte,cuilDni,dni,dni,apellido,nombre,email,mods[0],rol,mods[1],rol,mods[2],rol,mods.length]);
   }
  //SI ALGUNA DE LAS CONDICIONES NO SE CUMPLE AGREGO LA ENTRADA SIN MODULOS
   else  { var check = false;
   array.push([check,url,sec,resolucion,soporte,cuilDni,dni,dni,apellido,nombre,email,'','','','','','','']);
   }
 } 
 
 var lastRw = sheetMoodleCSV.getLastRow();    
//Si hay estudiantes nominalizados 
 if (array.length > 0){
  sheetMoodleCSV.getRange(lastRw + 1, 1, array.length, array[0].length).setValues(array); //vuelco las entradas tratadas en el "Listado - Moodle CSVs"
  //Convierto el array de estudiantes al formato csv
      var toCsv = [];
       for (var r in array) { var arrayRw = array[r];
        if (arrayRw[0] === true) {
         var newRw = arrayRw.slice(6,17);
         toCsv.push(newRw);
        }
      } toCsv.unshift(headers);
  
   //Si hay estudiantes genero el csv y lo declaro
   if (toCsv.length > 1) {
     toCsv = multi_replace(toCsv); //saco los caracteres invalidos
     var resource = {
      title: 'Moodle CSV - S'+sec+' - '+date,
      mimeType: MimeType.CSV,
      parents: [{id: '1dR7W-jf6jFwoJ0C1_A79Mhaw1Yw0yr5s'}],
     }; 
     var csvRows = toCsv.join("\n");
     var blob = Utilities.newBlob(csvRows, "text/csv")
     blob.setContentType('application/octet-stream');   
     var created = Drive.Files.insert(resource, blob); 
    var results_row = [open.getName(),open.getUrl(),'Done']; 
   } else {var results_row = [open.getName(),open.getUrl(),'No']} //Si no hay estudiantes que exportar lo declaro y no genero el csv
 }
//Si NO hay estudiantes nominalizados  
 else { var results_row = [open.getName(),open.getUrl(),'No, planilla vacía.']}
  return results_row 
   }
catch (e) {
   var results_row = [open.getName(),open.getUrl(),e.message];
   PropertiesService.getUserProperties().setProperty('ERROR_COUNTER', parseInt(PropertiesService.getUserProperties().getProperty('ERROR_COUNTER')) + 1)
   return results_row
   }
}

//-------------------------------FUNCIONES AUXILIARES-------------------------------

//1 - SPREADSHEET DESTINO
   //genera la planilla
function generarPlanillaMoodleCSV(){                                                
  var folder = dapp.getFolderById('1t2M43wCVa1IAJf9uGxUuHoKQz1kCjMv6');
  var modelo = dapp.getFileById('1dXacqT09sjiXtP00sAjbozrwpCg6sHoWQQnsNIR-zAM');
  var copy = modelo.makeCopy('Listado - Moodle CSVs '+ date, folder);
    sheet.getRange('D11').setValue(copy.getUrl());
    openListado = sapp.openByUrl(copy.getUrl());
    sheetMoodleCSV = openListado.getSheetByName('Listado moodle');}
    //referencia a la planilla genera
function abrirPlanillaMoodleCSV(){
    openListado = sapp.openByUrl(sheet.getRange('D11').getValue());
    sheetMoodleCSV = openListado.getSheetByName('Listado moodle');}

//2 - REPLACE, para sacar los caracteres no validos en el csv
function multi_replace(data){
  var clean_data=[];
  var sust_array = ['Ñ','Á','É','Í','Ó','Ú',',',';']; //objetos a reemplazar
  var remp_array = ['N','A','E','I','O','U','',''];   //objetos de reemplazo
  for (var i in data){
   var rw=data[i];
   var joined_data=rw.join('|').toUpperCase(); //convierto a string y paso a minusculas
    for (var n in sust_array) {
     var x = new RegExp(sust_array[n], 'g');
     var newStr = joined_data.replace(x, remp_array[n]);
     joined_data = newStr;
    }
   var new_rw = newStr.split('|');
   clean_data.push(new_rw);
   }
return clean_data;
}

//3 - VACIAR CARPETA "CSVs"
function vaciarCarpetaCSV(){
 var folderMoodleCSV = dapp.getFolderById('1dR7W-jf6jFwoJ0C1_A79Mhaw1Yw0yr5s');
 var allCSV = folderMoodleCSV.getFiles();
 while (allCSV.hasNext()){
  var CSV = allCSV.next();
  CSV.setTrashed(true);
 }}
