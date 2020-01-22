var openInforme,modulosSheet,informeSheet;

function informeDeCambios(file){
//-------------------------------RUTINA PRINCIPAL-------------------------------
try {
    var fileId = file.getId();
    var open = SpreadsheetApp.openById(fileId);
//SETTINGS:
var estSheet,estValues;
    estSheet = open.getSheetByName('ESTUDIANTES');
    estValues = estSheet.getRange('G6:G').getValues();  
var camSheet, camValues;
    camSheet = open.getSheetByName('Cambios fuera de plazo');
    camValues = camSheet.getRange('A5:AW55').getValues();
var cant = camSheet.getRange('B2').getValue();
var headers = camSheet.getRange('A3:AW3').getValues();
var sec = open.getSheetByName('ESTUDIANTES').getRange('B15').getValue();   

  var estC = [];  // Array de filas con Cambios.
  var rows = [];  // Indices para borrar filas con cambios.
  var infoPersonal = []; // Informe.
  var infoModulos = [];  // Cambios en los módulos (Bajas e Inicios).

        if(cant > 0){
          cambios(estC,rows,camValues); // 1. Genero Array con cambios.
        }
        if(estC.length > 0){ // Si hay cambios...
          informePersonal(estC,campModules,sec,infoPersonal,infoModulos); // 2. Genero informes.
          modEstudiantes(estC,rows,estValues,estSheet,camSheet,open); // 3. Borro las planillas y aplico cambios.
        }
     var results_row = [open.getName(),open.getUrl(),'Done'];
     return results_row   
     }
catch (e){
     var results_row = [open.getName(),open.getUrl(),e.message];
     PropertiesService.getUserProperties().setProperty('ERROR_COUNTER', parseInt(PropertiesService.getUserProperties().getProperty('ERROR_COUNTER')) + 1)
     return results_row
     }
}

//-------------------------------FUNCIONES AUXILIARES-------------------------------

//1 - GENERO ARRAY DE CAMBIOS.
function cambios(estC,rows, camValues){  // Genero Array con cambios y Genero Array con índices.
  for(var i in camValues){
    var filaCambio = camValues[i];  // Fila de cambios.
    if(filaCambio[1] > 0){  // Si realizaron algún cambio.
      estC.push(filaCambio);  // Armo el array con filas con cambios.
      rows.push(+i+5);
    }
  }
  return estC,rows;
};

//2 - GENERO INFORMES DE CAMBIOS DE INFO. PERSONAL Y MÓDULOS.
function informePersonal(estC,campModules,sec,infoPersonal,infoModulos){
  for(var i in estC){  // Itero sobre el array con Cambios. Los reparto en distintos informes (Virtual o Módulos)
    var cambio = estC[i];  // Fila de Cambios.
      if(cambio[0] == 'VIRTUAL'){ // Genero informe si son de VIRTUAL.
        if(cambio[5] != '' || cambio[6] != '' || cambio[7] != '' || cambio[10] != '' || cambio[16] != ''){
          infoPersonal.push([date,sec,cambio[4],cambio[5],cambio[6],cambio[7],cambio[10], cambio[16]]);
        }
       for(var x=17;x<cambio.length;x++){
         var modulo = campModules[x-17] + ' - S'+sec;  // Reemplazo por nombre de campus y concateno sección.
           if(cambio[x] != '' && cambio[x] != 'BAJA' && cambio[x][0] != 'A'){  // Si es un inicio.
            infoModulos.push([date, sec, cambio[4], 'INICIO',cambio[x], modulo]);
           } else if(cambio[x] == 'BAJA'){  // Si es Baja.
              infoModulos.push([date, sec, cambio[4], 'BAJA',cambio[x], modulo]);
             }
             else if(cambio[x][0] == 'A'){  // Si es (A)probado.
              infoModulos.push([date, sec, cambio[4], 'APROBADO',cambio[x], modulo]);
             }
       }
      }
  }
  if(infoPersonal.length > 0 || infoModulos.length > 0){
     
      if(infoPersonal.length > 0){
        var lastRi = informeSheet.getLastRow() + + 1;
         informeSheet.getRange(lastRi, 1, infoPersonal.length, infoPersonal[0].length).setValues(infoPersonal);
      }
      if(infoModulos.length > 0){
        var lastRm = modulosSheet.getLastRow() + + 1;
         modulosSheet.getRange(lastRm, 1, infoModulos.length, infoModulos[0].length).setValues(infoModulos);
      }
  }
};

//3 - APLICO LOS CAMBIOS EN "ESTUDIANTES" Y LIMPIO LA PLANILLA DE "CAMBIOS".
function modEstudiantes(estC,rows,estValues,estSheet,camSheet,open){ 
  for(var i in estC){  // Itero sobre el array con Cambios. Los reparto en distintos informes (Virtual o Módulos)
    var cambio = estC[i];  // Fila de Cambios.
      for(var j in estValues){
        var dniE = estValues[j];
          if(dniE == cambio[4]){  //Busco los índices de los DNI de estudiantes.
            var row = +j+6;
              for(var k=5;k<cambio.length;k++){
                if(cambio[k] != '' && cambio[k] != 'BAJA'){
                  estSheet.getRange(row, k).setValue(cambio[k]);
              } else if (cambio[k] == 'BAJA'){
                estSheet.getRange(row, k).clearContent();
                }
              }
          }
      }
      camSheet.getRange(rows[i], 5, 1, 47).clearContent();  // Borro los cambios plasmados en 'CAMBIOS'.
  }
  //refrescar la validacion y el formato de la columna dni
  var dniRng = camSheet.getRange('E5:E55');
  dniRng.setDataValidation(SpreadsheetApp.newDataValidation().setAllowInvalid(true).setHelpText('Introduzca un "Cuil/DNI" tal cual se encuentra en la solapa "ESTUDIANTES".').requireValueInRange(open.getSheetByName('Resumen').getRange('$D$9:$D$1005'), true).build());
  dniRng.setNumberFormat('@');
  camSheet.getRange('G1').setValue('Los cambios se han plasmado en la planilla de "ESTUDIANTES" el día : '+date);
}

//4 - SPREADSHEET DESTINO
   //genera la planilla
function generarPlanillaInformeDeCambios(){                                                
  var folder = dapp.getFolderById('1GMsWZh075LPoyCi08Unz3fisqPUCuT5U');
  var modelo = dapp.getFileById('1KMEAAz_5kq4bToGywXoNP8FxaVODBHS2KatS27zaTjw');
  var copy = modelo.makeCopy('Informe, Cambios '+ date, folder);
    sheet.getRange('D9').setValue(copy.getUrl());
    openInforme = sapp.openByUrl(copy.getUrl());
    modulosSheet = openInforme.getSheetByName('1. Cambios módulos');
    informeSheet = openInforme.getSheetByName('2. Cambios personales');
    dapp.getFileById(copy.getId()).addViewer('modalidadadistancia.adultos@gmail.com');}
    //referencia a la planilla genera
function abrirPlanillaInformeDeCambios(){
    openInforme = sapp.openByUrl(sheet.getRange('D9').getValue());
    modulosSheet = openInforme.getSheetByName('1. Cambios módulos');
    informeSheet = openInforme.getSheetByName('2. Cambios personales');}
 //-------------------------------------------------------------------
