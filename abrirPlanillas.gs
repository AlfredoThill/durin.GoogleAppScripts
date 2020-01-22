function abrirPlanillas(file){
try {
    var fileId = file.getId();
    var open = SpreadsheetApp.openById(fileId);  
// CIERRO PLANILLA DE "CAMBIOS FUERA DE PLAZO": 
       var cambios = open.getSheetByName('Cambios fuera de plazo');
        
        var destinationsheet = cambios;
        var remove_protection = destinationsheet.protect().remove();
        var destination_protection = destinationsheet.protect();
         destination_protection.setDescription('Planilla Cerrada');
        var removeeditor = destination_protection.removeEditors(destination_protection.getEditors());
                
// ABRO PLANILLA DE ESTUDIANTES: 
    var estudiantes = open.getSheetByName('ESTUDIANTES');    
       
        var destinationsheet = estudiantes;
        var ranges_to_unprotect = destinationsheet.getRangeList(['B7','B9','B11','B13','B17','E6:AV500']).getRanges();
        var destination_protection = destinationsheet.protect();
        destination_protection.setDescription('Planilla Abierta');
        var removeeditor = destination_protection.removeEditors(destination_protection.getEditors());
        var update_protection = destination_protection.setUnprotectedRanges(ranges_to_unprotect);
       
     var results_row = [open.getName(),open.getUrl(),'Done'];
     return results_row             
     }
catch (e) {
     var results_row = [open.getName(),open.getUrl(),e.message];
     PropertiesService.getUserProperties().setProperty('ERROR_COUNTER', parseInt(PropertiesService.getUserProperties().getProperty('ERROR_COUNTER')) + 1)
     return results_row
     }
}
