function cerrarPlanillas(file){ 
try {
    var fileId = file.getId();
    var open = SpreadsheetApp.openById(fileId);
// CIERRO PLANILLA DE ESTUDIANTES:
    var estudiantes = open.getSheetByName('ESTUDIANTES');
       
         estudiantes.protect().remove();
          var estudiantesProtection = estudiantes.protect().setDescription('Planilla cerrada');
          
             var me = Session.getEffectiveUser();
              estudiantesProtection.addEditor(me);
              estudiantesProtection.addEditor('modalidadadistancia.adultos@gmail.com');
                 var editors = estudiantesProtection.getEditors();
                  for(var e in editors){
                    if(editors[e].getEmail() != 'modalidadadistancia.adultos@gmail.com'){
                      estudiantesProtection.removeEditor(editors[e]);
                    }
                  }
                if (estudiantesProtection.canDomainEdit()){
                 estudiantesProtection.setDomainEdit(false);
                }
                
// ABRO PLANILLA DE "CAMBIOS FUERA DE PLAZO":
       var cambios = open.getSheetByName('Cambios fuera de plazo');
       
         var cambiosProtection = cambios.protect().setDescription('Planilla abierta');
           var rangeDatos = cambios.getRange('E5:AW55');
          
           var unprotected = cambiosProtection.getUnprotectedRanges();
             unprotected.push(rangeDatos);
             cambiosProtection.setUnprotectedRanges(unprotected);
             
             var me = Session.getEffectiveUser();
              cambiosProtection.addEditor(me);
               cambiosProtection.removeEditors(cambiosProtection.getEditors());
                if (cambiosProtection.canDomainEdit()){
                 cambiosProtection.setDomainEdit(false);
                }
     var results_row = [open.getName(),open.getUrl(),'Done'];
     return results_row            
     }
catch (e) {
     var results_row = [open.getName(),open.getUrl(),e.message];
     PropertiesService.getUserProperties().setProperty('ERROR_COUNTER', parseInt(PropertiesService.getUserProperties().getProperty('ERROR_COUNTER')) + 1)
     return results_row
     }
}     
