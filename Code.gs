/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current spreadsheet.
 */

function onInstall(e) {
    Logger.log("OnINSTALL");
  onOpen(e);
}

function onEdit(e){
  // Set a comment on the edited cell to indicate when it was changed.
//  var range = e.range;
//  updateStatusColor(range);
  Logger.log("OnEDIT");
  updateStatusColor();
  updatePercentage();
}
function onOpen(e) {
  Logger.log("OnOPEN");
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Sincronizar Calendario', 'calendarSync')
      .addItem('Preparación del Documento PCM', 'showSidebar')
      .addToUi();
}
//---------------------------------------------------------------------------------------------------------------------------------------------------------

/**
 * Runs when the add-on is installed; Prepares format, locale and data examples.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function preparePcmSheets(e){
  
  var CalendarColumns = 12;
  var CalendarRows = 20;
  
  var CtrlColumns = 5;
  var CtrlRows = 10;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetLocale('es_MX');
  var sheet = SpreadsheetApp.getActiveSheet();
  SpreadsheetApp.getActiveSpreadsheet().renameActiveSheet("<Nombre_del_Proceso>");

  var MaxColumns = parseInt(sheet.getMaxColumns());
  var MaxRows = parseInt(sheet.getMaxRows());
  sheet.deleteColumns(CalendarColumns, MaxColumns-CalendarColumns);
  sheet.deleteRows(CalendarRows, MaxRows-CalendarRows);
  
  sheet.getRange('A1').setValue('Calendar_id');
  sheet.getRange('B1').setValue('Landscape');
  sheet.getRange('C1').setValue('Env');
  sheet.getRange('D1').setValue('S. Type');
  sheet.getRange('E1').setValue('SID');
  sheet.getRange('F1').setValue('Ticket Novis');
  sheet.getRange('G1').setValue('Ticket EPH');
  sheet.getRange('H1').setValue('Status');
  sheet.getRange('I1').setValue('Inicio');
  sheet.getRange('J1').setValue('Fin');
  sheet.getRange('K1').setValue('Ejecutor');  
  sheet.getRange('L1').setValue('Observaciones');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
 
  var statusCol = sheet.getRange("H:H");  //H is the Status Column
  var statusRule = SpreadsheetApp.newDataValidation().requireValueInList(["EJECUTADO","PROGRAMADO","EN\ EJECUCION","SUSPENDIDO","CANCELADO"],true).setAllowInvalid(false).build();
  statusCol.setDataValidation(statusRule);  
  var datetimeCol = sheet.getRange("I:J");  //I:J are the Status Column
  var datetimeRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  datetimeCol.setDataValidation(datetimeRule);
  datetimeCol.setNumberFormat("dddd dd/mm/yyyy HH:mm");
  datetimeCol.setHorizontalAlignment("right");
 


  sheet=ss.insertSheet('CTRL', {template: false});
  var MaxColumns = parseInt(sheet.getMaxColumns());
  var MaxRows = parseInt(sheet.getMaxRows());
  sheet.deleteColumns(CtrlColumns, MaxColumns-CtrlColumns);
  sheet.deleteRows(CtrlRows, MaxRows-CtrlRows);

  sheet.getRange('A1').setValue('Nombre de\nCalendario:');
  ss.setNamedRange("CAL_NAME", sheet.getRange('B1'));
  sheet.getRange("A1").setBackground('#CCCCCC');
  sheet.getRange("B1").setBackground('#EEEEEE');
  sheet.getRange("A1:B1").setBorder(true, true, true, true, true, true);

  sheet.getRange('B3').setValue('Porcentaje');
  sheet.getRange('A4').setValue('Programado');
  ss.setNamedRange("PLAN", sheet.getRange('B4'));
  sheet.getRange('A5').setValue('Real');
  ss.setNamedRange("REAL", sheet.getRange('B5'));
  sheet.getRange("A3:B5").setBackground('#CCCCCC');
  sheet.getRange("B4:B5").setBackground('#EEEEEE');
  sheet.getRange("A3:B5").setBorder(true, true, true, true, true, true);

  var id = SpreadsheetApp.getActiveSpreadsheet().getId();
  var file = DriveApp.getFileById(id);
  file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.EDIT);
  
}
//---------------------------------------------------------------------------------------------------------------------------------------------------------
function calendarSync(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var startRow = 2;
  var numRows = sheet.getLastRow();
//  var numRows = 1;
  var dataRange = sheet.getRange(startRow, 1, numRows, 12);
  var data = dataRange.getValues();
  
  var calendar_name = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("CAL_NAME").getValue();
//var calendar_color = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("CAL_COLOR").getBackground();

  var cal = CalendarApp.getCalendarsByName(calendar_name)[0];
  Logger.log(cal);
  if(!cal){
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('El Calendario (\" '+calendar_name+' \") que has indicado en la pestaña CTRL no es válido.');
    throw new Error("El usuario está wn(CAL_NAME)");
  }

  var title = sheet.getName(); 
  Logger.log(title);
  
  var id = SpreadsheetApp.getActiveSpreadsheet().getId();
  var file = DriveApp.getFileById(id);
  var url = file.getUrl();
  
  for (i in data) {
    var row = data[i];
    
    var calendar_id = row[0];
    var landscape   = row[1];
    var environment = row[2];
    var s_type      = row[3];
    var sid         = row[4];
    var tkt_novis   = row[5];
    var tkt_eph     = row[6];
    var status      = row[7];
    var tstart      = new Date(row[8]);
    var tstop       = new Date(row[9]);
    var ejecutor    = row[10];
    var obs         = row[11];
    var j           = parseInt(i)+startRow;
    
    var desc        =     "Ticket Novis\t:"+tkt_novis+"\n";
    var desc        =desc+"Ticket EPH\t:"+tkt_eph+"\n";
    var desc        =desc+"\n";
    var desc        =desc+"Landscape: "+landscape+"\n";
    var desc        =desc+"Ambiente\t: "+environment+"\n";
    var desc        =desc+"S. Type\t: "+s_type+"\n";
    var desc        =desc+"SID\t\t: "+sid+"\n";
    var desc        =desc+"\n";
    var desc        =desc+"Ejecutor     : "+ejecutor+"\n";
    var desc        =desc+"Observaciones: "+obs+"\n";
    var desc        =desc+"\n";
    var desc        =desc+"URL     : "+url+"\n";
    
    var event_name  =sid+" - "+title+" - "+status;
    Logger.log(calendar_id+" "+sid+" "+tstart+" "+tstop);
    
    if(tstart.getTime() && tstop.getTime() ){
      //Given there is no eventUpdate method (only for eventSeries
      //instead of updating the event, the process is deleting the current and creating a new one
      if(calendar_id){
        var event = cal.getEventSeriesById(calendar_id);
        event.deleteEventSeries();
        sheet.getRange('A'+j).setValue('');
      }
      var event = cal.createEvent(event_name, tstart, tstop, {description:desc});
      event.addPopupReminder(120);      
      //update row with event_id
      sheet.getRange('A'+j).setValue(event.getId());
      Logger.log(event.getId());
    }else{
      if(calendar_id){
        var event = cal.getEventSeriesById(calendar_id);
        event.deleteEventSeries();
        sheet.getRange('A'+j).setValue('');
      }
    }
  }
}


//---------------------------------------------------------------------------------------------------------------------------------------------------------
/**
 * Opens a sidebar. The sidebar structure is described in the Sidebar.html
 * project file.
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate()
      .setTitle("Preparación de Hoja de Cálculo");
  SpreadsheetApp.getUi().showSidebar(ui);
}


//---------------------------------------------------------------------------------------------------------------------------------------------------------

function updateStatusColor(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var numRows = sheet.getLastRow();
  for(i=1;i<numRows;i++){
    var cell = sheet.getRange("H"+i);  //H is the Status Column
//    var cell=range;
    
    switch(cell.getValue()){
      case "EJECUTADO":
        cell.setFontColor('#007D1E');
        cell.setBackground('#D3EDD3');
        break;
      case "PROGRAMADO":
        cell.setFontColor('#004E96');
        cell.setBackground('#CBE2F4');
        break;

      case "EN\ EJECUCION":
        cell.setFontColor('#C69100');
        cell.setBackground('#FFF4CA');
        break;

      case "SUSPENDIDO":
        cell.setFontColor('#B7B7B7');
        cell.setBackground('#EEEEEE');
        break;

      case "CANCELADO":
        cell.setFontColor('#A90000');
        cell.setBackground('#FCC9CA');
        break;
      default:
        cell.setFontColor('#000');
        cell.setBackground('#FFF');
    }
  }
}

//---------------------------------------------------------------------------------------------------------------------------------------------------------
function updatePercentage(){
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var firstSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var numRows = firstSheet.getLastRow();
  var total=0;
  var plan=0;
  var real=0;
  var i=0;
  for(i=1;i<numRows;i++){
    var cell = firstSheet.getRange("H"+i);  //H is the Status Column
    switch(cell.getValue()){
      case "EJECUTADO":
        total=total+1;
        plan=plan+1;
        real=real+1;
        break;
      case "PROGRAMADO":
        total=total+1;
        break;
      case "EN\ EJECUCION":
        total=total+1;
        plan=plan+1;
        real=real+1;
        break;
      case "SUSPENDIDO":
        total=total-1;
        break;
      case "CANCELADO":
        total=total+1;
        plan=plan+0;
        real=real-1;
        break;
      default:
        total=total+1;
    }
  }
  var planPorcentaje = (plan)/total;
  var realPorcentaje = (real)/total;

  Logger.log(planPorcentaje+" v/s "+realPorcentaje);
  cell.setNumberFormat("#%");
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("PLAN").setNumberFormat("##.##%").setValue(planPorcentaje);
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("REAL").setNumberFormat("##.##%").setValue(realPorcentaje);
}
