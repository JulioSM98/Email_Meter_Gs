 function addDays(date, days) {
  var result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

function addHour(date,hour){
  var result = new Date(date);
  result.setHours(result.getHours()+hour);
  return result;
}


function convertDateEpoch(date){
  var result = new Date(date);
  var epoch = Math.round(result.getTime()/1000);
  return epoch;
}

function getMailDate(after,before){
  var response = Gmail.Users.Messages.list("me",{
    "maxResuts":400,
    "q":"after:"+after+" before:"+before
  });
  return response["resultSizeEstimate"]
}

function addRowDay(lastDate){
  var dateRow=addDays(lastDate,1);
  var row=[];
  row.push(dateRow);
  for(i=0;i<24;i++){
    var after=addHour(dateRow,i);
    var before=addHour(after,1);
    after=convertDateEpoch(after);
    before=convertDateEpoch(before);
    row.push(getMailDate(after,before));
  }
  return row;
}

function Actualizar_MesHours(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow= sheet.getLastRow();
  var lastCell = sheet.getRange(lastRow,1);
  var lastDate = lastCell.getValue();
  var date = new Date();
  
  if(lastDate == ''){
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert('Verifique la Fila '+lastRow+' este totalmente vacia',ui.ButtonSet.OK);
  }else{
    if(lastDate.getMonth() == date.getMonth() ){
    if(lastDate.getDate() == date.getDate()){
      sheet.deleteRow(lastRow);
      lastRow=sheet.getLastRow();
      lastCell = sheet.getRange(lastRow,1);
      lastDate = lastCell.getValue();
      sheet.appendRow(addRowDay(lastDate));
    }
  }else{
      while(lastDate.getDate() != date.getDate()){
        sheet.appendRow(addRowDay(lastDate));
        lastRow=sheet.getLastRow();
        lastCell = sheet.getRange(lastRow,1);
        lastDate = lastCell.getValue();
      }
    }
  }
  
}

function onOpen(){
  var uis=SpreadsheetApp.getUi();
  uis.createMenu('Email Meter')
  .addItem('Actualizar mes por hora', 'Actualizar_MesHours' )
  .addToUi();
}


