function doGet(e){
    return  remove();
}


function remove(){
  var SS = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/10mxT3Tl0fuA3VM3lkrcW4UOmhLMezhPculemrLiP4Ng/edit#gid=0');
  var sheet = SS.getSheetByName('RemoveCalenderEvent');
    var copylist = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()-1).getValues();
  var ID = copylist[0][2];
  var calendar = CalendarApp.getCalendarById(ID);
  var i,j,project,output = [],year1,year2,date1,date2,event;
  var this_year = new Date().getFullYear()

  //Logger.log(copylist[0][5])
  for(i=0;i<copylist.length;i++){
    try{
      if(copylist[i][0]==1){
        
        if(copylist[i][4].length > 5){//格式含年
          year1 = copylist[i][4].substring(0,3);
          date1 = copylist[i][4].substring(5);
          year2 = copylist[i][6].substring(0,3);
          date2 = copylist[i][6].substring(5);
        }else{
          year1 = this_year
          year2 = this_year;
          date1 = copylist[i][4];
          date2 = copylist[i][6];
        }
        
         event = calendar.getEvents(new Date(date1+','+year1+' '+copylist[i][5]+' GMT+08:00'),
                                    new Date(date2+','+year2+' '+copylist[i][7]+' GMT+08:00'),
                                  {search: copylist[i][3]}
                                   );
        for(j=0;j<event.length;j++){
         event[j].deleteEvent();
        }
        Logger.log(date1+','+year1+' '+copylist[i][5],date2+','+year2+' '+copylist[i][7])
          
          output.push([new Date()]);
       
      }else{
        output.push([copylist[i][8]]);
      }
    }catch(f){
      output.push([f])
    }
  }
  sheet.getRange(2,9,output.length,1).setValues(output);
//Logger.log('Event ID: ' + event.getId());
  
}



