function doGet(e){
    return  calendar();
}

function calendar(){
  var SS = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1aMSx1Mm8x1RL3xCMovXuUzNaLAkHeLlIjzzKTkWVefU/edit#gid=677037886');
  var sheet = SS.getSheetByName('CalenderSetting');
  var copylist = [];
  var copylist = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  var ID = copylist[0][2];
  //var ID = 'mara.wang@flow.tw';
  var calendar = CalendarApp.getCalendarById(ID);
  var i,j,project,output = [],year1,year2,date1,date2,event;
  var this_year = new Date().getFullYear();
  
  //Logger.log(copylist);
  var color_dict ={
    "紅": "11" ,
    "橙": "6" ,
    "黃": "5" ,
    "綠": "10" ,
    "藍": "9" ,
    "紫": "3" ,
    "灰": "8" ,
    "淺綠": "2" ,
    "粉紅": "4", 
    "青": "7" ,    
  };
  
  
  for(i=0;i<copylist.length;i++){
    try{
      if(copylist[i][0]==1){
        if(copylist[i][3]==""&&copylist[i][4]==""){
          output.push(["",""]);
          continue;
        }
        
        
        if(copylist[i][5].length > 5){//格式含年
          year1 = copylist[i][5].substring(0,3);
          date1 = copylist[i][5].substring(5);
          year2 = copylist[i][7].substring(0,3);
          date2 = copylist[i][7].substring(5);
        }else{
          year1 = this_year
          year2 = this_year;
          date1 = copylist[i][5];
          date2 = copylist[i][7];
        }
        
        
        if(copylist[i][12]==""){//沒有ID>>新活動
              event = calendar.createEvent(copylist[i][3]+"_"+copylist[i][4],
                                           new Date(date1+','+year1+' '+copylist[i][6]),
              new Date(date2+','+year2+' '+copylist[i][8]),
                {location: copylist[i][10],
                  description:copylist[i][11]}
          ).setColor(color_dict[copylist[i][9]]);
          
          output.push([event.getId() ,new Date()]);
        }else{
          event = calendar.getEventById(copylist[i][12]).setTime(new Date(date1+','+year1+' '+copylist[i][6]),new Date(date2+','+year2+' '+copylist[i][8])).setDescription(copylist[i][11]).setLocation(copylist[i][10]);
          output.push([copylist[i][12] ,new Date()]);
        }
      }else{
        output.push([copylist[i][12],copylist[i][13]]);
      }
    }catch(f){
      output.push(["",f]);
    }
  }
  sheet.getRange(2,13,output.length,2).setValues(output);
Logger.log(output);
}
