function doGet(e){

   return getURL();

}


function getURL() {
  var url = 'https://docs.google.com/spreadsheets/d/1uGpDL87zbgysBAF9Fk2TYKacpGFpCneepoUkzKfKhHY/edit#gid=1010445339'
  var name = '獲得檔案URL'
  var SS = SpreadsheetApp.openByUrl(url);
  var sheet = SS.getSheetByName(name);
  var sheet2 = SS.getSheetByName("結果_獲得檔案URL");
  var copylist = sheet.getDataRange().getValues();
  var folderurl,folderID,i,folder,file,name,mime,files,url_n,nowtime=[];
  var nowtime = sheet2.getRange(1,1,sheet2.getLastRow(),4).getValues();
  var output = sheet.getRange(4,3,sheet.getLastRow(),1).getValues();
  
  for(i=4;i<copylist.length;i++){
    if(copylist[i][0]==1){
      try{
        folderurl = copylist[i][1];
        folderID = folderurl.match(/[-\w]{25,}/);
        folder = DriveApp.getFolderById(folderID);
        files = folder.getFiles();
        
        while (files.hasNext()) {
          file = files.next();
          url_n = file.getUrl();
          name = file.getName();
          nowtime.push([folderurl,name,url_n,new Date()]);
        }
      sheet.getRange(i+1,3).setValue(new Date());
      }catch(f){
      sheet.getRange(i+1,3).setValue(f);
      }
    }
  }
  sheet2.getRange(sheet2.getLastRow(),1,nowtime.length,4).setValues(nowtime);
}
  