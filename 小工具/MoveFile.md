function doGet(e){
  return move_file();
}
function move_file(){
  var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1uGpDL87zbgysBAF9Fk2TYKacpGFpCneepoUkzKfKhHY/edit#gid=1010445339');
  var sheet = ss.getSheetByName('移動檔案');
  var copylist = sheet.getDataRange().getValues();
  var a,b,c,d,i,j;
  var sheet1,output = [];
  var files,file,fileId,folderId,original_folderId;
  
  output = sheet.getRange(4,6,sheet.getLastRow(),2).getValues();

    for(i=4;i<copylist.length;i++){
      if(copylist[i][0]==1){
        try{
          original_folderId = copylist[i][1].match(/[-\w]{25,}/);
          folderId = copylist[i][3].match(/[-\w]{25,}/);

          if(copylist[i][2]!=""){    
            fileId = copylist[i][2].match(/[-\w]{25,}/);
            file = DriveApp.getFileById(fileId);
            DriveApp.getFolderById(folderId).addFile(file);     
            if(copylist[i][4]=="是"){
              DriveApp.getFolderById(original_folderId).removeFile(file);
            }
            
            
          }else{
            files = DriveApp.getFolderById(original_folderId).getFiles();
            while(files.hasNext()){
              fileId = files.next().getId();
              file = DriveApp.getFileById(fileId);
              DriveApp.getFolderById(folderId).addFile(file);
              if(copylist[i][4]=="是"){
                DriveApp.getFolderById(original_folderId).removeFile(file);
              }
            }
            
            
          
          
          }
          
          
          
        output[i-3]=[,new Date()];
        }catch(f){
        output[i-3]=[f,new Date()]
        }
      }
    }
  sheet.getRange(4,6,output.length,2).setValues(output);
}
