function myFunction2222() {
   var baseFolder = DriveApp.getFolderById('google drive folder id');
   var ite = baseFolder.getFiles();
  
  var folder_temp,name,key,url,owner,time,blob,Ctype,Mime,file;
    var temp_final_array_folder=[];
    //撈資料夾圖片資料
    while (ite.hasNext()){
      folder_temp = ite.next();
      //依序為檔案ID、Blob?、內容屬性、Mime屬性(同左)、檔名、建立時間、url、擁有者
      key = folder_temp.getId();
      try{
      blob = DriveApp.getFileById(key).getBlob();
      
      Ctype = blob.getContentType();
      //if(Ctype=='image/png')
      }catch(e){}
      Mime = DriveApp.getFileById(key).getMimeType()
      
      file = DriveApp.getFileById(key)
       if(Mime.substring(0,5) == 'image'){
       file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        Logger.log("更改權限")
       }
      
      name = folder_temp.getName();
      time = folder_temp.getDateCreated();
      url = folder_temp.getUrl();
      owner = folder_temp.getOwner().getName();
      
      
      temp_final_array_folder.push([name,url,key,owner,blob,Ctype,Mime]);
    }
    var P = SpreadsheetApp.openByUrl('upload sheet url');
  var upload_sheet = P.getSheetByName('工作表1');
  upload_sheet.getRange(1, 1,temp_final_array_folder.length,7).setValues(temp_final_array_folder);
}
#
