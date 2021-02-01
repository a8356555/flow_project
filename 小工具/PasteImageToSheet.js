function doGet(e){
  
   return PasteImageToSheet();

}


function PasteImageToSheet() {
  var SS = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1p1xMGoJN2EHu5kzrvAaoYJsUGvL4KSgZqVyl7F82AX4/edit#gid=491517896');
  var sheet = SS.getSheetByName('PasteImageToSheet');
  var copylist = [];
  var nowtime = [];

  copylist = sheet.getRange(2, 2, sheet.getLastRow()-1,5).getValues();
  
  var SS2 = SpreadsheetApp.openByUrl(copylist[0][1]);
  var sheet2 = SS2.getSheetByName(copylist[0][2]); 
  var images,image,blob,ID,isgd,row,col,calcu,img_w,img_h;
  var i,j;
  //Logger.log(copylist[1][2])

  for(i=0;i<copylist.length ;i++){

      try {
    //主程式
        calcu = copylist[i][3].split("");
        if(parseInt(calcu[1])>=0){//A1notation轉數字
          col = calcu[0].toLowerCase().charCodeAt(0) - 96;
          row = parseInt(copylist[i][3].substring(1));
        }else{
          col = (calcu[0].toLowerCase().charCodeAt(0) - 96)*26 + (calcu[1].toLowerCase().charCodeAt(0) - 96);
          row = parseInt(copylist[i][3].substring(2));
        }
       
        //插入圖片
        ID = copylist[i][0].match(/[-\w]{25,}/); 
        isgd = copylist[i][0].substring(8,20);
        if(isgd =='drive.google'){ 
          blob = DriveApp.getFileById(ID).getBlob();   
          image =  sheet2.insertImage(blob, col,row);
        }else{
         image =   sheet2.insertImage(copylist[i][0],col,row);
        }
     
        //調整圖片大小&儲存格大小
        
        
        nowtime.push([new Date()]);    
      } catch (f) {
        nowtime.push([f]);
      }    
  }
        images = sheet2.getImages();
        for(j=0;j<images.length;j++){
          img_w = images[j].getWidth();
          img_h = images[j].getHeight();     
          if(img_w/img_h>1){
            images[j].setWidth(405).setHeight(img_h*405/img_w).setAnchorCellYOffset((405-img_h*405/img_w)/2);
          }else if(img_w/img_h==1){
            images[j].setWidth(405).setHeight(405);
          }else{
            images[j].setWidth(img_w*405/img_h).setHeight(405).setAnchorCellXOffset((405-img_w*405/img_h)/2);
          }
        }
   
    sheet2.setColumnWidths(1, sheet2.getMaxColumns(), 410).setRowHeights(1,sheet2.getMaxRows(),410);
    sheet.getRange(2, 6,nowtime.length,1).setValues(nowtime);
  
}