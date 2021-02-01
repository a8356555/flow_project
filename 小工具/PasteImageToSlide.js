function doGet(e){
  
   return PasteImageToSlides();

}


function PasteImageToSlides() {
  var SS = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1p1xMGoJN2EHu5kzrvAaoYJsUGvL4KSgZqVyl7F82AX4/edit#gid=491517896');
  var sheet = SS.getSheetByName('PasteImageToSlides');
  var copylist = [];
    //log for count
  var execution_list = [];
  //var SSname = SS.getName();
  //var sheetname = sheet.getName();
  var nowtime = [];
  var sheetlog;  
  //log for count  
  var deck,ppturl,ID,isgd,slide,blob,image;
  var pageWidth,pageHeight,imageW,imageH,imageWH,scale,left,top;
  var lastSlide; 

  copylist = sheet.getRange(2, 2, sheet.getLastRow()-1,1).getValues();
  ppturl = sheet.getRange(2, 3).getValues();
  deck = SlidesApp.openByUrl(ppturl[0]);

  
  for(var i=0;i<copylist.length ;i++){
    slide = deck.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      try {
    //主程式
    ID = copylist[i][0].match(/[-\w]{25,}/); 
    isgd = copylist[i][0].substring(8,20)
        if(isgd =='drive.google'){ 
          blob = DriveApp.getFileById(ID).getBlob();   
          image = slide.insertImage(blob);
        }else{
         image = slide.insertImage(copylist[i][0]);
        }
        
    pageWidth = deck.getPageWidth();
    pageHeight = deck.getPageHeight();
    imageW = image.getWidth();
    imageH = image.getHeight();
    imageWH = imageW/imageH;
   //left = pageWidth/2-imageW/2;
    //top =  pageHeight/2-imageH/2;
   

        //計算如何圖片放大及置中
    if(imageWH > 720/405 ){//比較寬
    scale = pageWidth/imageW;
    top = pageHeight/2-imageH*scale/2;
    image.setWidth(pageWidth).scaleHeight(scale).setTop(top).sendToBack();
    
    }else{//比較長(垂直)
    scale = pageHeight/imageH;
    left = pageWidth/2-imageW*scale/2;
    image.scaleWidth(scale).setHeight(pageHeight).setLeft(left).sendToBack();
    }
      
    nowtime.push([new Date()]);
  } catch (f) {
    nowtime.push([f]);
    //如果出現Exception就把新增的空白ppt刪掉
    lastSlide=deck.getSlides().pop()
    lastSlide.remove();
  }    
  /*      
  //log for count
  execution_list.push(copylist[i]);
  execution_list[execution_list.length-1].unshift(nowtime,user,SSname,sheetname);
  //log for count*/
   
  }
  sheet.getRange(2, 4,nowtime.length,1).setValues(nowtime);
  /*
    //log for count    
    var SSlog = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/173vDyRBlF6-ZmqBgOn4UaUYfj2dDszaXWwouWpa_M8w/edit#gid=735902864');
    var sheetlog =SSlog.getSheetByName('count');
    sheetlog.getRange(sheetlog.getLastRow()+1, 1, execution_list.length, execution_list[0].length).setValues(execution_list);
  //log for count */
}