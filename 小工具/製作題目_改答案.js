function main(){
  pasteImageToSlides()
  getUrl()
  copySlides()
  getAnswerMain()
  checkAnswerMain()
}

function pasteImageToSlides() {
  //設定試算表網址
  //指定試算表名
  var url = 'https://docs.google.com/spreadsheets/d/1pnVkrM3s9l8SLWk2jYCGpj2SdsURagl50LdlCgHYQi4/edit#gid=1515001339'
  var name = '上傳圖片(設定到背景)'
  var SS = SpreadsheetApp.openByUrl(url);
  var sheet = SS.getSheetByName(name);
  var copylist = [];
  
  var execution_list = [];
  var nowtime = [];
  var sheetlog; 
  
  var deck,ppturl,ID,isgd,slide,blob,image;
  var pageWidth,pageHeight,imageW,imageH,imageWH,scale,left,top;
  var lastSlide,slides, page_background; 

  copylist = sheet.getRange(2, 2, sheet.getLastRow()-1,2).getValues();

  deck = SlidesApp.openByUrl(copylist[0][0]);

  
  for(var i=0;i<copylist.length ;i++){
    slide = deck.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    //slides = deck.getSlides()[i+1];
    page_background = slide.getBackground();

      try {
    //主程式
    ID = copylist[i][1].match(/[-\w]{25,}/); 
    isgd = copylist[i][1].substring(8,20);
        if(isgd =='drive.google'){ 
          blob = DriveApp.getFileById(ID).getBlob(); 
          page_background.setPictureFill(blob);
        }else{
         page_background.setPictureFill(copylist[i][1]);
        }
    nowtime.push([new Date()]);
  } catch (f) {
    nowtime.push([f]);
    //如果出現Exception就把新增的空白ppt刪掉
    lastSlide=deck.getSlides().pop()
    lastSlide.remove();
  }    
   
  }
    sheet.getRange(2, 4,nowtime.length,1).setValues(nowtime);
}

function getUrl() {
  var url = 'https://docs.google.com/spreadsheets/d/1roi6w3nfbOtqK44tdKnuIWqI2iPdSJN7zjqdK0CqmlE/edit?usp=drive_web&ouid=106203360429558678203'
  var name = '獲得圖片url'
  var SS = SpreadsheetApp.openByUrl(url);
  var sheet = SS.getSheetByName(name);
  var folderurl = sheet.getRange(2, 2).getValue();
  var UrlList = [], nowtime = [];
  var url_n;
  var folderID = folderurl.match(/[-\w]{25,}/);
  var folder = DriveApp.getFolderById(folderID);
  var file,name,mime;
  
  var files = folder.getFiles();
  
  while (files.hasNext()) {
    try{
    file = files.next();
    url_n = file.getUrl();
    mime = file.getMimeType();
    name = file.getName();
    if(mime.substring(0,5) == 'image'){
    nowtime.push([name,url_n,new Date()]);  
    }
    }catch(f){
    nowtime.push([name,url_n,f]); 
    }
  }
    sheet.getRange(2,3,nowtime.length,3).setValues(nowtime);
}


function copySlides() {
  //設定試算表網址
  //指定試算表名
  var url = 'https://docs.google.com/spreadsheets/d/1roi6w3nfbOtqK44tdKnuIWqI2iPdSJN7zjqdK0CqmlE/edit?usp=drive_web&ouid=106203360429558678203'
  var name = '複製題目檔'
  var SS = SpreadsheetApp.openByUrl(url);
  var sheet = SS.getSheetByName(name);
  var copylist = [];
  var List = [];
    //log for count
  var execution_list = [];
  var nowtime = [];
  var sheetlog; 
  //log for count  
  
  var Pre_Id, folder_Id,newFile,newPre,url;

  copylist = sheet.getRange(2, 1, sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  
  for(var i=0;i<copylist.length ;i++){
    
    if(copylist[i][0]==1){
      try {
    //主程式
  folder_Id = copylist[i][2].match(/[-\w]{25,}/)
  Pre_Id = copylist[i][1].match(/[-\w]{25,}/); 

  newFile = {
   title: copylist[i][3],
  parents: [{id: folder_Id}]
  };
        
  newPre = Drive.Files.copy(newFile, Pre_Id);
  url = newPre.alternateLink;
        
   nowtime.push([url,new Date()]);
  } catch (f) {
   nowtime.push([url,f]);
  }    
    }else{
    nowtime.push([copylist[i][4],copylist[i][5]]);
    }

  }
      sheet.getRange(2, 5,nowtime.length,2).setValues(nowtime);      

}





function GetAnswer(shape,writer,ID)
{
  var num = parseInt(shape.getText().asString());
  return [writer,num,shape.getLeft(),shape.getTop(),shape.getWidth(),shape.getHeight(),shape.getObjectId(),shape.getShapeType(),ID];
  }


function GetSlideAndShape(url,writer)
{
   var deck, slides, slide, j, k, shape, ID;
   var push = [], cache = [];
   deck = SlidesApp.openByUrl(url);
   slides = deck.getSlides();
   ID = deck.getId();

  
  for(j=1;j<slides.length;j++){ 
    slide = slides[j];
    //Ele = slide.getPageElements();      
    shape =  slide.getShapes() ;      
    for (k = 0; k < shape.length; k++) {
      push = GetAnswer(shape[k],writer,ID)
      cache[push[1]]=push;
    }
       //0誰、1題號、2左座標、3上座標、4寬度、5高度、6shape_ID、7shape_type、8pre_ID
   }  
  return cache;
  }

function Judge(answer,correct,output,writer)
{
  var judge, judge2,j; 
  for(j=1;j<correct.length;j++){
      if(answer[j]!=null){//判斷有沒有作答
          judge = Math.abs(answer[j][2]-correct[j][2])+Math.abs(answer[j][3]-correct[j][3]);
          judge2 = Math.abs(answer[j][2]+answer[j][4]-correct[j][2]-correct[j][4])+Math.abs(answer[j][3]+answer[j][5]-correct[j][3]-correct[j][5]);
          //Logger.log(j,'題','作答左',answer[j][2],'答案左',correct[j][2],'作答右',answer[j][3],'答案右',correct[j][3])    
          if(answer[j][7]=="ELLIPSE"){
            //誰、第幾題、題型、判斷1、判斷2、pre_ID、shapeID、
            output.push([answer[j][0],j,"關鍵點",judge,"",answer[j][8],answer[j][6],""]); 
          }      
          else if(answer[j][7]=="RECTANGLE"){
            output.push([answer[j][0],j,"拉框",judge,judge2,answer[j][8],answer[j][6],""]); 
          }else{
          }   
      }else{
      output.push([writer,j,"未作答","未作答","未作答","未作答","未作答","",""]);
      }
  }
  return output;
}



function getAnswerMain() {
  var url = 'https://docs.google.com/spreadsheets/d/1roi6w3nfbOtqK44tdKnuIWqI2iPdSJN7zjqdK0CqmlE/edit?usp=drive_web&ouid=106203360429558678203';
  var name = '取得答案';
  var sheet2,name2,SS2,l;
  var SS = SpreadsheetApp.openByUrl(url);
  var sheet = SS.getSheetByName(name);
  var copylist = [],answer=[],correct=[],output = [],nowtime = [];
  
  copylist = sheet.getRange(2, 2, sheet.getLastRow()-1,5).getValues();
  //////////////抓答案  
  correct = GetSlideAndShape(copylist[0][0],"Answer")
  output[0]=["作答者","題號","題型","判斷1","判斷2","presentation_ID","ShapeID","=\"對/錯\"&CHAR(10)&\"請填入True/False\""]
  //Logger.log(correct);
  
  for(l=0;l<copylist.length;l++){
    try{
      ////////////////抓作答
      answer = [];
      answer = GetSlideAndShape(copylist[l][3],copylist[l][2])
      //Logger.log(answer);
      ///////////////判定對錯
      Judge(answer,correct,output,copylist[l][2])

  //Logger.log(output);
  nowtime.push([new Date()]);
  } catch (f) {
   nowtime.push([f]);
  } 
    }
    SS2 = SpreadsheetApp.openByUrl(copylist[0][1]);
    sheet2 = SS2.getSheetByName('_Record');
  
  if(sheet2=="Sheet"){//如果試算表已存在
      sheet2.getRange(SS2.getLastRow()+1,1,output.length,output[0].length).setValues(output);
  }else{
    SS2.insertSheet(0).setName('_Record').getRange(1,1,output.length,output[0].length).setValues(output);
  }
   
    sheet.getRange(2, 6,nowtime.length,1).setValues(nowtime);

  //Logger.log(output);
}


function SetRectangleBorderRed(pID,shapeId){ 
  var deck;
  deck = SlidesApp.openById(pID);
  deck.getPageElementById(shapeId).asShape().getBorder().getLineFill().setSolidFill("#FF0000");
}

function SetRectangleBorderBlue(pID,shapeId){
  var deck;
  deck = SlidesApp.openById(pID);
  deck.getPageElementById(shapeId).asShape().getBorder().getLineFill().setSolidFill("#0000FF");
}

function SetPointRed(pID,shapeId){
  var deck, shape;
  deck = SlidesApp.openById(pID);
  shape = deck.getPageElementById(shapeId).asShape();
  shape.getBorder().getLineFill().setSolidFill("#FF0000");
  shape.getFill().setSolidFill("#FF0000");
}

function SetPointBlue(pID,shapeId){
  var deck, shape;
  deck = SlidesApp.openById(pID);
  shape = deck.getPageElementById(shapeId).asShape();
  shape.getBorder().getLineFill().setSolidFill("#0000FF");
  shape.getFill().setSolidFill("#0000FF");
}


function checkAnswerMain() {
  var url = 'https://docs.google.com/spreadsheets/d/1roi6w3nfbOtqK44tdKnuIWqI2iPdSJN7zjqdK0CqmlE/edit?usp=drive_web&ouid=106203360429558678203'
  var name = '改答案'
  var SS = SpreadsheetApp.openByUrl(url);
  var sheet = SS.getSheetByName(name);
  var sheet2;
  var copylist = [];
    //log for count
  var execution_list = [];
  var nowtime = [];
  var sheetlog; 
    //log for count  
  var deck,deck2,slide,slide2,Ele,shape,shape2,ppturl,ID;
  var slides,slides2;
  var push,push2;
  var c_range = 3,judge,j1;
  var answer_list=[];
  var i=0,j,k,l;
  
  copylist = sheet.getRange(2, 1, sheet.getLastRow()-1,2).getValues();
  
    
//Logger.log(copylist,"\n");
  
  for(l=0;l<copylist.length;l++){

    if(copylist[l][0]==1){
//Logger.log(copylist[l][1],"\n");
  sheet2 = SpreadsheetApp.openByUrl(copylist[l][1]).getSheetByName('_Record');
  answer_list = sheet2.getRange(2, 1, sheet2.getLastRow()-1,8).getValues();

  //Logger.log(answer_list,"\n");
    //判定對錯
    
  for(j=0;j<answer_list.length;j++){  
    //Logger.log(answer_list[j][5],answer_list[j][6],"\n");
 if(answer_list[j][7]==0||answer_list[j][7]=="FALSE"){
    //Logger.log('yes');
    
      if(answer_list[j][2]=="關鍵點"){
        SetPointRed(answer_list[j][5],answer_list[j][6]);
      
        Logger.log('修改錯誤',answer_list[j][0],'第',answer_list[j][1],'題');
      
      }else if(answer_list[j][2]=="拉框"){
        SetRectangleBorderRed(answer_list[j][5],answer_list[j][6]);
        Logger.log('修改錯誤',answer_list[j][0],'第',answer_list[j][1],'題');      
      }else{Logger.log('no');}
    
  }else{
       if(answer_list[j][2]=="關鍵點"){
        SetPointBlue(answer_list[j][5],answer_list[j][6]);
      
        Logger.log('修改正確',answer_list[j][0],'第',answer_list[j][1],'題');
      
      }else if(answer_list[j][2]=="拉框"){
        SetRectangleBorderBlue(answer_list[j][5],answer_list[j][6]);
        Logger.log('修改正確',answer_list[j][0],'第',answer_list[j][1],'題');
      }else{Logger.log('no');}
       
  }
    
  }
  
  //寫入回饋
  //slides[0].getShapes()[0].getText().setText("錯誤題數："+count_err+"\n良率："+(1-count_err.length/(correct.length-1))).getTextStyle().setForegroundColor('#ff0000').setFontSize(30);
      sheet.getRange(l+2, 3).setValue(new Date());
  }
  }
    
}



