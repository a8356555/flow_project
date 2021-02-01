function GetSlideNote(url,output)
{
   var deck, slides, j, k, text, name;
   deck = SlidesApp.openByUrl(url);
   slides = deck.getSlides();
   name = deck.getName();
  for(j=0;j<slides.length;j++){ 
        
    text = slides[j].getNotesPage().getSpeakerNotesShape().getText().asString();
    output.push([name,j+1,text])
   }  
  return output;
  }

function DoThis() {
  var url = 'https://docs.google.com/spreadsheets/d/1p1xMGoJN2EHu5kzrvAaoYJsUGvL4KSgZqVyl7F82AX4/edit#gid=0';
  var name = 'PasteTextToSheet';
  var SS = SpreadsheetApp.openByUrl(url),SS2,sheet2;
  var sheet = SS.getSheetByName(name);
  var copylist = [],answer=[],correct=[],output = [];
  
  copylist = sheet.getRange(2, 2, sheet.getLastRow()-1,4).getValues();
 
  output[0]=["PPT名稱","頁次","Note內容"]
  for(l=0;l<copylist.length;l++){
    try{
      ////////////////抓作答
      GetSlideNote(copylist[l][0],output)

  //Logger.log(output);
  sheet.getRange(l+2, 4).setValue(new Date());
  }catch(f){
    sheet.getRange(l+2, 4).setValue(f);
  } 
  
    }
  SS2 = SpreadsheetApp.openByUrl(copylist[0][1]);
  sheet2 = SS2.getSheetByName('_SlidesNote');
  if(sheet2=="Sheet"){//如果試算表已存在
      sheet2.getRange(SS2.getLastRow()+1,1,output.length,output[0].length).setValues(output);
  }else{
    SS2.insertSheet(0).setName('_SlidesNote').getRange(1,1,output.length,output[0].length).setValues(output);
  }
  //Logger.log(output);
}

