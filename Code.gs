var databaseSsId="1nwJE0i3qTvjO8KW8BhneMYOCOVf74hVgDKoE7mx9wmE"; //original inventory list
var folderId="0B15F6wpL3VKOWF9KMGpOWUYtMFU";
var liveId=SpreadsheetApp.getActiveSpreadsheet().getId();


var mode='nzcu';
var allLenFrm='="N:"&LEN(R[0]C[2])&"; B4:"&LEN(R[0]C[8]) & "; T:"&(LEN(R[0]C[9])+LEN(R[0]C[23]))';
var amPrice1='=ROUNDUP(1.4*R[0]C[16])-0.01';  // on sale
var amPrice2='=ROUNDUP(1.3*R[0]C[16])-0.01';  // not on sale

function onOpen() {
  var ui = SpreadsheetApp.getUi();
 
  /*  
    ui.createMenu('Transfer')
      .addItem('Transfer to UPLOAD', 'transferToUpload')
      .addToUi();
  */    
      
    ui.createMenu('Scripts')
      .addItem('Import Ad','importFromSource1')
      .addItem('Activate Rows', 'activateFormulas')
      .addItem('Deactivate Rows', 'deactivateFormulas')
      .addItem('Mark as Review', 'onEdit3')

      .addSeparator()

      //.addItem("Download Images", 'getOsImages')
      .addItem('Show Images', 'imShowSideBar')
      .addItem("Show Detail Add", 'showSidebar')
      .addItem("Automate Terms", 'onEditAmTitle')
      .addItem("Check Errors", 'manualValidity')
      .addItem("Show Description", 'getOverstockDescription')
      
      .addToUi();   
      
      ui.createMenu('Formulas')
      .addItem('Turn Off Formulas', 'turnOffFormulas')
      .addItem('Turn On Formulas', 'turnOnFormulas')
      .addItem('Import Each from Source1', 'importAllEachAd')
      .addToUi();   
       
      
      ui.createMenu('Supervisor')
      .addItem('Check Images by Variation', 're_checkImagesByVariation')
      .addItem('Check Images by Image Position', 're_checkImagesByImagePosition')
      .addItem('Variation Checking', 'variationChecking')
      .addItem('Complete All', 'markAllCompleted')
      .addItem('Import All', 'allImport')
      .addItem('Make Batch Ads', 'makeBatchAds')
      .addItem('Make NFL Caps','nfl_Caps')
      .addItem('Make NFL Rugs', 'nfl_rugs')
      .addItem('Make NFL Throw', 'nfl_throw')
      .addItem('Make NFL Disney Throw', 'nfl_Dysney_throw')
      .addItem('Make Hand Towel', 'nfl_HandTowels')
      
      .addToUi();   
       
      
      
      
}



function checkErrorAllRows()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheets=ss.getSheets();
    
    for (var i=0; i=sheets.length; i++)
    {
          
    }



}













function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Add Texts')
      .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(html);
}





function importProfitCalculator()
{
        var ss=SpreadsheetApp.openById('1I_Etaz0BJ4K40K7FQjqedFRycbdl2y2VZO8s6OXb6Gs');
        var sheet=ss.getSheetByName('Lister');
        var frms=sheet.getRange("R1:U7").getFormulas();
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Lister').getRange("R1:U7").setFormulas(frms);



}


function completeAllSheets()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet()
    var sheets=ss.getSheets();
    
    for (var i=0; i<sheets.length; i++)
    {
        
        completeSheet_(sheets[i])
    
    }
    
    

   SpreadsheetApp.getActiveSpreadsheet().copy(SpreadsheetApp.getActiveSpreadsheet().getName()+Utilities.formatDate( new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd-MMMM-yyyy"))


}

function cleanUpAllSheets()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet()
    var sheets=ss.getSheets();
    
    for (var i=0; i<sheets.length; i++)
    {
        
        cleanUpSheet_(sheets[i]);
    
    }
    
    

   SpreadsheetApp.getActiveSpreadsheet().copy(SpreadsheetApp.getActiveSpreadsheet().getName()+Utilities.formatDate( new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd-MMMM-yyyy"))


}


function completeSheet_(sheet)  //clears posted sheet
{
      
      
      if(isLister(sheet)==false){return 0}
      
      
      
      var values=sheet.getDataRange().getValues();
      var colD=sheet.getRange(1, 4,values.length,1).getValues();
      var colDFrms=sheet.getRange(1, 4,values.length,1).getFormulasR1C1();
      

      for (var i= 20; i<values.length; i++)
      {
            var status=values[i][11-1];
            var frm=colDFrms[i][0]
            //if (colDFrms[i][0]!=""){colD[i][0]==colDFrms[i][0]};
            if(status=="COMPLETE")
            {
                   var getRow=i+1; 
                   var prevRow=getRow-1;
                   colD[i][0]='=IFERROR(IF($F'+getRow+'<>"",JOIN(",",UNIQUE(FILTER(Database!$B$1:$B,Database!$A$1:$A=$F'+getRow+'))),""),"")';
            
            }
            
            else if(frm!="")
            {
                      var r=i+1; 
                      var pr=r-1;  
                      colD[i][0]='=IF(COUNTIF(F$8:F'+pr+',$F'+r+')>0, "Repeat", IFERROR(IF($F'+r+'<>"",JOIN(",",UNIQUE(FILTER(Database!$B$1:$B,Database!$A$1:$A=$F'+r+'))),""),"New"))';
            }  
      
      }

      sheet.getRange(1, 4,values.length,1).setValues(colD);


}














function cleanUpSheet_(sheet)  //clears posted sheet
{
      
      if(isLister(sheet)==false){return 0}

      var values=sheet.getDataRange().getValues();

      for (var i= 20; i<values.length; i++)
      {
            var status=values[i][11-1];
            
            var init=replaceAll(values[i][2-1].trim()," ","").toLowerCase();
            var prevInit=replaceAll(values[i][4-1].trim()," ","").toLowerCase();
            
            if(status!="COMPLETE"){continue};
            {
                if(prevInit.indexOf(init)>=0)
                {
                    sheet.getRange(i+1, 1, 1, sheet.getMaxColumns()).clearContent();
                
                }
            
            }
      
      }




}



function deleteRows_()  //delete blanks rows posted sheet
{
      var ss=SpreadsheetApp.getActiveSpreadsheet()
      var sheet=ss.getActiveSheet();
      
      var values=sheet.getRange(1,1,sheet.getMaxRows(), sheet.getMaxColumns()).getValues();
      var count=0;// how any rows
      for (var i= values.length-1; i>15; i--)
      {
            var source1=values[i][8-1];

            if(source1=="")//consecutive blank rows
            {
                count++
            
            }
            
            
            
            
            if(source1!="")
            {
                 if(count>=13) //3 consequtive blank rows found
                 {
                  sheet.deleteRows(i+1+3, count-3);
                 } 
                  
                  count=0; //reset count
            
            }
            
             Logger.log(count)
            
      
      }

      


}






function makeCSV4Upload()
{

      var ss=SpreadsheetApp.getActiveSpreadsheet()
      var csvSheet=ss.getSheetByName('CSV');
      csvSheet.getRange(2, 1, csvSheet.getMaxRows()-1, csvSheet.getMaxColumns()).clearContent();
      var sheets=ss.getSheets();
      var arrs=[];
     
   for(var s=0; s<sheets.length; s++)      
   {   
                
                
                
                var sheet=sheets[s];
                var lr=sheet.getLastRow();
                lr=lr<6?6:lr;
                
                if(sheet.getName().indexOf("Lister")!=0){continue};
                var values=sheet.getRange(6,1,lr-6+1, 52).getValues();
                
                for (var i= 10; i<values.length; i++)
                {
                      var status=values[i][11-1];
                      var init=values[i][2-1];
                      var prevInit=values[i][4-1]
                      
                      if(status=="COMPLETE")
                      {
                            if(prevInit.indexOf(init)<0) //not posted
                            {
                                  arrs.push(values[i]);
                            
                            }
                      
                      }
                
                }//end of i for 

   }


    csvSheet.getRange(2, 1, arrs.length, 52).setValues(arrs)




}



function addProfit()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getSheetByName("CSV")
    
    var values=sheet.getDataRange().getValues();
    
    var ssLive=SpreadsheetApp.openById(liveId);
    var sheetLister=ssLive.getSheetByName("Lister");
    
    for (var i=1; i<values.length; i++)
    {
        var url=values[i][8-1];
        
        if(url.indexOf('overstock.com')>0)
        {
                    var isSale=isOnSale2(url);
                                       
                    
                    if(isSale==false)
                    {
                        var profit=sheetLister.getRange("U2").getFormulaR1C1();
                        var amPrice= "=ROUND((R[0]C[17]-(-(R[0]C[16])+((R[0]C[16]*0.12))+((R[0]C[16])-(R[0]C[16]*0.12))*0.0688))/0.85,0)-0.01";
                        }
            
                    else
                     {
                        var profit=sheetLister.getRange("U3").getFormulaR1C1();
                        var amPrice="=ROUND(((R[0]C[17]-(-(R[0]C[16])+(R[0]C[16]*0.1188)))/0.85),0)-0.01";
        
                        }
                    sheet.getRange(i+1, 32).setFormulaR1C1(profit);
                    sheet.getRange(i+1, 15).setValue(amPrice);
                    
        
        }// if os
        
        
        else if(url.indexOf('walmart.com')>0)
        {
                        var profit=sheetLister.getRange("U5").getFormulaR1C1();
                        var amPrice='=ROUNDUP(R[0]C[17]-(-R[0]C[16]-((R[0]C[16]*0.06))+(R[0]C[16]*0))/0.85)-0.01';
                        sheet.getRange(i+1, 32).setFormulaR1C1(profit);
                        sheet.getRange(i+1, 15).setValue(amPrice);
              
        }
        
    
    }//end of for


}




















function importFromSource1()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssLive=SpreadsheetApp.openById(liveId);
  var sheet = ss.getActiveSheet();
  
  var rng = sheet.getActiveRange();
  var row = rng.getRow();
  var col = rng.getColumn();
  LockService.getScriptLock().releaseLock()
  var gLock=LockService.getScriptLock();
  var a= sheet.getRange('J2429').getFormulaR1C1();
    
  
  //when entering soruce 1
  if(col==8 && isLister(rng.getSheet()))
  {
       var sourceUrl=rng.getValue();     
       if(sourceUrl.indexOf('overstock')>=0)
            {              
                        getOverstockData(rng) 
            
            }
          
       else if(sourceUrl.indexOf('walmart')>=0)  
            {
                            importWMdata(rng);
            }
       
       
        if(mode=='nzcu')  //repeat formula in first row
              {     
                    var getRow=row; 
                    var prevRow=getRow-1;
                    var repeatFrm='=IF(COUNTIF(F$8:F'+ prevRow+',$F'+getRow+')>0, "Repeat", IFERROR(IF($F'+getRow+'<>"",JOIN(",",UNIQUE(FILTER(Database!$B$1:$B,Database!$A$1:$A=$F'+getRow+'))),""),"New"))';
                    
                    //var repeatFrm='=IF(COUNTIF(F8:F'+prevRow+',F'+getRow+')+COUNTIF(Database!A1:A,F'+getRow+')=0,"New","Repeat")';
                    sheet.getRange(getRow, 4).setValue(repeatFrm);
              }
           
         
         
         
     
  }
    
    
    
    

}
















function showadd()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    var row=sheet.getActiveRange().getRow();
    
    
    var values1=sheet.getRange(8, 1,1,32).getValues();
    var values2=sheet.getRange(row, 1,1,32).getValues();
    
    var body="<p>";
    
    for (var i=0; i<values1[0].length; i++)
    {
          if(values1[0][i]==""){continue}
          body=body+'<b>'+values1[0][i]+":</b><br>"+values2[0][i]+"<br><br>";
          
    
    }

             return replaceAll(body,"_"," ");
}









//batch operation for rows to complete
function markAllCompleted()
{

            var ss=SpreadsheetApp.getActiveSpreadsheet();
            var sheet=ss.getActiveSheet();
    
             var rng=ss.getActiveRange();
             var rowO=rng.getRow();  //original row
            rng=sheet.getRange(rowO, 1, rng.getValues().length, 51);
            var valuesAll=rng.getValues();
            var mapValues=ss.getSheetByName("Mapping").getDataRange().getValues();
            var ddValues=ss.getSheetByName("Dropdown Options").getDataRange().getValues();
            
            
            for (var i=0; i<valuesAll.length; i++)
            {
                for (var j=0; j<valuesAll[0].length; j++)
                {
                    Logger.log(valuesAll[i][j].toString().toLowerCase())
                    if(valuesAll[i][j].toString().toLowerCase().indexOf('error')>=0)
                    {
                        Browser.msgBox("Error found in "+ Number(Number(rowO)+Number(i)));
                        return 0
                    }
                
                }
            
            }
            
            
      
            for (var r=0; r<valuesAll.length; r++)
            {     
                                            var values=[valuesAll[valuesAll.length-r-1]] ; //a 2D array just for this row
                                            if(values[0][11-1]==""){continue};  //skip if empty status found
                                           //following columsn will be done when it is marked as complete
                                           
                                           
                                           
                                           ///check empty columns
                                           for (var i=0; i< 25; i++)
                                           {
                                                   
                                                   //skip non mandatory columns
                                                   if(i==4||i==5 || i==6 || i==9 || i==8 || i== 12 || i==3){continue}
                                                   
                                                   if(values[0][i]=="")
                                                   {
                                                         var msg=sheet.getRange(rowO+r, i+1).getA1Notation()+" is empty!"
                                                         Browser.msgBox(msg);
                                                         sheet.getRange(rowO+r, 11).clearContent();
                                                         return 0;
                                                   }
                                                   
                                                   
                                                   if(i==12-1)
                                                   {
                                                                 var temp=replaceAll(values[0][i],"_"," ");
                                                                 values[0][i]=toTitleCase(temp.trim())
                                                   
                                                   }
                                                   
                                                   
                                                           
                                                   
                                                   if(i==16-1 || i==17-1 || i==18-1)
                                                   {
                                                                 var temp=replaceAll(values[0][i],"_"," ");
                                                                 values[0][i]=toSentence(temp.trim());
                                                   
                                                   }
                                                   
                                                   
                                                   if(i==19-1)
                                                   {
                                                                 var temp=replaceAll(values[0][i],"_"," ");
                                                                 values[0][i]=(temp.toLowerCase().trim());
                                                   
                                                   }
                                           
                                           }  //end of for
                                           
                                           
                                               // sheet.getRange(row, 2).setValue(Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MM/dd/yyyy")); //set todays date
                                                
                                          var row=rowO+valuesAll.length-r-1    
                                          var terms=""; //combineTerms(values).toLowerCase();
                                          sheet.getRange(row, 19).setValue(terms);
                                           
                                           var amTitle=values[0][12-1];//product name
                                          
                                          
                                          if(amTitle.toString().length>200){
                                            Browser.msgBox("Invalid Title!");
                                            return 0;
                                          } 
                                           
                                           
                                           
                                          values[0][11-1]="COMPLETE"; 
                                          sheet.getRange(row,1,1,values[0].length).setValues(values);
                                    
                                         
                                           
                                           
                                           
                                         sheet.getRange(row,1,1,23).setBackground("#ffffff");
                                         sheet.getRange(row,1,1,23).setFontLine("line-through");
                                        
                                              
                                         if(mode=='nzcu')
                                              {     
                                                var getRow=row; 
                                                var prevRow=getRow-1;
                                                var repeatFrm='=IFERROR(IF($F'+getRow+'<>"",JOIN(",",UNIQUE(FILTER(Database!$B$1:$B,Database!$A$1:$A=$F'+getRow+'))),""),"")';
                                                
                                                //var repeatFrm='=IF(COUNTIF(F8:F'+prevRow+',F'+getRow+')+COUNTIF(Database!A1:A,F'+getRow+')=0,"New","Repeat")';
                                                sheet.getRange(getRow, 4).setValue(repeatFrm);
                                              }
            
      
      }  //end of r for






}


function markAllasReview()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
  
    var rng = sheet.getActiveRange();
    var row = rng.getRow();
    var lr=rng.getLastRow();
    var col = rng.getColumn();
    
    var valuesAll=sheet.getRange(row, 1, lr-row+1, sheet.getLastColumn()).getValues();
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    

}








function onEdit3(e) {

 
  //Browser.msgBox("Script triggered")
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var ss2=SpreadsheetApp.openById(databaseSsId);
  var ssLive=SpreadsheetApp.openById(liveId);
  var sheet = ss.getActiveSheet();
  
  var rng = sheet.getActiveRange();
  var row = rng.getRow();
  var col = rng.getColumn();

  var gLock=LockService.getScriptLock();
  //gLock.releaseLock();
 // gLock.waitLock(30000);
  var a= sheet.getRange('J2429').getFormulaR1C1();
    
  
  
  
  
  
  
  
  
  
  if(col==11)      //when marks as complete
  {
       var values=sheet.getRange(row,1,1,sheet.getLastColumn()).getValues();
       var msg= checkUndefined(values[0],row);
       
       if(msg!="")
       {
             Browser.msgBox(msg);
             return 0;
       }
       
       if(values[0][11-1]=="REVIEW")
       {
           
                   sheet.getRange(row,1,1,23).setBackground("#93c47d");
                   sheet.getRange(row, 3).setValue(Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MM/dd/yyyy"));
                   
                    
                 //skip non mandatory columns
                 for (var i=0; i< 25; i++)
                 {
                 
                         //skip non mandatory columns
                         if(i==4||i==5 || i==6 || i==9 || i==8 || i== 12 || i==3){continue}
                         
                         if(values[0][i]=="")
                         {
                               var msg=sheet.getRange(row, i+1).getA1Notation()+" is empty!"
                               Browser.msgBox(msg);
                               sheet.getRange(row, 11).clearContent();
                               return 0;
                         }
                }
                 
                 
                 
               
                   return 0;

       
       }  // end of review
       
       
       
       if(values[0][11-1]=="COMPLETE"){return 0};
       
       
       if(values[0][11-1]!="COMPLETE"){return 0};
       //following columsn will be done when it is marked as complete
       
       ///check empty columns
       for (var i=0; i< 25; i++)
       {
               
               //skip non mandatory columns
               if(i==4||i==5 || i==6 || i==9 || i==8 || i== 12 || i==3){continue}
               
               if(values[0][i]=="")
               {
                     var msg=sheet.getRange(row, i+1).getA1Notation()+" is empty!"
                     Browser.msgBox(msg);
                     sheet.getRange(row, 11).clearContent();
                     return 0;
               }
               
               
               if(i==12-1)
               {
                             var temp=replaceAll(values[0][i],"_"," ");
                             values[0][i]=toTitleCase(temp.trim())
               
               }
               
               
                       
               
               if(i==16-1 || i==17-1 || i==18-1)
               {
                             var temp=replaceAll(values[0][i],"_"," ");
                             values[0][i]=toSentence(temp.trim());
               
               }
               
               
               if(i==19-1)
               {
                             var temp=replaceAll(values[0][i],"_"," ");
                             values[0][i]=(temp.toLowerCase().trim());
               
               }
               
               
               
               
               
           
               
       
       
       }  //end of for
       
       
           // sheet.getRange(row, 2).setValue(Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MM/dd/yyyy")); //set todays date
            
            
                   var terms=combineTerms(values).toLowerCase();
                   sheet.getRange(row, 19).setValue(terms);
       
       
       
       
       
       
         sheet.getRange(row,1,1,values[0].length).setValues(values);

         var amTitle=values[0][12-1];//product name
       
       
         if(amTitle.toString().length>200){
               Browser.msgBox("Invalid Title!");
               return 0;
         }
       
       
       
        sheet.getRange(row,1,1,23).setBackground("#ffffff");
        sheet.getRange(row,1,1,23).setFontLine("line-through");
    
          
          if(mode=='nzcu')
          {     
            var getRow=row; 
            var prevRow=getRow-1;
             var repeatFrm='=IFERROR(IF($F'+getRow+'<>"",JOIN(",",UNIQUE(FILTER(Database!$B$1:$B,Database!$A$1:$A=$F'+getRow+'))),""),"")';
            
            //var repeatFrm='=IF(COUNTIF(F8:F'+prevRow+',F'+getRow+')+COUNTIF(Database!A1:A,F'+getRow+')=0,"New","Repeat")';
            sheet.getRange(getRow, 4).setValue(repeatFrm);
          }
          
    
    
    
 
     }
     
  
  
  
  
  
    gLock.releaseLock();

}






function replaceAll(string, find, replace) {
  return string.replace(new RegExp(escapeRegExp(find), 'g'), replace);
}

function escapeRegExp(string) {
    return string.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1");
}













function toTitleCase(str)
{
    return str.replace(/\w\S*/g, function(txt){return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();});
}

function toSentence(string) {
    return string.charAt(0).toUpperCase() + string.slice(1);
}




function getOsImages()
{


        
        
        var ss=SpreadsheetApp.getActiveSpreadsheet();
        var sheet=ss.getActiveSheet();
        var rng=sheet.getActiveRange();
        
        if(!(isLister(sheet)) || rng.getColumn()!=8){Browser.msgBox("Please select a valid Overstock link and try again"); return 0}
         var option = {
                      'muteHttpExceptions' : true
          };
         var getURL=rng.getValue(); 
         
         
         
         
         
         
        
      
       
        var html = UrlFetchApp.fetch(getURL, option).getContentText();
        var htmlOrig=html;
       
       
        var n1=html.indexOf('s-h-title');
        var n2=html.indexOf("<",n1);
        var title=html.slice(n1+11,n2-1) 
       
       
       var imPhrase=Browser.inputBox("Please enter image name phrase")
       
       
       
       
       
        var tempFolder=DriveApp.getFolderById(folderId).createFolder(Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd-MMMM-yyyy hh:mm a")+"--"+title);; 
        var folderUrl=tempFolder.getUrl();
       
       var n1=html.lastIndexOf("thumb-frame");
       var n2=html.indexOf('</ul>',n1)
       var html2=html.slice(n1,n2)
       
       var imgUrlArr=html2.split('data-max-img');

              for (var i=1; i<imgUrlArr.length; i++)//when there is variation, index 0 has garbage data
              {
                      var longUrl=imgUrlArr[i];
                      var l1=longUrl.indexOf("ak1");
                      var l2=longUrl.indexOf(">",l1);
                      var imUrl=longUrl.slice(l1,l2-1);
                      
                      //var imageURL=(imUrl).replace("ostkcdn.com","ostkcdn.com.rsz.io")+"?flip=x"
                     
                      var imageURL="http://res.cloudinary.com/demo/image/fetch/a_hflip/http://"+imUrl;
                      var imBlob=UrlFetchApp.fetch(imageURL).getBlob();
                      
                      var imFile=tempFolder.createFile(imBlob);
                      imFile.setName(imPhrase+" "+ i+".jpg");
                  
              }



} //end of function










function myLookup(val, mapVals, col)
{
     val=val.toLowerCase();
     for (var i=1; i<mapVals.length; i++)
     {
           if(mapVals[i][col-1]==""){continue};
           if(val.indexOf((mapVals[i][col-1]).toLowerCase())>=0)
           {return i+1}         
     } 
     
     return null;


}









function removeChars(validChars, inputString) {
    var regex = new RegExp('[^' + validChars + ']', 'g');
    return inputString.replace(regex, '');
}














function isOnSale(url)
{
  //var url="http://www.overstock.com/Home-Garden/HomePop-Large-Teal-Blue-Decorative-Storage-Ottoman/10293207/product.html";
  try
  {
      var html=UrlFetchApp.fetch(url).getContentText();
  }
  
  catch(e)
  {
      var r=true 
  }
  
  if(html==undefined){return true}
  
  var n1=html.indexOf("price-title");
  var n2=html.indexOf(">",n1);
  var n3=html.indexOf("<",n2);
  var priceTitle=html.slice(n2+1,n3);
  // GmailApp.sendEmail("sakib118.biz@gmail.com", "Sale", priceTitle)
  if(priceTitle.indexOf("Sale")>=0)
  {var r= true;}
  else
  {r= false;}
  
  if(html.indexOf('DoorBustersIcon')>0)
  {
    r=true;
  }
  
  if(html.indexOf('DoorbusterIcon')>0)  //for weekly flash deals
  {
    r=true;
  }
  
  
  
  
  
  return r;
}














function replaceUnwantedFromOs(title)
{

        var idx1 = title.indexOf("&amp;");
          if (idx1 > -1) {
            title = replaceAll(title, "&amp;", "&");
          }
          
          if (title.indexOf("#") > -1) {  // Fix special characters if found on the title
            for (var i=0; i<title.length; i++)
            {
              var char = title.charAt(i);
              if (char == '#') {
                var semicolon = title.indexOf(';');
                var spCode = title.slice(i-1, semicolon+1);
                var decodedVal = "";
                if (spCode == "&#x27;") {
                  decodedVal = "'";
                } else if (spCode == "&#x26;") {
                  decodedVal = "&";
                } else if (spCode == "&#x25;") {
                  decodedVal = "%";
                } else if (spCode == "&#x24;") {
                  decodedVal = "$";
                } else if (spCode == "&#x23;") {
                  decodedVal = "#";
                } else if (spCode == "&#x22;") {
                  decodedVal = '"';
                }
                
                title = title.replace(spCode, decodedVal);
              }
              
              
            }
          }



        return title;



}






function importVariationOs(html)
{
          
          
          var isSale=isOnSale(html);
          
          if(isSale=="No")
          {
                profit=sheetLister.getRange("F2").getFormulaR1C1();
                amPrice=sheetLister.getRange("D2").getFormulaR1C1();
    
    
           }
    
          else
           {
              profit=sheetLister.getRange("F3").getFormulaR1C1();
              amPrice=sheetLister.getRange("D3").getFormulaR1C1();
    
    
          }










}






















function last_row(sheet, col)
{
  //var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  //col=1
  var values=sheet.getDataRange().getValues(); //sheet.getRange(1, col,sheet.getLastRow(),1).getValues();
 
  
  for(var i=values.length-1; i>=0; i--)
  {
   if (values[i][0] != "")
   {break}
   
  }
  
 
  return i+1
  
}


























function myvariation(mainRng,keys,rkeys)
{
    
    if(mainRng==""){return ""};
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    var row=sheet.getActiveCell().getRow();
    var col=sheet.getActiveCell().getColumn();
    //var mainRng="This is shit."
    var vals=mainRng.toString();
    /*
    vals=replaceAll(vals,"(","");
    vals=replaceAll(vals,")","");
    vals=replaceAll(vals,"[","");
    vals=replaceAll(vals,"]","");
    */
    var valsArr=vals.split(" ");
    
   // var rkeys=sheet.getRange(row, 34,1,8).getValues();
    

    for (var i=0; i<keys[0].length; i++)
    {
          
          if(rkeys[0][i]==""){continue;}
          var find=(keys[0][i]).toString();
          var replace=(rkeys[0][i]).toString();
          
            for (var j=0; j<valsArr.length; j++ )
            {
                  var word=valsArr[j].toString();
                  
                  if(word.toLowerCase()==find.toLowerCase())
                  {
                        valsArr[j]=replace.toLowerCase();
                  }     
                   
                  else if(word.toLowerCase()==find.toLowerCase()+",")
                  {
                        valsArr[j]=replace.toLowerCase()+",";
                  } 
                   
                  else if(toTitleCase(word)==toTitleCase(find))
                  {
                        valsArr[j]=toTitleCase(replace);
                  }  
                  
                   else if(toTitleCase(word)==toTitleCase(find)+",")
                  {
                        valsArr[j]=toTitleCase(replace)+",";
                  }  
                  
                  
                  
                  
 
                  
            
            } 
    
    }
  
  
   return valsArr.join(" ")
  
  
  
  
  
  

}





