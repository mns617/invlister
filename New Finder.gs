
function WMnewFinder() {

        var ss=SpreadsheetApp.getActiveSpreadsheet();
        var sheetDB=ss.getSheetByName('Database');
        var sheet=ss.getSheetByName('Lister');
        
        var mode='nzcu';
        var dbValues=sheetDB.getRange("A2:A").getValues();
        var exlusionVales=ss.getSheetByName('Database Exclusion').getRange("A1:A").getValues();
        var listerVales=ss.getSheetByName('Lister Amirul').getRange("F1:F").getValues();      
        
        var oldItemNos=dbValues.join("|").split("|").concat(exlusionVales.join("|").split("|")).concat(listerVales.join("|").split("|"));
                
      for (var p=1; p<=15;p++)
      {
        var url='https://www.walmart.com/search/?cat_id=4044_539095&facet=retailer%3AWalmart.com&grid=true&page='+p+'&po=1&query=towel+sets&typeahead=towel+set&vertical_whitelist=home%2Cfanatics%2C#searchProductResult'
        var getRow=sheet.getLastRow()+5;
        var startRow=getRow;
        var row=getRow; 
        var rng=sheet.getRange(row, 1);
        
         var prevRow=getRow-1;
         var repeatFrm='=IF(COUNTIF(F$8:F'+ prevRow+',$F'+getRow+')>0, "Repeat", IFERROR(IF($F'+getRow+'<>"",JOIN(",",UNIQUE(FILTER(Database!$B$1:$B,Database!$A$1:$A=$F'+getRow+'))),""),"New"))';
     
        
        
        

        var col=rng.getColumn();
        if(!(isLister(sheet))){return 0}
        
        
         var option = {
                      'muteHttpExceptions' : true
          };

        var html = UrlFetchApp.fetch(url, option).getContentText();
        var jsonData=getMyJsonSearch(html);
        var myItems=jsonData.preso.items
        var arr=[];
        var blankArr=["", "", "", "", "", "", "", ""];
        var header=["Page No "+p, "", "", "", "", "", "", ""];
              
        var oldItemNos=dbValues.join("|").split("|");
        arr.push(header)
        arr.push(blankArr); arr.push(blankArr); arr.push(blankArr);
        
        for (var i in myItems)
        {
               var myItem=myItems[i];
               var itemNo=myItem.usItemId;
               
               if(oldItemNos.indexOf(itemNo)>=0){continue};
               if(oldItemNos.indexOf(itemNo.toString())>=0){continue};
               
               
               var prodUrl="https://www.walmart.com"+myItem.productPageUrl;
               var wmTitle="";//myItem.title;
               wmTitle=replaceAll(wmTitle, "<mark>", "");
               wmTitle=replaceAll(wmTitle, "</mark>", "");
               var initial="";
               var date="";
               
               var sku="";
               var skugridVar="";
               
               var prevRow=row+2+arr.length-1;
               getRow=prevRow+1;
               var repeatFrm='=IF(COUNTIF(F$8:F'+ prevRow+',$F'+getRow+')>0, "Repeat", IFERROR(IF($F'+getRow+'<>"",JOIN(",",UNIQUE(FILTER(Database!$B$1:$B,Database!$A$1:$A=$F'+getRow+'))),""),"New"))';
               var asin=repeatFrm;
        
               
               var tempArr=[wmTitle, initial, date, asin, sku, itemNo, skugridVar, prodUrl];
               arr.push(tempArr);
               
               
                      
        }
        
        sheet.getRange(row+2, 1, arr.length, 8 ).setValues(arr);

      }



  
}








function OSnewFinder() {

        var ss=SpreadsheetApp.getActiveSpreadsheet();
        var sheetDB=ss.getSheetByName('Database');
        var sheet=ss.getSheetByName('Lister');
        
        var mode='nzcu';
        var dbValues=sheetDB.getRange("A2:A").getValues();
                
  for(var p=1; p<=11; p++)    
  {
     
       // var pageNo=sheet.getRange("A1").getValue();
        var url="https://www.overstock.com/Bedding-Bath/Towels/Animal-Print~Geometric~Jacquard~Stripe~check~chevron~dot~embroidered~floral~graphic~holiday~kids~medallion~nature~novelty~ombre~plaid~wildlife,/pattern,/4529/cat.html?page="+p+"&sort=Avg.%20Customer%20Rating"
        var getRow=sheet.getLastRow()+5;
        var startRow=getRow;
        var row=getRow; 
        var rng=sheet.getRange(row, 1);
        
         var prevRow=getRow-1;
         var repeatFrm='=IF(COUNTIF(F$8:F'+ prevRow+',$F'+getRow+')>0, "Repeat", IFERROR(IF($F'+getRow+'<>"",JOIN(",",UNIQUE(FILTER(Database!$B$1:$B,Database!$A$1:$A=$F'+getRow+'))),""),"New"))';
     
        
        
        

        var col=rng.getColumn();
        if(!(isLister(sheet))){return 0}
        
        
         var option = {
                      'muteHttpExceptions' : true
          };

       var html = UrlFetchApp.fetch(url, option).getContentText();
        
        var n1=html.indexOf('window.__INITIAL_STATE__=')+('window.__INITIAL_STATE__=').length;
        var n2=html.indexOf('window.__HAS_RESULTS__=true;',n1);
        var n3=html.lastIndexOf("}",n2)+1;
        var html2=html.slice(n1,n3);
      //  GmailApp.sendEmail("sakib118.biz@gmail.com", "test", html2)
        var jsonData=JSON.parse(html2)
        var myItems=jsonData.products;
        var myItems=myItems[Object.keys(myItems)[0]].products;
        
        
        var arr=[];
        var blankArr=["", "", "", "", "", "", "", ""];
        var oldItemNos=dbValues.join("|").split("|");
        
        arr.push(blankArr); arr.push(blankArr); arr.push(blankArr);
        
        for (var i in myItems)
        {
               var myItem=myItems[i];
              
               
              
               var itemNo=myItem.sku;
               
               if(oldItemNos.indexOf(itemNo)>=0){continue};
               if(oldItemNos.indexOf(itemNo.toString())>=0){continue};
               
               var prodUrl=myItem[Object.keys(myItem)[0]].productPage;
               
             //  var prodUrl="https://www.walmart.com"+myItem.productPageUrl;
               var wmTitle="";//myItem.title;
               wmTitle=replaceAll(wmTitle, "<mark>", "");
               wmTitle=replaceAll(wmTitle, "</mark>", "");
               var initial="";
               var date="";
               
               var sku="";
               var skugridVar="";
               
               var prevRow=row+2+arr.length-1;
               getRow=prevRow+1;
               var repeatFrm='=IF(COUNTIF(F$8:F'+ prevRow+',$F'+getRow+')>0, "Repeat", IFERROR(IF($F'+getRow+'<>"",JOIN(",",UNIQUE(FILTER(Database!$B$1:$B,Database!$A$1:$A=$F'+getRow+'))),""),"New"))';
               var asin=repeatFrm;
        
               
               var tempArr=[wmTitle, initial, date, asin, sku, itemNo, skugridVar, prodUrl];
               arr.push(tempArr);
               
               
                      
        }
        
        sheet.getRange(row+2, 1, arr.length, 8 ).setValues(arr);
        sheet.getRange("A1").setValue(p+1);


 }


  
}