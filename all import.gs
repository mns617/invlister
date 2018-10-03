

function makeBatchAds()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    var rng=sheet.getActiveRange();
    var sr=rng.getRow();
    var er=rng.getLastRow();
       
    for (var i=sr; i<=er; i++)
    {
          if(sheet.getRange(i, 8).getValue()==""){continue};
          sheet.getRange(i, 8).activate();
           nfl_throw();
        //   nflFootballMats()
           // nflFootballMats()           //nfl_Caps("", "");
          //nfl_Bathrobe("", "")
         // nfl_throw("", "")
          //nfl_throwPillow("", "");
         // nfl_GolfTowels("", "")
          //nfl_showerCurtain(5, 5);
         // nfl_BathTowels("", "");
          //nfl_rugs("", "")
    
    
    }

    


}
















function findTeam(title, col1, col2)
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet1=ss.getSheetByName('Sports Bedding');
    var values1=sheet1.getDataRange().getValues();
    
    
    var sheet2=ss.getSheetByName('Mapping3');
    var values2=sheet2.getDataRange().getValues();
    
    for (var i=1; i<values1.length; i++)
    {
        var tempTeamName=values1[i][col1-1];
        if(title.indexOf(tempTeamName)>=0)
        {
              break;
        
        }
        
    
    }
    var team=values1[i][col1-1]; 
    var color=values1[i][col1-1+1]; //column to right
    var longName=values1[i][col1-3]; 
    
    var flag=0;
    var material="";
    for (var i=1; i<values2.length; i++)
    {
        var tempMaterial=values2[i][1-1];
        if(title.indexOf(tempMaterial)>=0)
        {
              flag=1; break;
        
        }
        
    
    }
    
    
    if(flag==1)
     { material=values2[i][1-1];} 
    
     return [team, color, longName, material]
      


}





















function allImport() {

        var ss=SpreadsheetApp.getActiveSpreadsheet();
        var sheet=ss.getActiveSheet();
        var rng=sheet.getActiveRange();
        
          var mode='nzcu';

        if(rng.getColumn()!=1)
        {Browser.msgBox("Put this link in column A adn retry"); return 0};
        
        
        var url=sheet.getActiveRange().getValue();
        
        
        
        var getRow=rng.getRow();
        var startRow=getRow;
        var row=getRow; 
        
        
         var prevRow=getRow-1;
         var repeatFrm='=IF(COUNTIF(F$8:F'+ prevRow+',$F'+getRow+')>0, "Repeat", IFERROR(IF($F'+getRow+'<>"",JOIN(",",UNIQUE(FILTER(Database!$B$1:$B,Database!$A$1:$A=$F'+getRow+'))),""),"New"))';
     
        
        if(url.indexOf('overstock')>=0)
        {
           allImportOs(url)
           return 0
        
        }
        
        

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
        arr.push(blankArr); arr.push(blankArr); arr.push(blankArr);
        for (var i in myItems)
        {
               var myItem=myItems[i];
               var itemNo=myItem.usItemId;
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
               
               
               
               var myVariant=myItem.variants;
               
               if(myVariant==undefined) // no variation
               {
                   arr.push(blankArr); arr.push(blankArr); arr.push(blankArr); //three more blank rows
               
                   continue;
                 }
               var myVariants=myVariant.variantData;
               
               for ( var i in myVariants)
               {
                 
                   arr.push(blankArr);
               
               }
               arr.push(blankArr); arr.push(blankArr); //two more blank rows
               
               
        
        }
        
        sheet.getRange(row+2, 1, arr.length, 8 ).setValues(arr);
        var a=10






  
}









function allImportOs(url)
{

        var ss=SpreadsheetApp.getActiveSpreadsheet();
        var sheet=ss.getActiveSheet();
        var rng=sheet.getActiveRange();
        
        var mode='nzcu';
       // var url="https://www.overstock.com/Sports-Toys/Football/Nfl-Hats,/k,/6586/subcat.html?keywords=Nfl%20Hats&searchtype=Header";
        
        
        
        var getRow=rng.getRow();
        var startRow=getRow;
        var row=getRow; 
        
        
         var prevRow=getRow-1;
         var repeatFrm='=IF(COUNTIF(F$8:F'+ prevRow+',$F'+getRow+')>0, "Repeat", IFERROR(IF($F'+getRow+'<>"",JOIN(",",UNIQUE(FILTER(Database!$B$1:$B,Database!$A$1:$A=$F'+getRow+'))),""),"New"))';
     
        
        
        
        

        var col=rng.getColumn();
        if(!(isLister(sheet))){return 0}
        
        
         var option = {
                      'muteHttpExceptions' : true
          };

        var html = UrlFetchApp.fetch(url, option).getContentText();
        
        var n1=html.indexOf('window.__INITIAL_STATE__=')+('window.__INITIAL_STATE__=').length;
        var n2=html.indexOf(';window.__HAS_RESULTS__=true;',n1);
        var html2=html.slice(n1,n2);
        
        var jsonData=JSON.parse(html2)
        var myItems=jsonData.products;
        var myItems=myItems[Object.keys(myItems)[0]].products;
        
        var arr=[];
        var blankArr=["", "", "", "", "", "", "", ""];
        arr.push(blankArr); arr.push(blankArr); arr.push(blankArr);
        for (var i in myItems)
        {
               var myItem=myItems[i];
               var itemNo=myItem.sku;
               var prodUrl=myItem[Object.keys(myItem)[0]].productPage;
               var wmTitle="";//myItem.title;

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
               
               
               
               var myVariant=myItem.variants;
               
               if(myVariant==undefined) // no variation
               {
                   arr.push(blankArr); arr.push(blankArr); arr.push(blankArr); //three more blank rows
               
                   continue;
                 }
               var myVariants=myVariant.variantData;
               
               for ( var i in myVariants)
               {
                 
                   arr.push(blankArr);
               
               }
               arr.push(blankArr); arr.push(blankArr); //two more blank rows
               
               
        
        }
        
        sheet.getRange(row+2, 1, arr.length, 8 ).setValues(arr);





















}
















function importAllEachAd()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    var rng=sheet.getActiveRange();
    var sr=rng.getRow();
    var er=rng.getLastRow();
       
    for (var i=sr; i<=er; i++)
    {
          if(sheet.getRange(i, 8).getValue()==""){continue};
          if(sheet.getRange(i, 1).getValue()!=""){continue};
          sheet.getRange(i, 8).activate();

          var n= importFromSource1();
         // i=i+n; //increase by number of variations
    
    
    }

}



