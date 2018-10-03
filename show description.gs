function getOverstockDescription(rng) 
{
  var rng =SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange();
  var editedSheet = rng.getSheet().getName();
  var getRow = rng.getRow();
  var getCol = rng.getColumn();
  var getURL = rng.getValue();
  
   
  
  
  if (getCol == 8 && isLister(rng.getSheet())) 
  {
            var mode="nzcu";
            
            var ssLive=SpreadsheetApp.openById(liveId);


            
            if(getURL.indexOf("walmart")>=0){importWMDescription(); return 0;}
            var ss = SpreadsheetApp.getActiveSpreadsheet();
            var sheet = ss.getSheetByName("Lister");
            var option = {
              'muteHttpExceptions' : true
            };
            var html = UrlFetchApp.fetch(getURL, option).getContentText();
            var htmlOrig=html;
            
            var n1=html.indexOf('<span itemprop="description"')
            n1=html.indexOf(">",n1)+1;
            
            var n2=html.indexOf('span>',n1)-2;
            //n2=html.lastIndexOf('<br>')+4
            
            var subHtml='<br>'+html.slice(n1,n2);
            Logger.log(subHtml);
            
            var s1=htmlOrig.indexOf('<table class="table table-dotted table-extended table-header translation-table">');
            var s2=htmlOrig.indexOf('</table>',s1)+8;
            
            
            var specs=htmlOrig.slice(s1,s2);
            
            var th=specs.indexOf('<thead>');+7;
            var th2=specs.indexOf('</thead>');
            
            var specs=specs.slice(0,th)+specs.slice(th2, specs.length)
            
            
            specs='<br><br><h2>Specification</h2><div>'+specs+'</div>';
            
            
              
            var allHtml='<h2>Details</h2><div>'+subHtml+'</div>'+specs
            
             try
            {
                var html = HtmlService.createHtmlOutput(allHtml)
                               .setTitle('Specifications')
                               .setWidth(300);
                SpreadsheetApp.getUi() // 
                        .showSidebar(html);
             }
             
             catch(err)
             {
                    var allHtml= "<div> This OS page has malformed coding. So it can't be diplayed here. Please visit the OS page";
                     var html = HtmlService.createHtmlOutput(allHtml)
                               .setTitle('Specifications')
                               .setWidth(300);
                      SpreadsheetApp.getUi() // 
                        .showSidebar(html);
             
             }
            
            
    }


}









function importWMDescription()
{
        var ss=SpreadsheetApp.getActiveSpreadsheet();
        var sheet=ss.getSheetByName("Lister");
        var rng=sheet.getActiveRange();
        
          var mode='nzcu';

        
        
        var url=rng.getValue();
        var getRow=rng.getRow();
        
        var mapValues=ss.getSheetByName("Mapping").getDataRange().getValues();
        
        var ss2=SpreadsheetApp.openById(liveId)
        var sheetLister=ss2.getSheetByName("Lister");
        
        
        
        
        var col=rng.getColumn();
        if(!(isLister(sheet))){return 0}
        if(col!=8){return 0}
        
        
        
        
        
        
        
        
        
        
        
        
        
         var option = {
                      'muteHttpExceptions' : true
          };

        var html = UrlFetchApp.fetch(url, option).getContentText();
        
        var n1=html.indexOf('WML_REDUX_INITIAL_STATE')+13+15;
        var n2=html.indexOf('</script>',n1)-3;
        
        
        var vFrm='=myvariation(R'+getRow+'C[0],R'+getRow+'C35:R'+getRow+'C50,R[0]C35:R[0]C50)'

        
        
        
        var html2=html.slice(n1,n2);

        var jsonData=JSON.parse(html2)
        
        if(jsonData==undefined){
              Browser.msgBox("Error Fetching Data");
              return 0;
        
        }
        var a=10
        
        
        
        
        
        var productBasicInfo=jsonData.productBasicInfo;
        var selectedProduct=productBasicInfo.selectedProductId;
        var selectedProdDetails=productBasicInfo[selectedProduct];
        
        
        var product=jsonData.product;
        var products=product.products;
       
   
            var primaryProductId=product.primaryProduct; //varaition map starts with base product
            var proimaryProduct=products[primaryProductId];
            
         var prodDesc=proimaryProduct.productAttributes.detailedDescription;
         var details=product.idmlMap[primaryProductId].modules
         var shortDesc=details.ShortDescription.product_short_description.values[0];
         
         prodDesc=shortDesc+"<br><br>"+prodDesc;
         var specs=details.Specifications.specifications.values[0]
         
         var htmlSpec="<h2>Specifications</h2><ul>";

         for (var i in specs)
         {
             var value=specs[i];
             
             
             for (var j in value)
             {
                 var info=value[j];
                 htmlSpec+='<li>'+info.displayName+": "+info.displayValue+'</li>'
             
             }
             
  
         }
         
         htmlSpec+="</ul>"
         
         var allHtml= '<h2>Details</h2><div>'+prodDesc+'</div>'+htmlSpec
       
       
         try
            {
                var html = HtmlService.createHtmlOutput(allHtml)
                               .setTitle('Specifications')
                               .setWidth(300);
                SpreadsheetApp.getUi() // 
                        .showSidebar(html);
             }
             
             catch(err)
             {
                    var allHtml= "<div> This OS page has malformed coding. So it can't be diplayed here. Please visit the WM page";
                     var html = HtmlService.createHtmlOutput(allHtml)
                               .setTitle('Specifications')
                               .setWidth(300);
                      SpreadsheetApp.getUi() // 
                        .showSidebar(html);
             
             }
       
       
       
       



}

