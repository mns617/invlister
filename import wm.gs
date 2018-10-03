// this function imports walmart data. if no varaition is found then calls the no variation function at 132

function importWMdata(rng)
{
        var ss=SpreadsheetApp.getActiveSpreadsheet();
        var sheet=ss.getActiveSheet();
        var rng=sheet.getActiveRange();
        
          var mode='nzcu';

        
        
        var url=rng.getValue();
        if(url.indexOf('?')>0)
        {
          url=url.slice(0, url.indexOf('?'));
        }
        
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
        
       
        var vFrm='=myvariation(R'+getRow+'C[0],R'+getRow+'C35:R'+getRow+'C50,R[0]C35:R[0]C50)'


        var jsonData=getMyJson(html);
        
        if(jsonData==undefined){
              Browser.msgBox("Error Fetching Data");
              return 0;
        
        }
        var a=10
        
        
        var prodId=jsonData.productId;
        var prodName="";
        
        
        
        var productBasicInfo=jsonData.productBasicInfo;
        var selectedProduct=productBasicInfo.selectedProductId;
        var selectedProdDetails=productBasicInfo[selectedProduct];
        var wmTitle=selectedProdDetails.title;
        var itemNo=selectedProdDetails.usItemId;
        
        var initial="";
        var date=Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MM/dd/yyyy");
        var asin="";
        var sku="";
        var comment=allLenFrm//"=LEN(R[0]C[2])"// a formula that mesuares length of the Amazon Title//sheet.getRange("J515").getFormulaR1C1();
        
        var product=jsonData.product;
        var products=product.products;
        var upc="";
        var status="";
        var amTitle="";
        var type="BedAndBath";
        var category="";
        
        
        var b2="";
        var b3="";
        var b4="";
        var terms="";
        var imageMain="";
        var color="";
        var size="";
        var material="";

        
        
        
        
       /* 
         var matchedRow3=myLookup(productTitle, mapValues, 7); //match the category based on title
                
         if(matchedRow3!=null)
         {
           category=mapValues[matchedRow3-1][8-1];
           
         }
           */   
                
        
        var allOffers=jsonData.product.offers;
        
        //push all images to an array using 'variantCategoriesMap'
       
        var prodsForImages=[];
        var allImageIds=[];
        

            var primaryProduct=product.primaryProduct; //varaition map starts with base product
            var varMap=product.variantCategoriesMap[primaryProduct]; // first property is the primay product
            
            if(varMap==undefined)
            {
                  //no variation item
                  importWmNoVariation_(rng,jsonData);
                    
                  return 0
            }
            
            
            
               
            
            
            var flag1=0;
            var flag2=0;
            // these two arrays will all variation information
            var cv=varMap.actual_color;
            if(cv!=undefined)
              {var colorVars=cv.variants;}
            else
              {flag1=1;}
              
              
             var sv= varMap.size;
             if(sv!=undefined) 
              {var sizeVars=sv.variants;}
             else
             {flag2=1;}
            
                  
        
        
        
        var arr=[];
        //var amPrice='=ROUNDUP(R[0]C[17]-(-R[0]C[16]-((R[0]C[16]*0.06))+(R[0]C[16]*0))/0.85)-0.01';
        var rowTemp=getRow;
        var dataArr=[];
        
        var numVar=0;
        var currentValues=sheet.getRange(getRow, 1, 1000,1).getValues();// imagine maximum 1000 variation

        
        
        
        
        
        
        for(var i in products)
        {
                      //--check if overwrriting the exististing data
                      if(currentValues[numVar][0]!="")
                      {
                        
                        Browser.msgBox("This will overwrite data in row "+rowTemp);
                        return 0;  
                        
                      }
                     //---------------------
                   
                    var rowArr=[];
                    var dProd=products[i]; ///daughter product i.e. this product
                    //var itemNo=dProd.usItemId; // commented because using selected item id
                    
                    var variantsProp=dProd.variants; //variants of this product
                   
                     var count=0;
                     var variation="";
                     //get the variant details
                     if(flag1==0 && flag2==0)
                     {
                             var sizeProp=variantsProp.size;
                             var sizeName=sizeVars[sizeProp].name;
                             
                             var colorProp=variantsProp.actual_color;
                             var colorName=colorVars[colorProp].name;
                             
                             var skugridVar=sizeName+'|'+colorName;
                      }
                      
                      else if  (flag1==0)  //only color vari
                      {
                             var colorProp=variantsProp.actual_color;
                             var colorName=colorVars[colorProp].name;
                             
                             var skugridVar=colorName;
                      }
                      
                      
                      else if  (flag2==0)  //only color vari
                      {
                             var sizeProp=variantsProp.size;
                             var sizeName=sizeVars[sizeProp].name;
                             var skugridVar=sizeName;

                      }
                      
                      
                      
                      var color="";
                      var size="";
                      
                      var thisItem=dProd.usItemId;
                      var sellerUrl= 'https://www.walmart.com/product/'+thisItem+'/sellers'; // imports the list of sellers
                      
                        var option = {
                              'muteHttpExceptions' : true
                        };

                     var htmlSeller = UrlFetchApp.fetch(sellerUrl, option).getContentText();
                     var jsonDataSeller= getMySellerJson(htmlSeller);
          
                     var selectedProd=jsonDataSeller.product.selected.product;
                    
                     var detailsOfSelectedProd=jsonDataSeller.product.products[selectedProd];
                     var flagWm=0;
                     if(detailsOfSelectedProd!=undefined)
                      {
                        var myOffers=jsonDataSeller.product.products[selectedProd].offers
                        var allOffers=jsonDataSeller.product.offers;
                      }
                      
                     else
                      {
                          var flagWm=1;
                      }
                      
                      

                      
                      
                     
                      //var k=-1  //stop loop for texting
                      for(var k=0; k<myOffers.length && flagWm==0; k++)
                      {
                               var tempOfferId=myOffers[k];
                               var tempOffer=allOffers[tempOfferId];
                               var isStock=tempOffer.productAvailability.availabilityStatus;
                               
                               
                               
                                   
                              
                                     var sellerId=tempOffer.sellerId; //when there is more than one offer                                   
                                     if(sellerId=='F55CDC31AB754BB68FE0B39041159D63')
                                     {
                                         flagWm=1;  // break the loop by setting flag
                                         var currentPrice=tempOffer.pricesInfo.priceMap.CURRENT.price
                               
                                         if(currentPrice<35){currentPrice+=5};
                                         var profit=sheetLister.getRange("U5").getFormulaR1C1();
                                         //sheet.getRange("O518").getFormulaR1C1();
                                         break;
                                         
                                     }

                      
                      }  // emd of offers for
                      
                      if(flagWm==0)
                      {
                          currentPrice="NOT WALMART";
                          profit="0";
                          amPrice="NOT WALMART";
                      
                      }
 
                   //--------------------------//
                      
                      

                      var lenFrm="=LEN(L"+rowTemp+")"
                      var lenFrm2="=LEN(S"+rowTemp+")"
                      var lenFrm3="=LEN(R"+rowTemp+")"
                      var imFrm="=T"+rowTemp;
                      
                      
                      if(mode=='test' & numVar>0 )  // from test mode and variation is more than 1
                      {
                              amTitle=vFrm;
                              type=vFrm;
                              category=vFrm;
                              b2=vFrm;
                              b3=vFrm;
                              b4=vFrm;
                              terms=vFrm;
                              color=vFrm;
                              size=vFrm;
                      
                      
                      
                      }
                      
                      if(isStock=='IN_STOCK')
                      {
                            isStock=1;
                      
                      }

                      else
                      {
                          isStock=0;
                      }
                      

                      var lenFrm="=LEN(L"+rowTemp+")"
                      var lenFrm2="=LEN(S"+rowTemp+")"
                      var lenFrm3="=LEN(R"+rowTemp+")"
                      var imFrm="=T"+rowTemp;
                      
                      
                      if(mode=='nzcu' & numVar>0 )  // from nzcu mode and variation is more than 1
                      {
                              amTitle=vFrm;
                              type=vFrm;
                              category=vFrm;
                              b2=vFrm;
                              b3=vFrm;
                              b4=vFrm;
                              terms=vFrm;
                              material=vFrm;
                              size=vFrm;
                      
                      
                      
                      }
                      
                      if(isStock=='IN_STOCK')
                      {
                            isStock=1;
                      
                      }

                      else
                      {
                          isStock=0;
                      }
                      
                      
                      
                      
                      
                      
                      var tempArr=[wmTitle, initial, date, asin, sku, itemNo, skugridVar, url, upc, comment, status, amTitle, type, category, amPrice1, b2,b3,b4,terms,imageMain, color,size,material, imFrm, imFrm, imFrm,lenFrm3,lenFrm2,lenFrm,isStock, currentPrice, profit]
                      dataArr.push(tempArr);
                      var a=10;
                      rowTemp++
                
                      numVar++
              
              
   
        
        }//end of products for
        
              

        
        
        sheet.getRange(getRow, 1, dataArr.length, dataArr[0].length).setValues(dataArr);
        deactivateFormulas(sheet.getRange(getRow, 1, dataArr.length, dataArr[0].length));
      
      
       return dataArr.length;
        
        




}

















function importWmNoVariation_(rng,jsonData)
{

        var ss=SpreadsheetApp.getActiveSpreadsheet();
        var sheet=ss.getActiveSheet();
        var rng=sheet.getActiveRange();
        
        var mode='nzcu';

        
        
        var url=rng.getValue();
        
                
        var url=rng.getValue();
        if(url.indexOf('?')>0)
        {
          url=url.slice(0, getURL.indexOf('?'));
        }
        
        var getRow=rng.getRow();
        
        var mapValues=ss.getSheetByName("Mapping").getDataRange().getValues();
        
        var ss2=SpreadsheetApp.openById(liveId)
        var sheetLister=ss2.getSheetByName("Lister");
        
        
        
         var option = {
                      'muteHttpExceptions' : true
          };

        var html = UrlFetchApp.fetch(url, option).getContentText();
        
        
        

        
        var vFrm='=myvariation(R'+getRow+'C[0],R'+getRow+'C35:R'+getRow+'C50,R[0]C35:R[0]C50)'

        var jsonData=getMyJson(html);
        
        
        if(jsonData==undefined){
              Browser.msgBox("Error Fetching Data");
              return 0;
        
        }
        var a=10
        
        
        var prodId=jsonData.productId;
        var prodName="";
        
        
        
        var productBasicInfo=jsonData.productBasicInfo;
        var selectedProduct=productBasicInfo.selectedProductId;
        var selectedProdDetails=productBasicInfo[selectedProduct];
        var wmTitle=selectedProdDetails.title;
        
        
        var initial="";
        var date="" //Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MM/dd/yyyy");
        var asin="";
        var sku="";
        var comment=allLenFrm//"=LEN(R[0]C[2])"// a formula that mesuares length of the Amazon Title//sheet.getRange("J515").getFormulaR1C1();
        
        var product=jsonData.product;
        var products=product.products;
        var upc="";
        var status="";
        var amTitle="";
        var type="BedAndBath";
        var category="";
        
        
        var b2="";
        var b3="";
        var b4="";
        var terms="";
        var imageMain="";
        var color="";
        var size="";
        var material="";
        var mTerms="";
        
        
        
        
        
       /* 
         var matchedRow3=myLookup(productTitle, mapValues, 7); //match the category based on title
                
         if(matchedRow3!=null)
         {
           category=mapValues[matchedRow3-1][8-1];
           
         }
           */   
                
        
        var allOffers=jsonData.product.offers;
        
        //push all images to an array using 'variantCategoriesMap'
       
        var prodsForImages=[];
        var allImageIds=[];
        

        var primaryProduct=product.primaryProduct; //varaition map starts with base product
            
                  
        
        
        
        var arr=[];
       
        var rowTemp=getRow;
        var dataArr=[];
        
        var numVar=0;

        
        
        
        
        
        
        for(var i in products)
        {
                      
                     //---------------------
                   
                    var rowArr=[];
                    var dProd=products[i]; ///daughter product i.e. this product
                    var itemNo=dProd.usItemId;
                    
                    var variantsProp=dProd.variants; //variants of this product
                   
                    var count=0;
                    var variation="";
                     
                       var matchedRow=myLookup(variation, mapValues, 1);//map color vased on variation
                       var color="";                                                
                       if(matchedRow!=null)
                       {
                         color=mapValues[matchedRow-1][2-1];
                         
                       }
                       
                       
                        var matchedRow2=myLookup(variation, mapValues, 4)
                        var size="";
                        if(matchedRow2!=null)
                        {
                          size=mapValues[matchedRow2-1][5-1];
                          
                        }
   
       
                   //--------------------------////----------------------//
                      var thisItem=dProd.usItemId;
                      var sellerUrl= 'https://www.walmart.com/product/'+thisItem+'/sellers'; // imports the list of sellers
                      
                        var option = {
                              'muteHttpExceptions' : true
                        };

                     var htmlSeller = UrlFetchApp.fetch(sellerUrl, option).getContentText();
                     var jsonDataSeller= getMySellerJson(htmlSeller);
                     var selectedProd=jsonDataSeller.product.selected.product;
                    
                     var detailsOfSelectedProd=jsonDataSeller.product.products[selectedProd];
                     var flagWm=0;
                     if(detailsOfSelectedProd!=undefined)
                      {
                        var myOffers=jsonDataSeller.product.products[selectedProd].offers
                        var allOffers=jsonDataSeller.product.offers;
                      }
                      
                     else
                      {
                          var flagWm=1;
                      }
                      
                                  
                     
                      //var k=-1  //stop loop for texting
                      for(var k=0; k<myOffers.length && flagWm==0; k++)
                      {
                               var tempOfferId=myOffers[k];
                               var tempOffer=allOffers[tempOfferId];
                               var isStock=tempOffer.productAvailability.availabilityStatus;
                               
                               
                               
                                   
                              
                                     var sellerId=tempOffer.sellerId; //when there is more than one offer                                   
                                     if(sellerId=='F55CDC31AB754BB68FE0B39041159D63')
                                     {
                                         flagWm=1;  // break the loop by setting flag
                                         var currentPrice=tempOffer.pricesInfo.priceMap.CURRENT.price
                               
                                         if(currentPrice<35){currentPrice+=5};
                                         var profit=sheetLister.getRange("U5").getFormulaR1C1();
                                         //sheet.getRange("O518").getFormulaR1C1();
                                         break;
                                         
                                     }

                      
                      }  // emd of offers for
                      
                      if(flagWm==0)
                      {
                          currentPrice="NOT WALMART";
                          profit="0";
                          amPrice1="NOT WALMART";
                      
                      }
 
           
                      
                      if(mode=='nzcu' & numVar>0 )  // from nzcu mode and variation is more than 1
                      {
                              amTitle=vFrm;
                              type=vFrm;
                              category=vFrm;
                              b2=vFrm;
                              b3=vFrm;
                              b4=vFrm;
                              terms=vFrm;
                              color=vFrm;
                              size=vFrm;
                              mTerms=vFrm;
                      
                      
                      
                      }
                      
                      if(isStock=='IN_STOCK')
                      {
                            isStock=1;
                      
                      }

                      else
                      {
                          isStock=0;
                      }
                      
                      
                      
                      var imFrm="=T"+getRow;
                      var lenFrm="=LEN(L"+getRow+")";
                      var skugridVar="";
                      
                      
                      
                      
                      
                      var lenFrm="=LEN(L"+rowTemp+")"
                      var lenFrm2="=LEN(S"+rowTemp+")"
                      var lenFrm3="=LEN(R"+rowTemp+")"
                      var imFrm="=T"+rowTemp;
                      
                      var tempArr=[wmTitle, initial, date, asin, sku, itemNo, skugridVar, url, upc, comment, status, amTitle, type, category, amPrice1, b2,b3,b4,terms,imageMain, color,size,material, imFrm, imFrm, imFrm,lenFrm3,lenFrm2,lenFrm,isStock, currentPrice, profit]
                      dataArr.push(tempArr);
                      var a=10;
                      rowTemp++
                
                      numVar++
              
              
   
        
        }

        sheet.getRange(getRow, 1, dataArr.length, dataArr[0].length).setValues(dataArr);
        return 0;
  









}






