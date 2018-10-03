




function showWmImages()
{
      
      var ss=SpreadsheetApp.getActiveSpreadsheet();
        
        var mapValues=ss.getSheetByName("Mapping").getDataRange().getValues();
        var sheet=ss.getActiveSheet();
        var rng=sheet.getActiveRange();
        var row=rng.getRow();
        var col=rng.getColumn();
        
        if(isLister(sheet)==false){return 0}
       
        
                    var option = {
                      'muteHttpExceptions' : true
                    };
                    
                    
                    
                    var getURL = rng.getValue().toString();
                    if(getURL==""){return 0};
                    
                    
                    var html = UrlFetchApp.fetch(getURL, option).getContentText();
                    var htmlOrig=html;
                    
   
   
   
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
        var date= Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MM/dd/yyyy");
        var asin="";
        var sku="";
        var comment="=LEN(R[0]C[2])"// a formula that mesuares length of the Amazon Title//sheet.getRange("J515").getFormulaR1C1();
        
        
        var product=jsonData.product;
        var products=product.products;
        var upc="";
        var status="";
        var amTitile="";
        var type="BedAndBath";
        var category="";
        
       /* 
         var matchedRow3=myLookup(productTitle, mapValues, 7); //match the category based on title
                
         if(matchedRow3!=null)
         {
           category=mapValues[matchedRow3-1][8-1];
           
         }
           */   
                
                
                var allOffers=jsonData.product.offers;
                
                //push all images to an array using 'variantCategoriesMap'
                var varMap=product.variantCategoriesMap;
                
                var prodsForImages=[];
                var allImageIds=[];
                
        
                    var primaryProduct=product.primaryProduct; //varaition map starts with base product
                    var varMap=product.variantCategoriesMap[primaryProduct]; // first property is the primay product
                    
                    
                    
                    if(varMap==undefined)
                    {
                          //no variation item
                          var firstImUrl=showWmImagesNoVar_(rng,jsonData);
                         
                          return firstImUrl;
                    }
                    
                    
                     
                    var allImages=[];
                    var relatedColors=[];
                    var relatedSizes=[];
                    
                    
                    
                    
                    for (var p in products)
                    {
                          var thisProduct=products[p];
                          var thisProdVariants=thisProduct.variants;
                          var size=thisProdVariants.size;

                          if(size==undefined){size="N/A"}
                          else
                          {
                              size=replaceAll(size, "size-", "");
                          }
                          

                          var color=thisProdVariants.actual_color;
                          if(color==undefined){color="N/A"}
                          else color=color.split("-")[1]
                          
                          var thisProdImages=thisProduct.images;
                          if(thisProdImages==undefined){continue}

                          
                          for (var im=0; im<thisProdImages.length; im++)
                          {
                               allImages.push(thisProdImages[im]);
                               relatedColors.push(color);
                               relatedSizes.push(size);  //put an empty for zie
                          
                          }
                          
                    
                    }
                    
                    
                 



                    var sbHtml='<br>';
                    var htmlArr=[];
                    
                    var images=product.images
                    for (var i in images) 
                    {
                            var image=images[i];
                            var imId=image.assetId;
                            var imType= image.type
                            imType=('('+imType+')').toLowerCase();
                            var imUrl=image.assetSizeUrls.main;
                            var im1=imUrl.indexOf('i5.walmartimages.com');
                            var im2=imUrl.indexOf('.jpeg')
                            imUrl=imUrl.slice(im1,im2+5);
                            
                            var index=allImages.indexOf(imId);
                            var color=relatedColors[index];
                            var size=relatedSizes[index];
                            color=color+"|"+size
                             var imageURLFlipped="http://res.cloudinary.com/demo/image/fetch/a_hflip/http://"+imUrl;
                             var imageURLCropped="http://res.cloudinary.com/demo/image/fetch/a_hflip,h_0.95,w_0.999,c_crop,g_north_west/http://"+imUrl;
                             var imageURL="http://res.cloudinary.com/demo/image/fetch/http://"+imUrl;



                          var tempHtml='<br>Color: '+color+'   '+imType+'<img src="'+imageURL+'" alt="Mountain View" style="width:auto; height:200px;"><br><br>'                          //sbHtml=sbHtml+'<img src="'+imageURLFlipped+'" alt="Mountain View" style="width:100px;height:150px;"><br><br><br>'

                          +'<form>'
                            +'Flipped url: <input type="text" name="fname" value="'+imageURLFlipped+'"><br>'
                            +'Cropped url: <input type="text" name="fname" value="'+imageURLCropped+'"><br>'
                            +'Regular url: <input type="text" name="fname" value="'+imageURL+'"><br>'

                           +'</form>'
                           +'<br><hr>'
                          
                          htmlArr.push(tempHtml);
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                    
                    }
        
        
        
        
                   htmlArr=htmlArr.sort();
        
                   sbHtml=sbHtml+htmlArr.join('<br>');
        
        
                     var a=10
                   
                   
                   
                   
             
                        
                        
                  if(sbHtml!='<br>')
                  {
                      var imhtml = HtmlService.createHtmlOutput(sbHtml)
                          .setTitle('Images')
                          .setWidth(300);
                          SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
                          .showSidebar(imhtml);
                    
                   }













}

























// called by showWmimages function

function showWmImagesNoVar_(rng, jsonData)
{
      
      var ss=SpreadsheetApp.getActiveSpreadsheet();
        
        var mapValues=ss.getSheetByName("Mapping").getDataRange().getValues();
        var sheet=ss.getActiveSheet();
       var row=rng.getRow();
        var col=rng.getColumn();
        
       // if(sheet.getName()!="Lister"){return 0}
       // if(col!=8){return 0}        
        
        
        var prodId=jsonData.productId;
        var prodName="";
        
        
        
        var productBasicInfo=jsonData.productBasicInfo;
        var selectedProduct=productBasicInfo.selectedProductId;
        var selectedProdDetails=productBasicInfo[selectedProduct];
        var wmTitle=selectedProdDetails.title;
        
        
        var initial="";
        var date= Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MM/dd/yyyy");
        var asin="";
        var sku="";
        var comment="=LEN(R[0]C[2])"// a formula that mesuares length of the Amazon Title//sheet.getRange("J515").getFormulaR1C1();
        
        
        var product=jsonData.product;
        var products=product.products;
        var upc="";
        var status="";
        var amTitile="";
        var type="BedAndBath";
        var category="";
        
       /* 
         var matchedRow3=myLookup(productTitle, mapValues, 7); //match the category based on title
                
         if(matchedRow3!=null)
         {
           category=mapValues[matchedRow3-1][8-1];
           
         }
           */   
                
                
                var allOffers=jsonData.product.offers;
                
                //push all images to an array using 'variantCategoriesMap'
                var varMap=product.variantCategoriesMap;
                
                var prodsForImages=[];
                var allImageIds=[];
                
        
                    var primaryProduct=product.primaryProduct; //varaition map starts with base product
                    var varMap=product.variantCategoriesMap[primaryProduct]; // first property is the primay product
                    



                    var sbHtml='<br>';
                    var htmlArr=[];
                    
                    var firstImage="";
                    
                    var images=product.images
                    for (var i in images) 
                    {
                            var image=images[i];
                            var imId=image.assetId;
                            var imType= image.type
                            imType=('('+imType+')').toLowerCase();
                            var imUrl=image.assetSizeUrls.main;
                            var im1=imUrl.indexOf('i5.walmartimages.com');
                            var im2=imUrl.indexOf('.jpeg')
                            imUrl=imUrl.slice(im1,im2+5);
                            
                           // var index=allImages.indexOf(imId);
                           // var color=relatedColors[index];
                            
                            
                             var imageURLFlipped="http://res.cloudinary.com/demo/image/fetch/a_hflip/http://"+imUrl;
                             var imageURLCropped="http://res.cloudinary.com/demo/image/fetch/a_hflip,h_0.95,w_0.999,c_crop,g_north_west/http://"+imUrl;
                             var imageURL="http://res.cloudinary.com/demo/image/fetch/http://"+imUrl;
                           if(imType=="(primary)")
                             {
                             
                                 
                                 //sheet.getRange(row, 20).setValue(imageURL);
                                 //return imageURL;
                             
                             }
                          var color='N/A' 
                          var tempHtml='<br>Color: '+color+'   '+imType+'<img src="'+imageURL+'" alt="Mountain View" style="width:auto; height:200px;"><br><br>'                          //sbHtml=sbHtml+'<img src="'+imageURLFlipped+'" alt="Mountain View" style="width:100px;height:150px;"><br><br><br>'

                          +'<form>'
                            +'Flipped url: <input type="text" name="fname" value="'+imageURLFlipped+'"><br>'
                            +'Cropped url: <input type="text" name="fname" value="'+imageURLCropped+'"><br>'
                            +'Regular url: <input type="text" name="fname" value="'+imageURL+'"><br>'

                           +'</form>'
                           +'<br><hr>'
                          
                          htmlArr.push(tempHtml);
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                    
                    }
        
        
        
        
        
                   sbHtml=sbHtml+htmlArr.join('<br>');
        
        
                   
                   
                   
             
                        
                        
                  if(sbHtml!='<br>')
                  {
                      var imhtml = HtmlService.createHtmlOutput(sbHtml)
                          .setTitle('Images')
                          .setWidth(300);
                          SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
                          .showSidebar(imhtml);
                    
                   }


                   return firstImage;










}


