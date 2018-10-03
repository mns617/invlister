

//shows overstock images, 
//calls WM image show if url has walmart line 18

function imShowSideBar() {

    
        var ss=SpreadsheetApp.getActiveSpreadsheet();
        
        var mapValues=ss.getSheetByName("Mapping").getDataRange().getValues();
        var sheet=ss.getActiveSheet();
        var rng=sheet.getActiveRange();
        var row=rng.getRow();
        var col=rng.getColumn();
        
        if(!(isLister(sheet))){return 0}
        if(col!=8){return 0}        
        
                    var option = {
                      'muteHttpExceptions' : true
                    };
                    
                    
                    
                    var getURL = rng.getValue().toString();
                    if(getURL==""){return 0};
                    
                    
                    if(getURL.indexOf('walmart.com')>=0)
                    {
                            showWmImages();
                            return 0;
                            
                    
                    }
                    
                    
                    var headers = {                           
                      'ostkid': 'OSTK-VIP_18-A77359'                         
                    };                           
                    var option = {                            
                      "headers": headers,
                      'muteHttpExceptions' : true                           
                    };
                    
                    
                    
                    
                    
                    
                    
                    
                    var html = UrlFetchApp.fetch(getURL, option).getContentText();
                    var htmlOrig=html;
                    
                    
                    var n1=html.indexOf('s-h-title');
                    var n2=html.indexOf("<",n1);
                    var title=html.slice(n1+11,n2-1); 
                    
                //    var folderId = "0Bw-TXeLyArDnLTBabkVxUXBoeVk";
                    
                  //  var tempFolder=DriveApp.getFolderById(folderId).createFolder(Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd-MMMM-yyyy hh:mm a")+"--"+title);; 
                   // var folderUrl=tempFolder.getUrl();
                    
                    var n1=html.lastIndexOf('<div class="container">');
                    var n2=html.indexOf('</ul>',n1)
                    var html2=html.slice(n1,n2)
                    var sbHtml='<br>';
                    var imgUrlArr=html2.split('data-max-img');
                    
                        for (var j=1; j<imgUrlArr.length; j++)  //when there is variation, index 0 has garbage data
                        {
                           var longUrl=imgUrlArr[j];
                              var l1=longUrl.indexOf("ak1");
                              var l2=longUrl.indexOf(">",l1);
                              var imUrl=longUrl.slice(l1,l2-1);
                              
                              //var imageURL=(imUrl).replace("ostkcdn.com","ostkcdn.com.rsz.io")+"?flip=x"
                              
                              var imageURLFlipped="http://res.cloudinary.com/demo/image/fetch/a_hflip/http://"+imUrl;
                              var imageURLCropped="http://res.cloudinary.com/demo/image/fetch/a_hflip,h_0.95,w_0.999,c_crop,g_north_west/http://"+imUrl;
                              var imageURL="http://res.cloudinary.com/demo/image/fetch/http://"+imUrl;
    
                              //sheet.getRange(sheet.getActiveRange().getRow(), 20).setValue(imageURL); 
    
                              sbHtml=sbHtml+'<img src="'+imageURL+'" alt="Mountain View" style="width:auto; height:200px;"><br><br>'                          //sbHtml=sbHtml+'<img src="'+imageURLFlipped+'" alt="Mountain View" style="width:100px;height:150px;"><br><br><br>'
    
                              +'<form>'
                                +'Flipped url: <input type="text" name="fname" value="'+imageURLFlipped+'"><br>'
                                +'Cropped url: <input type="text" name="fname" value="'+imageURLCropped+'"><br>'
                                +'Regular url: <input type="text" name="fname" value="'+imageURL+'"><br>'
    
                               +'</form>'
                               +'<br><hr><br>'
                              
                              
                              
                              
                              
                              
                              
                              var imageURLFlipped="http://res.cloudinary.com/demo/image/fetch/a_hflip/http://"+imUrl;

                          //var imBlob=UrlFetchApp.fetch(imageURL).getBlob();
                          
                          //var imFile=tempFolder.createFile(imBlob);
                          //imFile.setName(imPhrase+" "+ j+".jpg");
                        }
                        
                        
                        
                        
                        
                        
                  if(sbHtml!='<br>')
                  {
                      var imhtml = HtmlService.createHtmlOutput(sbHtml)
                          .setTitle('Images')
                          .setWidth(300);
                          SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
                          .showSidebar(imhtml);
                    
                   }     
                        
                        
                        
                        
                        
                        
                        
                        

  }
  
  
  
  
  
  
  
  
 
/*


function showWmImages()
{
      
      var ss=SpreadsheetApp.getActiveSpreadsheet();
        
        var mapValues=ss.getSheetByName("Mapping").getDataRange().getValues();
        var sheet=ss.getActiveSheet();
        var rng=sheet.getActiveRange();
        var row=rng.getRow();
        var col=rng.getColumn();
        
        if(sheet.getName()!="Lister"){return 0}
        if(col!=8){return 0}        
        
                    var option = {
                      'muteHttpExceptions' : true
                    };
                    
                    
                    
                    var getURL = rng.getValue().toString();
                    if(getURL==""){return 0};
                    
                    
                    var html = UrlFetchApp.fetch(getURL, option).getContentText();
                    var htmlOrig=html;
                    
                   
                   var n1=html.indexOf('WML_REDUX_INITIAL_STATE')+13+15;
        var n2=html.indexOf('</script>',n1)-3;
        
        
        
        
        var html2=html.slice(n1,n2);
        //GmailApp.sendEmail('sakib118.biz@gmail.com', "a", html2)
        //var n1=html2.lastIndexOf(' ');
        
        
      
      //  var test=html.slice(172383-2, )
        //GmailApp.sendEmail('sakib118.biz@gmail.com', 'test', html.slice(n1,n2))
        var jsonData=JSON.parse(html2);
        
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
            
                
                
                var allOffers=jsonData.product.offers;
                
                //push all images to an array using 'variantCategoriesMap'
                var varMap=product.variantCategoriesMap;
                
                var prodsForImages=[];
                var allImageIds=[];
                
        
                    var primaryProduct=product.primaryProduct; //varaition map starts with base product
                    var varMap=product.variantCategoriesMap[primaryProduct]; // first property is the primay product
                    
                    var colorVars=varMap.actual_color.variants;
                    var allImages=[];
                    var relatedColors=[];
                    var relatedSizes=[];
                    
                        for (var j in colorVars)
                        {
                                var tempColorVar=colorVars[j];
                                var tempProducts=tempColorVar.products;
                                var tempImages=tempColorVar.images;
                                
                                for (var k=0; k<tempImages.length; k++)
                                {
                                            if(allImages.indexOf(tempImages[k])<0)
                                            {
                                                    allImages.push(tempImages[k]);
                                                    relatedColors.push(tempColorVar.name);
                                                    relatedSizes.push("");  //put an empty for zie
                                                    
                                            
                                            }
                                            
                                            else    /// image was pushed before
                                            {
                                                    var index=allImages.indexOf(tempImages[k])
                                                    relatedColors[index]=tempColorVar.name;
                                                    relatedSizes[index]="";
                                            
                                            
                                            }
                                      
                                
                                
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
                            
                            
                             var imageURLFlipped="http://res.cloudinary.com/demo/image/fetch/a_hflip/http://"+imUrl;
                             var imageURL="http://res.cloudinary.com/demo/image/fetch/http://"+imUrl;



                          var tempHtml='<br>Color: '+color+'   '+imType+'<img src="'+imageURL+'" alt="Mountain View" style="width:auto; height:200px;"><br><br>'                          //sbHtml=sbHtml+'<img src="'+imageURLFlipped+'" alt="Mountain View" style="width:100px;height:150px;"><br><br><br>'

                          +'<form>'
                            +'Flipped url: <input type="text" name="fname" value="'+imageURLFlipped+'"><br>'
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
  
*/ 
  
  
  
  
  

