
function getOverstockData(rng) 
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


            
            if(getURL.indexOf("overstock")<0){return 0;}
            var ss = SpreadsheetApp.getActiveSpreadsheet();
            var sheet = ss.getActiveSheet();
            var headers = {                           
              'ostkid': 'OSTK-VIP_18-A77359'                         
            };                           
            var option = {                            
              "headers": headers,
              'muteHttpExceptions' : true                           
            };
            
            var html = UrlFetchApp.fetch(getURL, option).getContentText();
            var htmlOrig=html;
            
            
            getURL=getURL.slice(0, getURL.indexOf('.html')+5);
            var mapValues=ss.getSheetByName("Mapping").getDataRange().getValues();
            
//    Logger.log(html);
            var arrVals = [];
            var arrVals2=[];
            var arrVals3=[]
            var m1 = html.indexOf("product-title");
            var m2 = html.indexOf("<h1>", m1);
            var m3 = html.indexOf("</h1>", m2);
            var title = html.slice(m2+4, m3);
            
            //get rid of special chars
                                                                        
            title = replaceUnwantedFromOs(title).trim();
            
            if(title.toLowerCase().indexOf('sweet jojo')>=0)
            {
                  Browser.msgBox("Sweet Jojo is a prohibited brand please skip this ad!")
                  return 0;
            
            }
            
            
            if(html.indexOf('Overstock Marketplace Seller')>0)
            {
                  Browser.msgBox("We cannot sell product of third party seller, please skip this ad!")
                  return 0;      
            }
            
            
            if(title.toLowerCase().indexOf('as is item')>=0)
            {
                   Browser.msgBox("Do not list As is Items!")
                  return 0;
            
            }
            
            var i1 = html.indexOf("item-number");
            var i2 = html.indexOf(">", i1);
            var i3 = html.indexOf("<", i2);
            var itemTxt = html.slice(i2+7, i3);
            var i4 = itemTxt.indexOf("ITEM#");
            var itemNo = itemTxt;
            itemNo=removeChars('0123456789', itemNo);
            
        
            
            var today = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MM/dd/yyyy");
            
            
            
            var amPrice="";
            var profit="";
            var isSale=isOnSale(html);
            var sheetLister=ss.getSheetByName("Lister");
            
            if(getURL.indexOf("Sports-Toys")>0)
            {
                isSale="Yes";
            }
                        
            
            if(isSale=="No")
            {
                profit=sheetLister.getRange("U2").getFormulaR1C1();
                amPrice= "=ROUND((R[0]C[17]-(-(R[0]C[16])+((R[0]C[16]*0.12))+((R[0]C[16])-(R[0]C[16]*0.12))*0.0688))/0.85,0)-0.01";
                }
    
            else
             {
                profit=sheetLister.getRange("U3").getFormulaR1C1();
                amPrice="=ROUND(((R[0]C[17]-(-(R[0]C[16])+(R[0]C[16]*0.1188)))/0.85),0)-0.01";

                }
                
                
            var initials=""; 
            var comment=allLenFrm;
            var variation="";
            var status="";
            
            
            
            
        
            
            
            
            
            
            
            
            
            if (mode=="nzcu")//nazmus is going to use this version for drafting
            {
                  var rowWhere='=IFERROR(MATCH(F'+getRow+',F2:F,0))';
                  var lenFrm='=LEN(TRIM(L'+getRow+'))';
                  initials="";
                  comment=allLenFrm;
            
            }
            
            
            
            
            
            
            
            var t1=html.indexOf('dropDownOptions');
            var htmlOption=html.slice(t1)
          
            
            //find out if three is a variation
            var n1=html.indexOf("options-dropdown");
            var n2=html.indexOf("</select>",n1);
            var prodOptions=html.slice(n1,n2);
            
            
            var asin="";
            var type='BedAndBath';
            var catagory="";
            
            
            var b2="";
            var b3="";
            var b4="";
            var terms="";
            var imageMain="";
            var color="";
            var size="";
            var material="";
            var b1="";
          
            
            
            
            if(n1==-1)  //when there is no variation
            {

                                    
                                    //upc portion
                                    
                                      
                                      
                                      var upc="";//sheetUPC.getRange(lrUPC, 1).getValue();
                                      //sheet.getRange(getRow, 9).setValue(upc);
                              
                                      var prodName="";//amazon title
                                      var sellerSku=""; "DH"+upc;
                                      //sheet.getRange(getRow, 5).setValue(sellerSku);
                                      var imFrm="=T"+getRow;
                                      
                                      
                                      //sheet.getRange(getRow, 24,1,3).setValues([[frmImg, frmImg, frmImg]])
                      
                                    
                                    var cell="I"+getRow;
                                    var prc1=html.indexOf('monetary-price-value');
                                    var prc2=html.indexOf("content=",prc1);
                                    var prc3=html.indexOf(">",prc2);
                                    var price=html.slice(prc2+9,prc3-1);
                                    
                                    
                                    var lenFrm="=LEN(L"+getRow+")"
                                    
                                    
                                    
                                    //get quantity
                                    
                                            var q1=html.indexOf('dropDownOptions');
                                            var q2=html.indexOf('fromOptionBasedRequest',q1);
                                            var qtyHtml=html.slice(q1,q2);
                                            var qtyArr=qtyHtml.split('maxQuantity');
                                    
                                            var longQty=qtyArr[1];
                                            var q1=longQty.indexOf(":");
                                            var q2=longQty.indexOf(",",q1);
                                            var qty=longQty.slice(q1+1,q2);
                                    
                                    
                                                 
                                                            
                                            var matchedRow3=myLookup(title, mapValues, 7)
                                            
                                            if(matchedRow3!=null)
                                            {
                                              catagory=mapValues[matchedRow3-1][8-1];
                                              
                                            }
                                    
                                                
                                    
                                    
                                    
                                    
                                    
                                    
                                    arrVals.push([title, initials, today, asin, sellerSku, itemNo, variation, getURL, upc, comment,status, prodName,type, catagory,amPrice,b2,b3,b4, terms, imageMain, color,size,material, imFrm, imFrm, imFrm,"","",lenFrm,qty, price, profit]);
                                    sheet.getRange(getRow, 1, arrVals.length, arrVals[0].length).setValues(arrVals);

                    
                                 //    sheetUPC.deleteRow(lrUPC);


            
            }
    
            
            else if(n1>=0)//when there is vairiation
            {
                                            
                                            
                                            
                                            var q1=html.indexOf('dropDownOptions:');
                                            var q2=html.indexOf('fromOptionBasedRequest',q1);
                                            var qtyHtml=html.slice(q1,q2);
                                            var qtyArr=qtyHtml.split('description');
                                            
                                            
                                            
                                            
                                            
                                            
                                            
                                            
                                            var arr=[];
                                            var rowTemp=getRow;
                                            var lastRow=getRow+qtyArr.length-2;
                                            var varFrm='=myvariation(A'+getRow+':AH'+getRow+',AA'+getRow+':AD'+lastRow+')';
                                            var validChars="0123456789."
                                            var prodName="";
                                            var mTerms="";
                                            var numSkus=qtyArr.length-1;
                                            //GmailApp.sendEmail("sakib118.biz@gmail.com", "nzcu", getURL+"\n\n"+ getRow+"      "+numSkus)
                                                                      //get current values to avoid overwrite
                                                                    
                                                                       if(numSkus>1) 
                                                                         {
                                                                              var currentValues=sheet.getRange(getRow, 1, numSkus,1).getValues()
                                                                         
                                                                         }
                                                                      else
                                                                      {
                                                                             var currentValues=[[sheet.getRange(getRow,1).getValues()]];
                                                                      
                                                                      }
                                            
                                            //the formula for creating variation
                                            var vFrm='=myvariation(R'+getRow+'C[0],R'+getRow+'C35:R'+getRow+'C50,R[0]C35:R[0]C50)'
                                            
                                            
                                            
                                            
                                            
                                            
                                           
                                            
                                             

                                                      for (var k=1; k<qtyArr.length ; k++) 
                                                      {
                                                            
                                                            if(currentValues[k-1][0]!="")
                                                            {
                                                               
                                                                Browser.msgBox("This will override data in "+rowTemp);
                                                                return 0;  
                                                            
                                                            }
                                                            
                                                            
                                                            var value=qtyArr[k];
                                                            var n2=value.indexOf(" ");
                                                            var id=value.slice(2,n2-1);
                                                            
                                                            
                                                            
                                                            var cell="I"+rowTemp;
                                                            if(mode=="nzcu")
                                                            {
                                                                  var rowWhere='=IFERROR(MATCH(F'+rowTemp+',F2:F,0))';
                                                                  var lenFrm='=LEN(TRIM(L'+rowTemp+'))';
                                                                  initials="";
                                                                  comment=allLenFrm;
                                                                  
                                                                      if(k>=2)
                                                                      {
                                                                              prodName=vFrm; //=myvariation(L$'+getRow+',$AG$'+getRow+':$AN$'+getRow+',$AG'+rowTemp+':$AN'+rowTemp+')';  
                                                                              b2=vFrm;
                                                                              b3=vFrm;
                                                                              b4=vFrm;
                                                                              terms=vFrm;
                                                                              imageMain=""//vFrm;
                                                                              color=vFrm;
                                                                              size=vFrm;
                                                                              material=vFrm;
                                                                              mTerms=vFrm;
                                                                              b1=vFrm;
                                                                              
                                                                      }
                                                                  
                                                            }
                                                            
                                                            //values
                                                            var upc=""//upcs[k-1][0];
                                                            var sellerSku=""; //"DH"+upc
                                                            var imFrm="=T"+rowTemp;
                                                            var lenFrm="=LEN(L"+rowTemp+")"
                                                            
                                                            var m1=value.indexOf(":")+2;
                                                            var m2=value.indexOf("containsProduct",m1);
                                                            var variation=value.slice(m1,m2-3);
                                                            variation = replaceAll(replaceAll(replaceAll(variation,"&quot;",'"'),"\t",""),'\\',"");
                                                            variation=variation.replace(/\r?\n|\r/g,"")
                                                            
                                                            
                                                                                                                        
                                                            var p1=value.indexOf("sellingPrice",m1);
                                                            var p2=value.indexOf("sellingPriceCurrency",p1);
                                                            var price=value.slice(p1+2,p2-3);
                                                            price=removeChars(validChars, price);
                                                            
                                                            
                                                           // var longQty=qtyArr[k-1];
                                                            var q1=value.indexOf("maxQuantity");
                                                            var q2=value.indexOf("status",q1);
                                                            var qty=value.slice(q1+13,q2-2);
                                                            
                                                            var lenFrm2="=LEN(S"+getRow+")";
                                                            var lenFrm3="=LEN(R"+getRow+")";
                                                            arrVals.push([title, initials, today, asin, sellerSku, itemNo, variation, getURL, upc, comment,status, prodName,type, catagory,amPrice, b2, b3, b4, terms, imageMain, color,size,material, imFrm, imFrm, imFrm,lenFrm3,lenFrm2,lenFrm,qty, price, profit, mTerms,b1]);
                                                            
                                                            rowTemp++;
                          
                                          
                                          
                          
                                    
                                                      }//end of for
                                                sheet.getRange(getRow, 1, arrVals.length, arrVals[0].length).setValues(arrVals);
                                                deactivateFormulas(sheet.getRange(getRow, 1, arrVals.length, arrVals[0].length));
            
            }//if there is variations
            
            

      
           
  
  
  
  
  
  
  
  
  
  
  
  
  
  }//end of if column is 8











}




  
 

  




