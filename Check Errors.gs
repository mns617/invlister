


function manualValidity()
{

        var ss=SpreadsheetApp.getActiveSpreadsheet();
        
        var mapValues=ss.getSheetByName("Mapping").getDataRange().getValues();
        var sheet=ss.getActiveSheet();
        var rng=sheet.getActiveRange();
        
        
        var startRow=rng.getRow();
        var lastRow=rng.getLastRow();
        var msg="<br>";
        
       for(var i=startRow; i<=lastRow; i++)
       {
        var row=i;//rng.getRow();
        var col=rng.getColumn();
        
        var rowValues=sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues();
        
        var amTitle=rowValues[0][12-1];
        if(amTitle=="")
        {
          continue
        }

        var amTitleN=amTitle.toString().trim().toLowerCase().replace(/[\W_]+/g," ");
        var returns=myLookupColor(replaceAll(amTitle,"_"," "), mapValues, 1, 27, 29)  //macth with column A
              var colors=returns[0];
              var sizes=returns[1];
              var patterns=returns[2];
              
        var category=findMyCategory(amTitle, mapValues);
        
        var cRow=row;
        msg= msg+ '<br><br><b>Row '+cRow+":</b> " +checkValidity(rowValues, mapValues, sizes, colors, patterns, category, cRow);
       }
       
       
         var html = HtmlService.createHtmlOutput(msg)
                               .setTitle('Error Check Results')
                               .setWidth(300);
          SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
                        .showSidebar(html);

}


function checkValidity(rowValues, mapValues, sizes, colors, patterns, category, cRow)
   {
           // normalizing using https://stackoverflow.com/questions/20864893/javascript-replace-all-non-alpha-numeric-characters-new-lines-and-multiple-whi
           var ss=SpreadsheetApp.getActiveSpreadsheet();
           
           var variation=rowValues[0][7-1];
           var sourceTitle=rowValues[0][1-1];
           var amTitle=rowValues[0][12-1];
           var msg="";
           
           var variationN="";
           
           
           
           if(variation==""){variation=sourceTitle;} //if variation column is empty try to get it from OS title
           if(variation!="")
             {
                    variationN=variation.toString().trim().toLowerCase().replace(/[\W_]+/g," ");
                    variationN=variationN.replace("|"," ");
               
             }
               
           
           var amTitleN=amTitle.toString().trim().toLowerCase().replace(/[\W_]+/g," ");
           
           var mapValues2=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mapping2').getDataRange().getValues();   
           var matchedSize="";
           var titleSize="";
          


           for(var i=12; i<mapValues2.length; i++)
           {
                       var sizeTemp=(mapValues2[i][4-1]).toString().toLowerCase();
                       if(sizeTemp==""){ continue; }
                       
                       
                      
                       if(fullWordIndexOf(variationN, sizeTemp)>=0)
                       {
                            matchedSize=sizeTemp;
                            titleSize=mapValues2[i][5-1];  //what should be in AM Title column E
                            break;
                       }
   
           
           
           }







           if(matchedSize=="")
           {
                    msg+='<br>-<font color="red"> We are unable determine size of the prouct on from OS/WM Name or Variation for row:  '+cRow +'</font>';
                     
          
            }
           
           
           
             //we have found a matched size
           else
           {
                       if(fullWordIndexOf(amTitleN, titleSize)<0)  // macthed size is not in variation
                       {
                            
                            msg+='<br>-<font color="red">'+ titleSize+' not found in AM Title</font>';
                            
                           // Browser.msgBox('"'+titleSize+ '" not found in Amazon Title');
                           // return 0;
                       
                       }
                       
                       else
                       {
                            //now find the matched size in AM title
                                           var matchedSize2="";
                                           for(var i=0; i<mapValues2.length; i++)
                                           {
                                                       var sizeTemp=mapValues2[i][5-1];
                                                       if(sizeTemp==""){ continue; }
                                                       
                                                       
                                                       
                                                       if(fullWordIndexOf(amTitleN, sizeTemp)>=0)
                                                       {
                                                            matchedSize2=sizeTemp;
                                                            break;
                                                       }
                                           }  
                                           
                                           if(titleSize==matchedSize2)
                                           { 
                                               msg+='<br>-<font color="green">'+ titleSize+' matched</font>';
                                               //msg=msg+"\n"+titleSize+ ' mathced'; 
                                               //Browser.msgBox(msg);
                                           }
                                           else {
                                           
                                                     { 
                                                           msg+='<br>-<font color="red">'+ titleSize+' did not match with '+ matchedSize2 +'</font>';
                                                   
                                                     }//Browser.msgBox(titleSize +" and "+ matchedSize2+ ' not mathced for row'+ cRow);}
                                                }
                                                
                       } //end of else
                 
           }
           
         // ----------------------***------------------
           //column L to S 
             
           
             var sheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
             var row=sheet.getActiveRange().getRow();
             for (var i=12-1; i<19-1; i++)
             {
                   var temp=rowValues[0][i];
                   temp=temp.toString().trim().toLowerCase().replace(/[\W_]+/g," ");
                   //checks if duve cover missing
                   if(temp.indexOf('duvet')>=0 && temp.indexOf('duvet cover')<0)
                   {
                     
                           msg+='<br>-<font color="red">Duvet Cover missing in cell '+ sheet.getRange(row, i+1).getA1Notation() +'</font>';
                     //Browser.msgBox('Duvet Cover missing in cell '+ sheet.getRange(row, i+1).getA1Notation());
                     //return 0
                   
                   }
                   
                   
                   if(temp.indexOf('1 piece')>=0 || temp.indexOf('1 pc')>=0 || temp.toLowerCase().indexOf('single')>=0)
                   {
                       if(temp.indexOf("set")>=0)
                       {
                           
                           msg+='<br>-<font color="red">Set detected with one piece item</font>';
                           //Browser.msgBox('Set detected with one piece item for '+ cRow);
                           //return 0;
                       
                       }
                   
                   
                   }
                   
             
             }
            
             
             
             // https://stackoverflow.com/questions/1183903/regex-using-javascript-to-return-just-numbers
             var setIncludes=(rowValues[0][16-1]).toString().trim().toLowerCase();  //remove all spaces
             msg+=countIncludes(setIncludes, amTitle, mapValues, mapValues2);  //check validaity in set includes
             
             
             //----check category------------
            var catTitle= findMyCategory(amTitle, mapValues);
            var catOs=findMyCategory(rowValues[0][0],mapValues);
            if(catOs!=catTitle)
            {
                  if(catOs==null)
                  {
                      if(catTitle.indexOf('Bed in a Bag')>=0)
                      {
                          catOs="Comforter"
                      
                      }
                  
                  }
                  
                  
                  if(catOs!=null) 
                  {
                    
                    
                          msg+='<br>-<font color="red">'+ catOs+' did not match with '+ catTitle +' in AM Title</font>';
                    
                        //Browser.msgBox('"'+catOs+ '" found from source title didn\'t match '+catTitle+ ' found from Amazon Title for row '+cRow);
                  }
            
            }
            
            
            var catInc=findMyCategory(rowValues[0][15],mapValues);
            if(catInc!=catTitle)
            {
                  
                  
                  msg+='<br>-<font color="red">'+ catInc+' in Set Includes did not match with '+ catTitle +' in AM Title</font>';
                  //Browser.msgBox('"'+catInc+ '" found from Set Includes didn\'t match "'+catTitle+ '" found from Amazon Title for row '+cRow);
                 
            
            }
            
            
            
            var catDim=findMyCategory(rowValues[0][16],mapValues);
            if(catDim!=catTitle)
            {
                  
                  msg+='<br>-<font color="red">'+ catDim+' in Dimensions did not match with '+ catTitle +' in AM Title</font>';
                  //Browser.msgBox('"'+catDim+ '" found from Dimensions didn\'t match "'+catTitle+ '" found from Amazon Title for row '+cRow);
                        
            }
   
            var catTerms=findMyCategory(rowValues[0][18],mapValues);
            if(catTerms!=catTitle && catTerms!=null)
            {
                  
                  msg+='<br>-<font color="red">'+ catTerms+' in Terms did not match with '+ catTitle +' in AM Title</font>';
                  //Browser.msgBox('"'+catTerms+ '" found from Search Terms didn\'t match "'+catTitle+ '" found from Amazon Title for row '+cRow);
                        
            }
    
            var tempColN=rowValues[0][14-1];
            if(tempColN=="pillowcase-and-sheet-sets"){tempColN="sheet-sets"}
            var catColN=findMyCategory(replaceAll(tempColN,"-"," "),mapValues);
             if(catColN!=catTitle)
            {
                   msg+='<br>-<font color="red">'+ catColN+' in column N did not match with '+ catTitle +' in AM Title</font>';
                  //Browser.msgBox('"'+catColN+ '" found from column N didn\'t match "'+catTitle+ '" found from Amazon Title for row '+cRow);
                        
            }
            
            //term size checking
            
            var sizeTitle =sizes[0]; //size determined from title for search terms
            var terms=rowValues[0][19-1]; //col S
            for (var i=0; i<mapValues.length; i++)
            {
                  var tempSize=mapValues[i][28-1];
                  if(terms.indexOf(tempSize)>=0)  //this size is found in terms
                  {
                        if(tempSize!=sizeTitle && tempSize!="")
                        {     
                              
                              
                               msg+='<br>-<font color="red">Invalid size '+ tempSize+' found in serch term.</font>';
                              //Browser.msgBox('Invalid size '+sizeTemp + ' found in serch term for row '+ cRow);
                              //return 0;
                        
                        }
                  
                  }
            
            
            
            }
             
            return msg
   
            
   
   }





  function countIncludes(setIncludes, amTitle, mapValues, mapValues2)
  {
      
      var msg=""
      setIncludes=replaceAll(setIncludes, "_", "");  //eliminate under score
      
      setIncludes=replaceAll(setIncludes, "( ","(");  //elimiate any comma after the bracket
      setIncludes=replaceAll(setIncludes, " )",")");
       
       
      amTitle=replaceAll(amTitle.toLowerCase(), "_", " ");
      
      var setPlural=setIncludes;  //set icnludes for plural detection
      
      var count=0;
      
      for( var r=1; r<15 ; r++)  //random loop assuimg max 15 items in a set
      {      
      
            for (var i=1; i<mapValues2.length; i++)
            {
                    var temp=(mapValues2[i][7]).toString().toLowerCase().trim();
                    if(setIncludes.indexOf(temp)>=0 && temp!="")
                    {     
                          var tempCount=mapValues2[i][8];  //col H
                          Logger.log(temp)
                          
                          count=count+Number(tempCount);
                          setPlural=replaceAll(setPlural,temp,"|"+temp);  // add a seperator for later use to find plural error
                          Logger.log(setPlural)
                          setIncludes=setIncludes.replace(temp, "");  //remove the matched count so it does not get counted again
                         
                          break;
                    }
            
            
            }
       }     
      if(count==0){
         msg+='<br>-<font color="red">No Set Includes found.</font>';
      
      };
      
      
      var titlePieces="";
      for (var i=0; i<mapValues.length; i++)
      {
                var tempValue=(mapValues[i][17-1]).toLowerCase();//
                
                if(tempValue !="" && fullWordIndexOf(amTitle,tempValue)>=0)
                {
                        titlePieces=tempValue;
                        break;
                
                }
      
      }
      
      
      
      if(titlePieces=="")
      {
             msg+='<br>-<font color="red">No piece found in AM Title.</font>';
      
      }
      var titleCount=Number(titlePieces.match(/\d+/g)[0]);  //https://stackoverflow.com/questions/1183903/regex-using-javascript-to-return-just-numbers
     
      if(count>1 || titleCount>1)
      {
              if(amTitle.toLowerCase().indexOf("set")<0)
              {
                  msg+='<br>-<font color="red">Set not found with multiple peice item</font>';
              }
      
      }
     
     
      if(count==titleCount)
      {
           
           msg+='<br>-<font color="green">'+count+' piece matched</font>';
           // Browser.msgBox(count +' piece matched');
           // return 0;
      }
      
      else
      {
           msg+='<br>-<font color="red">Set Includes count '+count+' not matched with Title piece count- '+titleCount+'</font>';
          //Browser.msgBox("Set includes count: "+count+ " didn't match with title piece count "+titleCount);
          //return 0;
      
      }
      
      
      //plural error finding start
      
      var setPluralArr=setPlural.split("|");
      
          for (var i=0; i<setPluralArr.length; i++)
          {
              var set=setPluralArr[i].toString().toLowerCase();
              
              
              if(set.indexOf('one')>=0 || set.indexOf('1')>=0)
             {
             
                    set=replaceAll(set," ","");
                    var lastChar=set.slice(-1);
                    if(lastChar=='s')
                    {
                        msg+='<br>-<font color="red">Suspected plural noun used with 1 piece</font>';
                  
                    }
              }
              
          
          }
      
      
      
      
      
      
      
      
      
      
      return msg;
      
      
      
      
      
  
  
  }











function variationChecking()
{
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var rng=ss.getActiveRange();
  var sheet=rng.getSheet();
  
  var sr=rng.getRow();
  var lr=rng.getLastRow();
  
  var values=sheet.getRange(sr,1, lr-sr+1, sheet.getMaxColumns()).getValues();
  var result="<br>";
     
     for (var c=12-1; c<=34-1 ; c++) //column L to ah
     {
           var refText=values[0][c]; 
           if(refText==""){continue};
           
           if(c==13-1 || c==14-1 || c==15-1 || c==20-1 || c==21-1){continue};
           
           if(c>=24-1 && c<=32-1){continue};  //column X to AF
           
           for (var r=1; r<values.length; r++)  //move along the rows
           {
                     var tempText=values[r][c];
                     
                     for (var k=35-1; k<=50-1; k++)  //for all keys
                     {
                           var refKey=values[0][k];
                           var tempKey=values[r][k];
                           
                           if(refKey==""){continue}
                           if(tempKey==""){tempKey=refKey};
                           
                           var n1=countOccurance(refKey, refText);
                           var n2=countOccurance(tempKey, tempText);
                           if(n1 != n2 && n1>0)
                           {
                                Logger.log(refText+"\n"+refKey+"\n"+n1+"\n"+tempKey+"\n"+n2)
                              result+="<br>"+ sheet.getRange(sr+r, c+1).getA1Notation()+ "--> "+tempKey+ " for "+refKey ;
                           
                           }
                           
                        
                           
                           
                     }//end of k for
           
           }//end of r for
     
     }//end of c for
     
             
                  if(result!='<br>')
                  {
                      var imhtml = HtmlService.createHtmlOutput(result)
                          .setTitle('Results')
                          .setWidth(300);
                          SpreadsheetApp.getUi() 
                          .showSidebar(imhtml);
                    
                   }     
                        
  
  

}





function countOccurance(phrase, string)
{
    var array=string.toString().split(" ");
    var count=0;
    for (var i=0; i<array.length; i++)
    {
        if(array[i].toLowerCase()==phrase.toLowerCase())
        {count++};
    
    }
    return count

}








