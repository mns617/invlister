
function onEditAmTitle()
{
      
      var ss=SpreadsheetApp.getActiveSpreadsheet();
      var sheet=ss.getActiveSheet();
    
      var rng=ss.getActiveRange();
      var row=rng.getRow();
      rng=sheet.getRange(row, 1, rng.getValues().length,sheet.getMaxColumns());
      var values=rng.getValues();
      var mapValues=ss.getSheetByName("Mapping").getDataRange().getValues();
      var ddValues=ss.getSheetByName("Dropdown Options").getDataRange().getValues();
      
      for (var i=0; i<values.length; i++)
      {     
            var vals=[values[i]]; //send one row but as 2D
            if(vals[0][12-1]==""){continue;}
            var col=12;
            onEditAmTitleEachRow(vals, row+i, col, mapValues, ddValues)
      
      }
      

}





function onEditAmTitleEachRow(vals, row, col, mapValues, ddValues) 
{

        var ss=SpreadsheetApp.getActiveSpreadsheet();
        
        var sheet=ss.getActiveSheet();
       

        if(vals==undefined)
        {
                     var rng=ss.getActiveRange();
                     var row=rng.getRow();
                     var col=rng.getColumn();
                     rng=sheet.getRange(row, 1, rng.getValues().length,sheet.getMaxColumns());
                     var values=rng.getValues();
                     var mapValues=ss.getSheetByName("Mapping").getDataRange().getValues();
                     var ddValues=ss.getSheetByName("Dropdown Options").getDataRange().getValues();
                     vals=values[0];
        
        
        }

        
        if(isLister(sheet))  //when someone is editing amazon title
        {
              var amTitle=vals[0][12-1].toString()+" "+vals[0][34-1].toString();
              if(amTitle==""){return 0};
              
              var returns=myLookupColor(replaceAll(amTitle,"_"," "), mapValues, 1, 27, 29)  //macth with column A
              var colors=returns[0];
              Logger.log(colors)
              var sizes=returns[1];
              var patterns=returns[2];
              
               if(patterns.length<1)
               {
                    var searchTerms='No known pattern found, please make Terms manually';
               }
               else if(amTitle.indexOf('error')>=0)
               {
                       var searchTerms='Error found, delete, undo and retry';
               }
               else
               {
                    var searchTerms=   getSearchTerms(patterns, colors, replaceAll(amTitle,"_"," "), sizes, mapValues) 
               }
               sheet.getRange(row, 19).setValue(searchTerms)
                return 0; 
             //find material 
              var rowMaterial= myLookupFullWord(replaceAll(amTitle,"_"," "), mapValues, 11)
              if(rowMaterial!=null)
               { 
                  var material=mapValues[rowMaterial-1][11-1+1];
                  sheet.getRange(row, 23).setValue(material);
               }
              
              //create keyword sentence
              
              var category=findMyCategory(replaceAll(amTitle,"_"," "), mapValues);
              
              if(category!=null)
              {
                  var baseLine= 'Beautiful CCC KKK features a PPP pattern and design.';
                  
                  var pattern=mapValues[patterns[0]-1][29-1]; //patterns[0] is the row number of firt pattern  
                  var kwLine=baseLine.replace('CCC', colors[0].toLowerCase()).replace('KKK', category).replace('PPP', pattern.toLowerCase());
                 //Browser.msgBox(sheet.getRange(row, 18).getFormula());
                 if(sheet.getRange(row, 18).getFormula()!="" || sheet.getRange(row, 18).getValue()=="")    // only overwrite if ther is a formula or a value
                  {
                        //sheet.getRange(row, 18).setValue(kwLine)
                  
                  };
                  
                  
   
              }//category is not null
             
              
              
                    var matchedRow=myLookup(amTitle, mapValues, 1)
                    var color="";
                    if(matchedRow!=null)
                    {
                      color=mapValues[matchedRow-1][2-1];
                      
                    }
                    sheet.getRange(row, 21).setValue(color)
                    
                    var matchedRow2=myLookup(amTitle, mapValues, 4)
                    
                    var size="";
                    if(matchedRow2!=null)
                    {
                      size=mapValues[matchedRow2-1][5-1];
                      
                    }
                    sheet.getRange(row, 22).setValue(size);
                    
                   
                   var matchedRow3=myLookup(amTitle + vals[0][0], mapValues, 7)
                    var mycatagory="";
                    if(matchedRow3!=null)
                    {
                      mycatagory=mapValues[matchedRow3-1][8-1];
                      
                    }
                   sheet.getRange(row, 14).setValue(mycatagory)

              
              
              
              
              
              

        }// when column 12 is edited
       

  
}






































function findMyCategory(amTitle, mapValues)
{
     for(var i=30-1; i<mapValues[0].length; i++)
     {
           var phrase=(mapValues[0][i]).toString().toLowerCase();
           if(fullWordIndexOf(amTitle, phrase)>=0)
           {
               return phrase;
           
           }
           
           
     }

    return null


}









function myLookupColor(val, mapVals, col1, col2, col3)
{
     val=val.toLowerCase();
     var colors=[];
     var sizes=[];
     var patterns=[];
    // var valArr=replaceAll(val,',','').toLowerCase().split(' ');
     var index=10;
     for (var i=1; i<mapVals.length; i++)
     {
          
           //if(mapVals[i][col1-1]==""){continue};
           
           //color
           var tempColor=mapVals[i][col1-1];
           var nn1=fullWordIndexOf(val, mapVals[i][col1-1]);
           if(nn1>=0 && tempColor!='')
           {
                
                if(mapVals[i][col1]!="")
                    {colors.push([nn1,(mapVals[i][col1-1])])}; //store the color in variation
                if(mapVals[i][col1]!="")                
                    {colors.push([nn1,(mapVals[i][col1])])};  //push the color that matches with the mapped base color
                   //return i+1;
            } 
            
            
            //size
            if(fullWordIndexOf(val, mapVals[i][col2-1])>=0)
           {
                
                if(mapVals[i][col2]!="")
                    {sizes.push((mapVals[i][col2]))}; //store the size in variation
                
            } 
            
            
            //
            //pattern
            var multiplePatterns=mapVals[i][col3-1];   
            var patternArr=multiplePatterns.split(','); // we can enter comma seperated patterns in mapping 
            
            for( var p=0; p<patternArr.length; p++)
            {
                   var indx=fullWordIndexOf(val, patternArr[p].trim())
                   if(indx>=0)
                   {
                        
                        
                        if(mapVals[i][col3-1]!="")
                            {
                                            patterns.push([indx, i+1]); //push the row in patterns
                            }
                        
                    } 
            }
            
            
            
            
     } 
     
      //http://stackoverflow.com/questions/16096872/how-to-sort-2-dimensional-array-by-column-value
      patterns.sort(sortFunction);
      colors.sort(sortFunction);
      
      function sortFunction(a, b) {
          if (a[0] === b[0]) {
              return 0;
          }
          else {
              return (a[0] < b[0]) ? -1 : 1;
          }
      }
      
      
     var patternsTemp=[];
    
     for (var p=0; p<patterns.length; p++)
     {
           
           patternsTemp.push(patterns[p][1]);
     
     }
     
     
       
     var colorsTemp=[];
    
     for (var c=0; c<colors.length; c++)
     {
           
           colorsTemp.push(colors[c][1]);
     
     }
     
     
     
     
     
     
     
     var a=colorsTemp.filter(function(item, i, ar){ return ar.indexOf(item) === i; });  //only unique values
     var b=sizes.filter(function(item, i, ar){ return ar.indexOf(item) === i; });  //only unique values
     var c=patternsTemp.filter(function(item, i, ar){ return ar.indexOf(item) === i; });  //only unique values
     
     Logger.log(colorsTemp)
     if(patternsTemp.length>3) {patternsTemp=patternsTemp.slice(0,3)}; //limit maximum 3 patterns
     return [a,b,c];  
}






function getSearchTerms(patterns, colors, amTitle, sizes, mapValues) 
{
     
     var terms=[];
     var size=sizes[0];  //just one size
     
     for(var i=30-1; i<mapValues[0].length; i++)
     {
           var amT=amTitle.toLowerCase()
           var tempC=mapValues[0][i].toLowerCase()
        if(amT.indexOf(tempC)>=0 || amT.indexOf(tempC+"s")>=0)
        {
           
                     for (var j=0; j<patterns.length; j++)
                     {
                                var patternRow=patterns[j];
                                var sTerm=mapValues[patternRow-1][i];  // this is the base search term
                                
                                  for (var k=0; k<colors.length; k++)
                                  {
                                              var sTerm1=replaceAll(sTerm, "CCC", colors[k]);
                                              sTerm1=replaceAll(sTerm1, "SSS", size);
                                              terms.push(sTerm1); 
                                
                                  }
                                
                                
                     
                     }
           
       }
     }//end of for
  
 
  
  
  var searchTerms= terms.join(", ").split(", "); 
  var uniqTerms= searchTerms.filter(function(item, i, ar){ return ar.indexOf(item) === i; })
  
  var ret=[];//uniqTerms[0]; //uniqTerms.join(", ");
  
     for (var i=1; i<uniqTerms.length; i++)
     {
           if(ret.length<850)
           {
               if(uniqTerms[i]==""){continue};
              ret.push(uniqTerms[i]);
           
           }
     
     }
  
    return ret.join(", ").toLowerCase();
  
  
  
}




function lookup5(l_value, sheet2, lookup_col, pick_up_col, value_or_row) {
  var last_row2 = sheet2.getLastRow();
  
  if (last_row2<2) { last_row2 = 2; }
  
  var ar=sheet2.getRange(2,lookup_col,last_row2-1, pick_up_col-lookup_col+1).getValues();
  
  var flag=0;
  for (var i=0; i<last_row2-1; i++)
  {
    var temp1 = ar[i];
    var temp=temp1[0];
    if(temp==l_value)
    {
      flag=1;
      if (value_or_row=="value")   
      {
        return ar[i][temp1.length-1];
        break;
      }
      
      if (value_or_row=="row") {
        return i+2;
        break;
      }
    }
  }
  if(flag==0){ return null; }
}




























//These scripts will help to show already listed images in sidebar for reviewing







function re_checkImagesByVariation()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    
    var rng=ss.getActiveRange();
    var row=rng.getRow();
    rng=sheet.getRange(row, 1, rng.getValues().length,sheet.getMaxColumns());
    
    
    var values=rng.getValues();
    var formulas=rng.getFormulas();
    var formulasR1C1=rng.getFormulasR1C1();
    
     var sbHtml='<br>';
     var htmlArr=[];
                    
    for (var i=0; i<values.length; i++)
    {
          var variation=values[i][7-1];
          var img1=values[i][20-1];
          var img2=values[i][24-1];
          var img3=values[i][25-1];
          var img4=values[i][26-1];
          
          var type1="Regular";
          if(img1.indexOf('a_hflip')>=0)
          {
              type1="Flipped";
          }
          
          if(img1.indexOf('c_crop')>=0)
          {
              type1="Cropped";
          
          }
          
          
           var type2="Regular";
          if(img2.indexOf('a_hflip')>=0)
          {
              type2="Flipped";
          }
          
          if(img2.indexOf('c_crop')>=0)
          {
              type2="Cropped";
          
          }
          
          var type3="Regular";
          if(img3.indexOf('a_hflip')>=0)
          {
              type3="Flipped";
          }
          
          if(img3.indexOf('c_crop')>=0)
          {
              type3="Cropped";
          
          }
          
          
          var type4="Regular";
          if(img4.indexOf('a_hflip')>=0)
          {
              type4="Flipped";
          }
          
          if(img4.indexOf('c_crop')>=0)
          {
              type4="Cropped";
          
          }
          
          
          var tempHtml='<br><br>Variation: <b>'+variation+ '</b><br><br>Main Image: '+type1+'<br><img src="'+img1+'" alt="Not Available" style="width:auto; height:200px;"><br>'          
                                                    + '<br>Image 2 '+type2+'<br><br><img src="'+img2+'" alt="Not Available" style="width:auto; height:200px;"><br>'            
                                                    + '<br>Image 3 '+type3+'<br><br><img src="'+img3+'" alt="Not Available" style="width:auto; height:200px;"><br>'
                                                    + '<br>Image 4 '+type4+'<br><br><img src="'+img4+'" alt="Not Available" style="width:auto; height:200px;"><br>';
                                                    
          htmlArr.push(tempHtml);                                           
    
    }
    
    
    
        
        
                  
                   
                   
                   
      sbHtml=sbHtml+htmlArr.join('<br>');
      Logger.log(sbHtml)
          
                        
                        
      if(sbHtml!='<br>')
      {
            var imhtml = HtmlService.createHtmlOutput(sbHtml)
            .setTitle('Images')
            .setWidth(300);
            SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
            .showSidebar(imhtml);
        
      }



    
    
    
    
    
    
    
 }













function re_checkImagesByImagePosition()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    
    var rng=ss.getActiveRange();
    var row=rng.getRow();
    rng=sheet.getRange(row, 1, rng.getValues().length,sheet.getMaxColumns());
    
    
    var values=rng.getValues();
    var formulas=rng.getFormulas();
    var formulasR1C1=rng.getFormulasR1C1();
    
     var sbHtml='';
     var htmlArr=[];
     
  for (var k=1; k<=4; k++)
  {
          for (var i=0; i<values.length; i++)
          {
                var variation=values[i][7-1];
                
                if(k==1)
                {
                    var img=values[i][20-1];
                }
                else
                {
                    var img=values[i][24+k-2-1];
                
                }
                
                
                var type="Regular";
                if(img.indexOf('a_hflip')>=0)
                {
                    type="Flipped";
                }
                
                if(img.indexOf('c_crop')>=0)
                {
                    type="Cropped";
                
                }
                
                

                  var position="Main Image";
                
                
                 if(k>1)
                {
                  position="Image "+k;
                }
                
                
                
                
                var tempHtml='Variation: <b>'+variation+ '</b><br>'+position+': '+type+'<br><img src="'+img+'" alt="Not Available" style="width:auto; height:200px;"><br>'          
                
                htmlArr.push(tempHtml);                                           
          
            }
            
            sbHtml=sbHtml+"<br><br><br>"+htmlArr.join('<br>')+'-------------------------------------------<br>-------------------------------------------';
            htmlArr=[];

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






