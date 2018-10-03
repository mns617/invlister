












function deactivateFormulas(rng)
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    
   if(rng==undefined)
   {
    var rng=ss.getActiveRange();
   } 
    
    var row=rng.getRow();
    rng=sheet.getRange(row, 1, rng.getValues().length,sheet.getMaxColumns());
    
    
    var values=rng.getValues();
    var formulas=rng.getFormulas();
    var formulasR1C1=rng.getFormulasR1C1();

    
    
    
    for (var i=0; i<formulasR1C1.length; i++)
    {
          for(var j=0; j<formulasR1C1[0].length; j++)
          {
                if(formulasR1C1[i][j]!="")
                {
                        values[i][j]=replaceAll("'"+formulasR1C1[i][j], row.toString(), '[-'+i+']');
                
                }
                
          
          
          }
    
    }

   rng.setValues(values);

}







function activateFormulas()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    
    var rng=ss.getActiveRange();
    var row=rng.getRow();
    rng=sheet.getRange(row, 1, rng.getValues().length,sheet.getMaxColumns());
    
    
    
    var values=rng.getValues();
    var formulas=rng.getFormulas();
    
    for (var i=0; i<values.length; i++)
    {
          for(var j=0; j<values[0].length; j++)
          {
                if(values[i][j].toString().indexOf("'=")>=0)
                {
                        values[i][j]=(values[i][j]).replace("'=","=");
                
                }
                
                else if(formulas[i][j]!="")
                {
                      values[i][j]=formulas[i][j]
                
                
                }
                
          
          
          }
    
    }

   rng.setValues(values);

}






// create alternate image ads

function createAlternate2()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    
    var rng=ss.getActiveRange();
    var values=rng.getValues();
    var formulas=rng.getFormulas();
    var getRow=rng.getRow();
    
    var targetCol=24; // for 2nd image column
    var vFrm='=myvariation(R'+getRow+'C[0],R'+getRow+'C33:R'+getRow+'C48,R[0]C33:R[0]C48)'
    var values2=values;
    for (var i=0; i<values.length; i++)
    {
         
               if(values[i][11-1]!="COMPLETE")
               {
                    continue
               }
          
          var valuesTemp=values[i][20-1];
          values2[i][20-1]=values[i][targetCol-1];
          values2[i][targetCol-1]=valuesTemp;
          
          
    
    }
   var len=values.length;
   var rowNew=getRow+len+2;
   sheet.insertRowsAfter(rowNew, len);
   var rng2=sheet.getRange(rowNew, 1, len, values[0].length )
   rng2.setValues(values2)


}



function createAlternate3()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    
    var rng=ss.getActiveRange();
    var values=rng.getValues();
    var formulas=rng.getFormulas();
    var getRow=rng.getRow();
    
    var targetCol=25; // for 2nd image column
    var vFrm='=myvariation(R'+getRow+'C[0],R'+getRow+'C33:R'+getRow+'C48,R[0]C33:R[0]C48)'
    var values2=values;
    for (var i=0; i<values.length; i++)
    {
         
               if(values[i][11-1]!="COMPLETE")
               {
                     continue;
               }
          var valuesTemp=values[i][20-1];
          values2[i][20-1]=values[i][targetCol-1];
          values2[i][targetCol-1]=valuesTemp;
          
          
    
    }
   var len=values.length;
   var rowNew=getRow+len+2;
   sheet.insertRowsAfter(rowNew, len);
   var rng2=sheet.getRange(rowNew, 1, len, values[0].length )
   rng2.setValues(values2)


}



function createAlternate4()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    
    var rng=ss.getActiveRange();
    var values=rng.getValues();
    var formulas=rng.getFormulas();
    var getRow=rng.getRow();
    
    var targetCol=26; // for 2nd image column
    var vFrm='=myvariation(R'+getRow+'C[0],R'+getRow+'C33:R'+getRow+'C48,R[0]C33:R[0]C48)'
    var values2=values;
    for (var i=0; i<values.length; i++)
    {
         
               if(values[i][11-1]!="COMPLETE")
               {
                     continue;
               }
          
          var valuesTemp=values[i][20-1];
          values2[i][20-1]=values[i][targetCol-1];
          values2[i][targetCol-1]=valuesTemp;
          
          
    
    }
   var len=values.length;
   var rowNew=getRow+len+2;
   sheet.insertRowsAfter(rowNew, len);
   var rng2=sheet.getRange(rowNew, 1, len, values[0].length )
   rng2.setValues(values2)


}







