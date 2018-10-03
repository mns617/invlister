

function importAll()
{
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sheet=ss.getActiveSheet();
    for (var i=503; i<600; i++)
    {
          if(sheet.getRange(i, 8).getValue()==""){continue};
          sheet.getRange(i, 8).activate();
          //nfl_showerCurtain(5, 5);
          nfl_BathTowels("", "");
          //importFromSource1();
          //nfl_rugs("", "")
    
    
    }

}








function nfl_throw(title, row) {
      var ss=SpreadsheetApp.getActiveSpreadsheet();
      var sheet=ss.getActiveSheet();
      var rng=sheet.getActiveRange();
      var row=rng.getRow();
      var col=rng.getColumn();
      
      var values=sheet.getRange(row, 1,1, sheet.getMaxColumns()).getValues();
      var sourceTitle=values[0][0];
      
      var details= findTeam(sourceTitle, 9, 1); //[team, color, longName, material]
      var fullName=details[2];
      var name=details[0];
      var color=details[1];
      var material=details[3];
      var partsTitle=sourceTitle.split("x");
      var size1=partsTitle[0].match(/\d+/)[0];
      var size2=partsTitle[1].match(/\d+/)[0];
      
      var amTitle='1 Piece Nfl '+name+' Throw Blanket '+size1+' X '+size2+' Inches, Football Themed Bedding Sports Patterned, Team Logo Fan Merchandise Athletic Team Spirit Fan, '+color+" "+material;
      var b1='1 Piece Nfl '+fullName+' Throw Blanket '+size1+' X '+size2+' Inches, Football Themed Bedding Sports Patterned, Team Logo Fan Merchandise Athletic Team Spirit Fan, '+color+" "+material;
      var dim='Dimensions: '+size1+' x '+size2+" inches";
      var includes="Includes: 1 throw blanket";
      
      
      
      
      sheet.getRange(row, 12).setValue(amTitle);
      sheet.getRange(row, 34).setValue(b1);
      sheet.getRange(row, 16).setValue(includes);
      sheet.getRange(row, 17).setValue(dim);
      var img=showWmImages();
      
      sheet.getRange(row, 20).setValue(img)
      
      Logger.log(amTitle)
      var a=10

  
}




function nfl_showerCurtain(title, row) {
      var ss=SpreadsheetApp.getActiveSpreadsheet();
      var sheet=ss.getActiveSheet();
      var rng=sheet.getActiveRange();
      var row=rng.getRow();
      var col=rng.getColumn();
      
      var values=sheet.getRange(row, 1,1, sheet.getMaxColumns()).getValues();
      var sourceTitle=values[0][0];
      
      var details= findTeam(sourceTitle, 9, 1); //[team, color, longName, material]
      var fullName=details[2];
      var name=details[0];
      var color=details[1];
      var material="Polyester";//details[3];
      var partsTitle=sourceTitle.split("x");
      var size1="72"//partsTitle[0].match(/\d+/)[0];
      var size2="72";//partsTitle[1].match(/\d+/)[0];
      
      var amTitle='1 Piece Nfl '+name+' Showe Curtain '+size1+' X '+size2+' Inches, Football Themed Bedding Sports Patterned, Team Logo Fan Merchandise Bathroom Curtain Athletic Team Spirit Fan, '+color+" "+material;
      var b1='1 Piece Nfl '+fullName+' Showe Curtain '+size1+' X '+size2+' Inches, Football Themed Bedding Sports Patterned, Team Logo Fan Merchandise Bathroom Curtain Athletic Team Spirit Fan, '+color+" "+material;
      var dim='Dimensions: '+size1+' x '+size2+" inches";
      var includes="Includes: 1 decorative shower curtain";
      
      
      
      
      sheet.getRange(row, 12).setValue(amTitle);
      sheet.getRange(row, 34).setValue(b1);
      sheet.getRange(row, 16).setValue(includes);
      sheet.getRange(row, 17).setValue(dim);
      sheet.getRange(row, 2).setValue("NZCU2NSC");
      sheet.getRange(row, 3).setValue("12/1/2017");
      sheet.getRange(row, 22).setValue(size2+ " inches")
      sheet.getRange(row, 14).setValue("shower-curtains")
      var img=showWmImages();
      
      sheet.getRange(row, 20).setValue(img)
      
      Logger.log(amTitle)
      var a=10

  
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
    var color=values1[i][col1-2]; 
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













function nfl_rugs(title, row) {
      var ss=SpreadsheetApp.getActiveSpreadsheet();
      var sheet=ss.getActiveSheet();
      var rng=sheet.getActiveRange();
      var row=rng.getRow();
      var col=rng.getColumn();
      
      
      var values=sheet.getRange(row, 1,1, sheet.getMaxColumns()).getValues();
      var sourceTitle=values[0][0];
      if(sourceTitle==""){return 0};
      
      var details= findTeam(sourceTitle, 9, 1); //[team, color, longName, material]
      var fullName=details[2];
      var name=details[0];
      var color=details[1];
      var material="Polyester";//details[3];
      var partsTitle=sourceTitle.split(" x ");
      var size1=partsTitle[0].match(/\d+/)[0];
      var size2=partsTitle[1].match(/\d+/)[0];
      
      var amTitle=size1+'" x '+size2+'" NFL '+name+' Mat For Boys, Football Themed Bath Rug Sports Patterned Rectangular Bathroom Carpet, Team Logo Fan Merchandise Athletic Spirit, '+color+" "+material;
      var b1=size1+'" x '+size2+'" NFL '+fullName+' Mat For Boys, Football Themed Bath Rug Sports Patterned Rectangular Bathroom Carpet, Team Logo Fan Merchandise Athletic Spirit, '+color+" "+material;
      var dim='Dimensions: '+size1+' x '+size2+" inches";
      var includes="Face: 100 percent micro polyester fabric, Filling: 100 percent polyurethane foam, Back: 100 percent PVC";
      
      
      
      
      sheet.getRange(row, 12).setValue(amTitle);
      sheet.getRange(row, 34).setValue(b1);
      sheet.getRange(row, 16).setValue(includes);
      sheet.getRange(row, 17).setValue(dim);
      var img=showWmImages();
      
      sheet.getRange(row, 20).setValue(img)
      
      Logger.log(amTitle)
      var a=10

  
}








function nfl_BathTowels(title, row) {
      var ss=SpreadsheetApp.getActiveSpreadsheet();
      var sheet=ss.getActiveSheet();
      var rng=sheet.getActiveRange();
      var row=rng.getRow();
      var col=rng.getColumn();
      
      
      var values=sheet.getRange(row, 1,1, sheet.getMaxColumns()).getValues();
      var sourceTitle=values[0][0];
      if(sourceTitle==""){return 0};
      
      var details= findTeam(sourceTitle, 9, 1); //[team, color, longName, material]
      var fullName=details[2];
      var name=details[0];
      var color=details[1];
      var material="Polyester";//details[3];
      var partsTitle=sourceTitle.split(" x ");
      var size1="25";//partsTitle[0].match(/\d+/)[0];
      var size2="50";//partsTitle[1].match(/\d+/)[0];
      
      var amTitle="1 Piece Nfl "+name +" Bath Towel "+size1+" X "+size2+" Inches, Football Themed Applique Shower Towel Sports Patterned, Team Logo Fan Merchandise Athletic Spirit, "+color+", "+material
      var b1="1 Piece Nfl "+fullName +" Bath Towel "+size1+" X "+size2+" Inches, Football Themed Applique Shower Towel Sports Patterned, Team Logo Fan Merchandise Athletic Team Spirit Fan, "+color+", "+material
      
      
      var dim='Dimensions: '+size1+' x '+size2+" inches";
      var includes="Includes: 1 bath towel";
      
      
      
     /* 
      sheet.getRange(row, 12).setValue(amTitle);
      sheet.getRange(row, 34).setValue(b1);
      sheet.getRange(row, 16).setValue(includes);
      sheet.getRange(row, 17).setValue(dim);
      sheet.getRange(row, 2).setValue("NZCU2NTS");
      sheet.getRange(row, 3).setValue("12/1/2017");
      sheet.getRange(row, 22).setValue(size2+ " inches")
      sheet.getRange(row, 14).setValue("towel-sets")
     */ 
      var img=showWmImages();
      
      sheet.getRange(row, 20).setValue(img)
      
      Logger.log(amTitle)
      var a=10

  
}










