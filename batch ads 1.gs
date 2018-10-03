



function nfl_Caps(title, row) {
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
      var size1="26";//partsTitle[0].match(/\d+/)[0];
      var size2="15";//partsTitle[1].match(/\d+/)[0];
      
       var amTitle="1 Piece Mens Nfl "+name+" Cap, Football Themed Hat Embroidered Team Logo Sports Patterned, Team Logo Fan Athletic Team Spirit Fan Comfortable, Blue White, Heavy Twill"
       var b1="1 Piece Mens Nfl "+fullName+" Cap, Football Themed Hat Embroidered Team Logo Sports Patterned, Team Logo Fan Athletic Team Spirit Fan Comfortable, Blue White, Heavy Twill"
       
       
      var dim='It comes in one size to fit most wearers.';
      var includes="Includes: 1 Cap"
      
      
     
      sheet.getRange(row, 12).setValue(amTitle);
      sheet.getRange(row, 34).setValue(b1);
      sheet.getRange(row, 16).setValue(includes);
      sheet.getRange(row, 17).setValue(dim);
      sheet.getRange(row, 2).setValue("NZCU2NWH");
      sheet.getRange(row, 3).setValue(new Date());
      sheet.getRange(row, 22).setValue("Standard Size")
      sheet.getRange(row, 14).setValue("skull-caps")
     
      var img=showWmImages();
      
      sheet.getRange(row, 20).setValue(img)
      
      Logger.log(amTitle)
      var a=10

  
}



































function nfl_throwPillow(title, row) {
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
      var size1=15;//partsTitle[0].match(/\d+/)[0];
      var size2=15;//partsTitle[1].match(/\d+/)[0];
      
      var amTitle="1 Piece NFL "+name+" Throw Pillow "+size1+" Inches, Football Themed Accent Pillow For Bedroom Sofa Sports Patterned, Team Color Logo Fan Merchandise Athletic Spirit "+color+" Polyester Cotton";
      var b1="1 Piece NFL "+fullName+" Throw Pillow "+size1+" Inches, Football Themed Accent Pillow For Bedroom Sports Patterned, Team Color Logo Fan Merchandise Athletic Team Spirit "+color+" Polyester Cotton";
      var dim='Dimensions: '+size1+' x '+size2+" inches";
      var includes="Includes: 1 throw pillow";
      
      
      
      
      sheet.getRange(row, 12).setValue(amTitle);
      sheet.getRange(row, 34).setValue(b1);
      sheet.getRange(row, 16).setValue(includes);
      sheet.getRange(row, 17).setValue(dim);
      var img=showWmImages();
      
      sheet.getRange(row, 20).setValue(img)
      
      sheet.getRange(row, 2).setValue("NZCU8NTP");
      sheet.getRange(row, 3).setValue("12/4/2017");
      sheet.getRange(row, 3).setValue("15 inches");
      
      Logger.log(amTitle)
      var a=10

  
}





function nfl_throw() {
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
      var partsTitle=sourceTitle.split(" x ");
      var size1=partsTitle[0].match(/\d+/)[0];
      var size2=partsTitle[1].match(/\d+/)[0];
      
      var amTitle='1 Piece Nfl '+name+' Throw Blanket '+size1+' X '+size2+' Inches, Football Themed Oversized Bedding Sports Patterned, Team Logo Fan Merchandise Athletic Team Spirit Fan, '+color+" "+material;
      var b1='1 Piece Nfl '+fullName+' Oversized Throw Blanket '+size1+' X '+size2+' Inches, Football Themed Bedding Sports Patterned, Team Logo Fan Merchandise Athletic Team Spirit Fan, '+color+" "+material;
      var dim='Dimensions: '+size1+' x '+size2+" inches";
      var includes="Includes: 1 NFL throw";
      
      
      
      sheet.getRange(row, 12).setValue(amTitle);
      sheet.getRange(row, 34).setValue(b1);
      sheet.getRange(row, 16).setValue(includes);
      sheet.getRange(row, 17).setValue(dim);
      var img=showWmImages();
      var imFrm='=IMAGE("'+img+'", 1)';
      sheet.getRange(row, 5).setValue(imFrm);
      sheet.getRange(row, 18).setValue('This throw can be used out at a game, on a picnic, in the bedroom, or cuddled under in the den while watching the game on TV.')
      sheet.getRange(row, 20).setValue(img);
      
      sheet.getRange(row, 2).setValue("NZCU4NTB");
      sheet.getRange(row, 3).setValue(new Date());
      
      Logger.log(amTitle)
      var a=10

  
}


function nfl_Dysney_throw(title, row) {
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
      
      var amTitle='2 Piece Nfl '+name+' Throw Blanket  Full Set With Disney Mickey Mouse Character Shaped Pillow, Sports Patterned Bedding Team Logo Fan '+color+" "+material;
      var b1='2 Piece Nfl '+fullName+' Throw Blanket Full Set Disney Mickey Mouse Character Shaped Pillow '+size1+' X '+size2+' Inches, Football Themed Bedding Sports Patterned, Team Logo Fan Merchandise Athletic Team Spirit Fan, '+color+" "+material;
      var dim='Dimensions: '+size1+' x '+size2+" inches";
      var includes="Includes: 1 throw blanket, 1 Pillow";
      sheet.getRange(row, 2).setValue("NZCU4NTB");
      sheet.getRange(row, 3).setValue("12/4/2017");
      
      
      
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








function nfl_HandTowels(title, row) {
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
      var size1="26";//partsTitle[0].match(/\d+/)[0];
      var size2="15";//partsTitle[1].match(/\d+/)[0];
      
      var amTitle="1 Piece Nfl "+name +" Hand Towel "+size1+" X "+size2+" Inches, Football Themed Applique Sports Patterned, Team Logo Fan Merchandise Athletic Spirit, "+color+", "+material
      var b1="1 Piece Nfl "+fullName +" Hand Towel "+size1+" X "+size2+" Inches, Football Themed Applique Sports Patterned, Team Logo Fan Merchandise Athletic Team Spirit Fan, "+color+", "+material
      
      
      var dim='Dimensions: '+size1+' x '+size2+" inches";
      var includes="Includes: 1 hand towel";
      
      
      
     
      sheet.getRange(row, 12).setValue(amTitle);
      sheet.getRange(row, 34).setValue(b1);
      sheet.getRange(row, 16).setValue(includes);
      sheet.getRange(row, 17).setValue(dim);
      sheet.getRange(row, 2).setValue("NZCU2NTS");
      sheet.getRange(row, 3).setValue("12/1/2017");
      sheet.getRange(row, 22).setValue(size1+ " inches")
      sheet.getRange(row, 14).setValue("towel-sets")
     
      var img=showWmImages();
      
      sheet.getRange(row, 20).setValue(img)
      
      Logger.log(amTitle)
      var a=10

  
}






function nfl_GolfTowels(title, row) {
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
      var size1="16";//partsTitle[0].match(/\d+/)[0];
      var size2="19";//partsTitle[1].match(/\d+/)[0];
      
      var amTitle="1 Piece Nfl "+name +" Golf Towel "+size1+" X "+size2+" Inches, Football Themed Applique Sports Patterned, Team Logo Fan Merchandise Athletic Spirit, "+color+", "+material
      var b1="1 Piece Nfl "+fullName +" Golf Towel "+size1+" X "+size2+" Inches, Football Themed Applique Sports Patterned, Team Logo Fan Merchandise Athletic Team Spirit Fan, "+color+", "+material
      
      
      var dim='Dimensions: '+size1+' x '+size2+" inches";
      var includes="Includes: 1 golf towel";
      
      
      
     
      sheet.getRange(row, 12).setValue(amTitle);
      sheet.getRange(row, 34).setValue(b1);
      sheet.getRange(row, 16).setValue(includes);
      sheet.getRange(row, 17).setValue(dim);
      sheet.getRange(row, 2).setValue("NZCU2NTS");
      sheet.getRange(row, 3).setValue("12/1/2017");
      sheet.getRange(row, 22).setValue(size2+ " inches")
      sheet.getRange(row, 14).setValue("towel-sets")
     
      var img=showWmImages();
      
      sheet.getRange(row, 20).setValue(img)
      
      Logger.log(amTitle)
      var a=10

  
}




function nfl_Bathrobe(title, row) {
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
      var size1="26";//partsTitle[0].match(/\d+/)[0];
      var size2="15";//partsTitle[1].match(/\d+/)[0];
      
       var amTitle="1 Piece Nfl "+name+" Mens Large / X-Large Bathrobe, Football Themed Bath Robe For Boys Embroidered Team Logo Sports Patterned, Fan Merchandise Athletic Team Spirit Fan, "+color+", Cotton"; 
      var b1="1 Piece Nfl "+fullName+" Mens Large / X-Large Bathrobe, Football Themed Bath Robe For Boys Embroidered Team Logo Sports Patterned, Fan Merchandise Athletic Team Spirit Fan, "+color+", Cotton";
      var dim='Size: Large / X-large';
      var includes=" It has two front patch pockets, a Silk Touch tie belt, and two belt loops on both the left and right sides for added adjustability.";
      
      
      
     
      sheet.getRange(row, 12).setValue(amTitle);
      sheet.getRange(row, 34).setValue(b1);
      sheet.getRange(row, 16).setValue(includes);
      sheet.getRange(row, 17).setValue(dim);
      sheet.getRange(row, 2).setValue("NZCU2NTS");
      sheet.getRange(row, 3).setValue("12/1/2017");
      sheet.getRange(row, 22).setValue(size1+ " inches")
      sheet.getRange(row, 14).setValue("towel-sets")
     
      var img=showWmImages();
      
      sheet.getRange(row, 20).setValue(img)
      
      Logger.log(amTitle)
      var a=10

  
}



