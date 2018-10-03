function getMyJson(html)
{
            
        var n1=html.indexOf('</div><script id="atf-content" type="application/json">')+('</div><script id="atf-content" type="application/json"> ').length;
        var n2=html.indexOf('</script>',n1);

        var html2=html.slice(n1,n2);
        
      if(html2.slice(0,1)=="{")
      {
        var jsonData=JSON.parse(html2)['atf-content'];
      }
      else
      {
                  var n1=html.indexOf('WML_REDUX_INITIAL_STATE')+13+15;
                  var n2=html.indexOf('</script>',n1)-3;
                  var html2=html.slice(n1,n2);
                   jsonData=JSON.parse(html2)
        
      
      }
     
       return jsonData;

}




function getMySellerJson(html)
{
            
                     var htmlSeller = html;
                     var n1=htmlSeller.indexOf('</div><script id="contentSeller" type="application/json"> ')+('</div><script id="contentSeller" type="application/json"> ').length;
                     var n2=htmlSeller.indexOf('</script>',n1);
                     var htmlSeller=htmlSeller.slice(n1,n2);
                     if(htmlSeller.slice(0,1)=="{")
                     {
                         var jsonDataSeller=JSON.parse(htmlSeller).contentSeller;
                     }
                     
                     else
                     {
                     
                         var n1=html.indexOf('window.__WML_REDUX_INITIAL_STATE__ = ')+('window.__WML_REDUX_INITIAL_STATE__ = ').length;
                         var n2=html.indexOf('</script>',n1)-3;
                         var htmlSeller=html.slice(n1,n2);
                         var jsonDataSeller=JSON.parse(htmlSeller);
                     }
                     
                     return jsonDataSeller;

}



function getMyJsonSearch(html)
{
            
        var n1=html.indexOf('</div><script id="atf-content" type="application/json">')+('</div><script id="atf-content" type="application/json"> ').length;
        var n2=html.indexOf('</script>',n1);

        var html2=html.slice(n1,n2);
        
      if(html2.slice(0,1)=="{")
      {
        var jsonData=JSON.parse(html2)['atf-content'];
      }
      else
      {
                  var n1=html.indexOf('WML_REDUX_INITIAL_STATE')+13+15;
                  var n2=html.indexOf('</script>',n1);
                  
                  
                //  Browser.msgBox(html.slice(n2-500, n2))
                  var html2=html.slice(n1,n2-1);
                  
                   jsonData=JSON.parse(html2)
        
      
      }
     
       return jsonData;

}

























// returns true if user is on Lisert Sheet
function isLister(sheet)
{
    if(sheet.getName().toLowerCase().indexOf('lister')==0) return true
    return false;
}



function combineTerms(values)
{
    var autoTerms=values[0][19-1];
    var manualTerms=values[0][33-1];
    if(manualTerms=="") {var combinedTerms=autoTerms}
    else {var combinedTerms=autoTerms+", "+manualTerms;}
    combinedTerms=replaceAll(combinedTerms, "_", " ").trim();
    combinedTerms=replaceAll(combinedTerms, "_ ", "_");
    
    return combinedTerms.split(", ").filter(function(item, i, ar){ return ar.indexOf(item) === i; }).join(", ");  //make unqiue

}





function checkUndefined(rowValues,row)
{
      for (var i=12-1; i<rowValues.length; i++)
      {
            var temp=rowValues[i].toString();
            if(temp.toLowerCase().indexOf('undefined')>=0)
            {
                    var msg="undefined error found in row "+row;
                    return msg;           
            }
            
            else if(temp.toLowerCase().indexOf('error')>=0)
            {
                    var msg="'#error' error found in row "+row;
                    return msg;           
            }
            
      
      }
      
      return "";


}
















function myLookup(val, mapVals, col)
{
     val=val.toLowerCase();
     for (var i=1; i<mapVals.length; i++)
     {
           if(mapVals[i][col-1]==""){continue};
           if(val.indexOf((mapVals[i][col-1]).toLowerCase())>=0)
           {return i+1}         
     } 
     
     return null;


}




function  myLookupFullWord(val, mapVals, col)

{

    
     val=val.toLowerCase();
     for (var i=1; i<mapVals.length; i++)
     {
           if(mapVals[i][col-1]==""){continue};
           var phrase= (mapVals[i][col-1]).toLowerCase();
           //if(val.indexOf((mapVals[i][col-1]).toLowerCase())>=0)
           if(fullWordIndexOf(val, phrase)>0)
           {return i+1}         
     } 
     
     return null;


}




function fullWordIndexOf(text, phrase)
{
    //text='this is a garbage text';
    //phrase='tex';
    text =replaceAll(text,',','')
    text=text.toString().toLowerCase();
    phrase=phrase.toString().toLowerCase();
    
    if(phrase==text){return 0};
    
    var n1=text.indexOf(phrase+' ');
    if(n1==0){return n1};              //if the phrase is the start of the sentence
    
    var n2=text.indexOf(' '+phrase+' ')
    if(n2>=0){return n2};
    
    var n3=text.indexOf(' '+phrase)
    var len=text.length; //length of 
    var len2=len-(phrase.length+1);  //length of prase ans a space,
    if(n3==len2){return n3};
    
    return -1

}


