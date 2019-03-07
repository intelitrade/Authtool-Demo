function getStyleSheet(unique_title) {
  
  debugger;
  debugger;
  unique_title = "testit";
  var sheet = document.styleSheets[1];
  sheet.rules.item(13).style.display="none";
  return;
  for(var i=0; i<document.styleSheets.length; i++) {
    var sheet = document.styleSheets[i];
    if(sheet.title == unique_title) {
      return sheet;
    }
  }
}