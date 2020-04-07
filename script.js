// Global variables
const setupSheet = "Setup";
const ss = SpreadsheetApp.getActiveSheet();
const targetLanguage = ss.getRange('B2').getValue();
const targetMarket = ss.getRange('B3').getValue();
const targetLocation = ss.getRange('B4').getValue();
const targetCountry = ss.getRange('B5').getValue();
const deviceType = ss.getRange('B6').getValue();
const clientName = ss.getRange('B7').getValue();

/*
Logger.log(targetLanguage);
Logger.log(targetMarket);
Logger.log(targetLocation);
Logger.log(targetCountry);
Logger.log(executionID); */


//add graphical interface to launch things
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('SerpCrawler 2.0')
      .addItem('Get data', 'summonGods')
      .addItem('Help documents', 'goToHelp')
      .addToUi();
}


// read keywords and their themes
function getKws () {
  const data = [];
  const sheet = SpreadsheetApp.getActive().getSheetByName(setupSheet);
  const values = sheet.getDataRange().getValues();
  for(var i = 8; i < values.length; i++){ // 8 refers to row 9 in the sheet where the actual KWs start
    data.push(values[i]);
  }

  //Logger.log(data);
  return data;
}

// create an unique executionID from the client name and the time at the beginning of the process. This is the same for all kws per one run.
function createExecutionID() {
    const clientName = ss.getRange('B7').getValue();
    const getCurrentTime = Utilities.formatDate(new Date(), "GMT+3", 'YYYYMMdd-HHmmss');
    
    // this is messy but does the job
    const replaceA = clientName.replace(/ä/g, "a").toLowerCase(); 
    const replaceO = replaceA.replace(/ö/g, "o");
    const replaceSpaces = replaceO.replace(/\s+/g, "")
    const formattedClientName = replaceSpaces.replace(/[^\w\s]/gi, "-")
    
    const executionID = formattedClientName + '-' + getCurrentTime;
    Logger.log(executionID);
    
    //write executionID to the sheet
    ss.getRange('E2').setValue(executionID);
    
    return executionID;    
     
}

function summonGods() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Please confirm","Are you sure you wish to summon the SERP Gods? :o ", ui.ButtonSet.YES_NO)
  
  if (response == ui.Button.YES) {
    Logger.log("Gods are summoned :o");
    
    // get executionID from the function
    const execID = createExecutionID();
   
    
    // get all the keywords from the list using getKws() function
    const kws = getKws();
    
    //ping our magical Python elves with requests about keywords to track via the formatPrayer() function.
    for(var i = 0; i < kws.length; i++){
      
      const kwtheme = kws[i][0];
      const kw = kws[i][1];
      Logger.log(kws[i]);
      format = formatPrayer(kwtheme, kw, execID);
    };
    
  } else if (response == ui.Button.NO){
     //log a msg, potentially show a visual msg to user as well
     Logger.log("The Gods may rest... for now.");
     
  } else { //error handling
    Logger.log("You've just confused everybody, including the SERP Gods.");
  }

}

function formatPrayer(kwtheme, kw, execID) {
   
  // replace spaces with + to get them formatted for Python Elves
  const nonspaceKwtheme = kwtheme.replace(/ /g, "+"); 
  const nonspaceKw = kw.replace(/ /g, "+");
  const nonspaceClientName = clientName.replace(/ /g, "+"); 
  const executionID = execID;
    
  const baseUrl = 'CENSORED';
  const fullReq = baseUrl + 'executionid=' + executionID + '&clientname=' + nonspaceClientName + '&keyword=' + nonspaceKw + '&keywordtheme=' + nonspaceKwtheme + '&language=' + targetLanguage + '&market=' + targetMarket + '&country=' + targetCountry + '&location=' + targetLocation;
  const formattedReq = UrlFetchApp.fetch(fullReq, {muteHttpExceptions: true}).getContentText();
   
  Logger.log(fullReq);
   
}

//offer people the option to see Nuclino documentation about how to use this thing.
//placeholder goes to oikio.fi because why not.
//this bit comes from https://support.google.com/docs/thread/16869830?hl=en
function goToHelp() {
  var url = "https://oikio.fi/code/";
  var html = HtmlService.createHtmlOutput('<html><script>'
  +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
  +'var a = document.createElement("a"); a.href="'+url+'"; a.target="_blank";'
  +'if(document.createEvent){'
  +'  var event=document.createEvent("MouseEvents");'
  +'  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'                          
  +'  event.initEvent("click",true,true); a.dispatchEvent(event);'
  +'}else{ a.click() }'
  +'close();'
  +'</script>'
  
  // Offer URL as clickable link in case above code fails.
  +'<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="'+url+'" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
  +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
  +'</html>')
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html, "Opening ..." );
   
}


