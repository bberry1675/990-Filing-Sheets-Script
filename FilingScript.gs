
function rangeAggregateXmlFiles(){
  var range = SpreadsheetApp.getActiveRange();
  var sheet = range.getSheet();
  if(sheet.getName() == 'Unique Id to URL'){
    var time = (new Date()).toLocaleString();
    var values = sheet.getRange(range.getRow(),1,range.getHeight(),4).getValues();
    var aggregate = XmlService.createDocument(XmlService.createElement('Return'));
    var attachs = {}
    var compressed;
    var lastEIN = parseInt(values[0][1],10);
    values.forEach(function(row){
      if(parseInt(row[1],10) != lastEIN){
        attachs[lastEIN] = aggregate;
        aggregate = XmlService.createDocument(XmlService.createElement('Return'));
        lastEIN = parseInt(row[1],10);
      }
      compressed = compressXMLDocument( getFilingXML(row[3]) );
      aggregate = aggregateDocument(aggregate, compressed, row[0]);
    });
    attachs[lastEIN] = aggregate;
    var savedFileNames = [];
    Object.keys(attachs).forEach(function(ein){
      savedFileNames.push(ein+'_'+time+'.xml');
      DriveApp.addFile(DriveApp.createFile(ein+'_'+time+'.xml', XmlService.getPrettyFormat().format(attachs[ein]), 'text/xml'))
    });
    
    SpreadsheetApp.getActiveSpreadsheet().toast(savedFileNames.join('\n'), 'Saved Files to Drive');
  }
  else{
    SpreadsheetApp.getActiveSpreadsheet().toast("Must be on 'Unique Id to URL'");
  }
  
}

function aggregateXmlFiles(){
  var startTime = (new Date()).getTime();
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("Unique Id to URL");

  var currentTime = startTime;
  var aggregateDoc;
  var lastEIN;
  var currentRow;
  var currentLineNumber;
  var currentFiling;
  var endLine;
  var props = PropertiesService.getScriptProperties();
  currentLineNumber = props.getProperty('aggregate_curr_row');
  try{
  //if there was no current row then set the current row to 1 which is the recond row in the spreadsheet
  currentLineNumber = (currentLineNumber == null) ? 0 : currentLineNumber;
  currentLineNumber = parseInt(currentLineNumber,10);

  console.log("starting at line: " + currentLineNumber);
  
  lastEIN = -1;
  
  while((currentTime - startTime < 300000) && currentLineNumber < sheet.getLastRow() ){
    
    currentRow = sheet.getRange(currentLineNumber + 2, 1, 1, 4).getValues()[0];
    
    if(currentRow[1] != lastEIN){
      //save the aggregate doc
      if(aggregateDoc != undefined){
        //Logger.log('finishing: ' + currentRow[1].toString());
        saveXmlObject(aggregateDoc,currentRow[1].toString());
      }
      
      //set aggregate doc to new one
      aggregateDoc = getExisitingAggregateDocument(currentRow[1].toString());
      //Logger.log('got to after getting existing aggregate document');
      
    }
    
    currentFiling = getFilingXML(currentRow[3]);
    aggregateDoc = aggregateDocument(aggregateDoc, compressXMLDocument(currentFiling) , currentRow[0].toString() );
    
    
    
    lastEIN = currentRow[1];
    currentLineNumber = currentLineNumber + 1;
    //get the time again
    currentTime = (new Date()).getTime();
  }
  
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger){
    if(trigger.getHandlerFunction()	== 'aggregateXmlFiles'){
      ScriptApp.deleteTrigger(trigger);
    }
  });
      
  
  //save the current document being aggregated
  if(currentLineNumber >= sheet.getLastRow()){
      props.deleteProperty('aggregate_curr_row');
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Aggregate Xml Filings', "Script has finished aggregating all xml filings in " + ss.getName() + " into 'Aggregate Filings' drive folder.");
  }
  else{
      props.setProperty('aggregate_curr_row', currentLineNumber);
      console.log('Finishing at line: ' + currentLineNumber);
      

      ScriptApp.newTrigger('aggregateXmlFiles').timeBased().after(90000).create();
  }
  
  saveXmlObject(aggregateDoc,lastEIN);
  
  }catch(err){
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(),'Failure to Aggregate Xml Filings', err.message);
  }


}



//takes document object of xml and saves it
function saveXmlObject(document,fileName){
  var oldFileID = getExistingAggregateDocumentID(fileName);
  var driveFile;
  var folderIter = DriveApp.getFoldersByName("Aggregate Filings");
  var aggregateFolder = folderIter.next();
  var file = DriveApp.createFile(fileName + '.xml', XmlService.getPrettyFormat().format(document), 'text/xml');

  if(oldFileID != null){
    DriveApp.getFileById(oldFileID).setTrashed(true);
  }

  aggregateFolder.addFile(file);
}

function aggregateDocument(aggregate, compressedDoc, compressedDocQuantifier){
  Object.keys(compressedDoc).reverse().forEach(function(header){
    //get ride of the beginning / and split the string into the node names
    var path = header.slice(1).split('/');
    //assuming the root element will always be Return
    path.shift();
    //Logger.log('Travelling path: ' + path);
    var currentNode = aggregate.getRootElement();
    var namespace = currentNode.getNamespace();
    //Logger.log(typeof compressedDocQuantifier);
    //Logger.log('<' + compressedDocQuantifier + '></'+compressedDocQuantifier+'>');
    //var elementToAdd = XmlService.parse('<' + compressedDocQuantifier + '></'+compressedDocQuantifier+'>').detachRootElement();
    var elementToAdd = XmlService.createElement('UNIQUE' + compressedDocQuantifier.trim());
    
    //Logger.log('after element to add');
    elementToAdd.setNamespace(namespace);
    elementToAdd.setText(compressedDoc[header]);
    
    path.forEach(function(currentPath){
      var children = currentNode.getChildren(currentPath,namespace);
      //if there is a child node with the name of the current header
      if(children.length > 0){
        //Logger.log('Found Child For: ' + currentPath);
        currentNode = children[0];
        //Logger.log(currentNode);
      }
      //if there isn't a child node with the same name means we have to create it
      else{
        //Logger.log('Making a new node for: ' + currentPath);
        currentNode.addContent(XmlService.createElement(currentPath, namespace));
        currentNode = currentNode.getChildren(currentPath,namespace)[0];
        //Logger.log(currentNode);
      }
    });

    currentNode.addContent(elementToAdd);    
  });
  
  return aggregate;
}

function getExisitingAggregateDocument(EIN){
  var id = getExistingAggregateDocumentID(EIN);
  if(id == null){
    return XmlService.createDocument(XmlService.createElement('Return'));;
  }
  else{
  
    return XmlService.parse(DriveApp.getFileById(id).getBlob().getDataAsString());
  }
}

//returns the ID of the document or null if it doesn't exist
function getExistingAggregateDocumentID(EIN){
  var folderIter =  DriveApp.getFoldersByName("Aggregate Filings");
  var folder;
  if(folderIter.hasNext() == false){
    folder = DriveApp.createFolder("Aggregate Filings");
  }
  else{
    folder = folderIter.next();
  }
  
  var fileIter = folder.getFiles();
  var file;
  var curr;
  while(file === undefined){
    if(fileIter.hasNext()){
      curr = fileIter.next();
      if(curr.getName().indexOf(EIN) != -1){
        file = curr;
      }
    }
    else{
      file = null;
    }
  }
  
  return (file == null) ? null : file.getId();
}

//takes a Document class from XmlService.parse
//returns an object with the headers being the keys and the values being values
function compressXMLDocument(xmlDoc){
  var root = xmlDoc.getRootElement();
  var retObj = {};
  var elementsToVisit = [];
  elementsToVisit.push(root);
  while(elementsToVisit.length > 0){
    var current = elementsToVisit.pop();
    if(current.getChildren().length == 0){
      var name = getElementPath(current);
      //Logger.log(name);
      var value = current.getText();
      retObj[name] = value;
    }
    else{
      current.getChildren().forEach(function(child){elementsToVisit.push(child);});
    }
  }
  return retObj;
}

//gets the Document object for an XML page at URL
function getFilingXML(url){
  //Logger.log(url);
  var response = UrlFetchApp.fetch(encodeURI(url).replace('%0D','')).getContentText();
  return XmlService.parse(response);
}

function getElementPath(element){
  if(element == null){
    return '';
  }
  else{
    return getElementPath(element.getParentElement()) + '/' + element.getName();
  }
}


function placeHeaders(ss){
  if(ss != null){
    var rawKeys = ss.getSheetByName('RawKeys');
    if(rawKeys != null){
      var headers = {};
      var keyMatrix = rawKeys.getRange(1,2,rawKeys.getLastRow(), rawKeys.getLastColumn() - 1).getValues();
      for(var i = 0 ; i < keyMatrix.length; i++){
        for(var j = 0; j < keyMatrix[i].length; j++){
          if(headers[keyMatrix[i][j]] === undefined){
            headers[keyMatrix[i][j]] = 1;
          }
          else{
            headers[keyMatrix[i][j]] = headers[keyMatrix[i][j]] + 1;
          }
        }
      }

      
    }
  }
}

function makeSSInFolder(folder, name){
  var newSS = SpreadsheetApp.create(name);
  newSS.insertSheet("RawKeys");
  newSS.insertSheet("RawValues");
  newSS.insertSheet("Compiled View");
  newSS.deleteSheet(newSS.getSheetByName('Sheet1'));
  var firstID = newSS.getId();
  var oldFile = DriveApp.getFileById(firstID);
  var newFile = oldFile.makeCopy(folder);
  
  oldFile.setTrashed(true);
  newFile.setName(newFile.getName().substring(8,newFile.getName().length));
  PropertiesService.getScriptProperties().deleteProperty('export_start_row');
  return newFile.getId();
  
}

function getParsingDocuments(){
  var folderiter = DriveApp.getFoldersByName('Parsed 990 Filings');
  var folder;
  var parseIterator;
  var parseSheets = {};
  //if there is no folders in drive then make the folder / files
  if(folderiter.hasNext() == false){
    //make the folder in drive
    //set to folder variable
    folder = DriveApp.createFolder('Parsed 990 Filings');
    
    //make the parse documents
    //set the parse files
    var listOfYears = getIndexYears();
    listOfYears.forEach(function(element) {
      parseSheets[element] = makeSSInFolder(folder, element+'_ParsedFilings');
    });
    
    
  }
  else{
    folder = folderiter.next();
    parseIterator = folder.getFiles();
    //if there are no files then make the parse documents
    var listOfYears = getIndexYears();
    //if the parse sheets are missing
    if(parseIterator.hasNext() == false){
      //make the parse sheets
      listOfYears.forEach(function(element) {
        parseSheets[element] = makeSSInFolder(folder, element+'_ParsedFilings');
      });
      
    }
    else{
      var tempiter;
      listOfYears.forEach(function(element) {
        tempiter = folder.getFilesByName(element + '_ParsedFilings');
        //if the folder has that file then add it to the return object
        if(tempiter.hasNext()){
          parseSheets[element] = tempiter.next().getId();
        }
        //if the file doesn't exist then make and add to return object
        else{
          parseSheets[element] = makeSSInFolder(folder, element+'_ParsedFilings');
          
        }
      });
      
    }
  }
  
  return parseSheets;
}

function testheir(element){
  if(element == null){
    return '';
  }
  else{
    return testheir(element.getParentElement()) + '->' + element.getName();
  }
}

function parseXML(url){
  var xml = UrlFetchApp.fetch(encodeURI(url).replace('%0D','')).getContentText();
  var document = XmlService.parse(xml);
  var root = document.getRootElement();
  var descptionList = [];
  var valueList = [];
  var runner = function(element){
    if(element.getChildren().length == 0){
      descptionList.push(testheir(element));
      valueList.push(element.getText());
    }
    else{
      element.getChildren().forEach(runner);
    }
  };
  root.getChildren().forEach(runner);
  return [descptionList,valueList];
}

function makeParseSSObject(years, documentIDs){
  var retValue = {};
  years.forEach(function(year){
    var tempObj = {};
    var tempSS = SpreadsheetApp.openById(documentIDs[year]);
    tempObj['rawValues'] = tempSS.getSheetByName('RawValues');
    tempObj['rawKeys'] = tempSS.getSheetByName('RawKeys');
    retValue[year] = tempObj;
  });
  return retValue;
}

function extendedExportFilingsToParseDocuments() {
    var startTime = (new Date()).getTime();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var scriptVariables = getStaticVariables();
    var toExportSheet = ss.getSheetByName(scriptVariables.exportSheet);
    if (toExportSheet != null) {
        try {
            ss.toast("Email will be sent to " + Session.getEffectiveUser().getEmail() + ' with script output.', "Starting to parse 990 Filings");
            
            //row to start, either left over from last execution or starting at 0
            var startRow = PropertiesService.getScriptProperties().getProperty('export_start_row') == null ? 1 : Math.floor(PropertiesService.getScriptProperties().getProperty('export_start_row'));

            console.info('Starting at row: %d', startRow);

            var parseCounter = 0;

            //2d array for the rows in export sheet
            var filingsToParse = toExportSheet.getRange(startRow + 1, 1, toExportSheet.getLastRow(), toExportSheet.getLastColumn()).getValues().filter(function(row){return row[0] != '';});

            //list of years for filings
            var years = getIndexYears();
            //nested object of parse sheets
            var parseSSs = makeParseSSObject(years, getParsingDocuments());

            for (var row = 0; row < filingsToParse.length; row++) {

                //if the max time is coming up then make the trigger and stop current execution
                if (((new Date()).getTime() - startTime) > 300000) {

                    var newRowStart = startRow + parseCounter;

                    //set the current row to be used next time
                    PropertiesService.getScriptProperties().setProperty('export_start_row', Math.floor(newRowStart));

                    console.info('Saving row %d for next time', newRowStart);

                    //clear any previous triggers
                    var triggers = ScriptApp.getProjectTriggers();
                    triggers.forEach(function(trigger) {
                        //if the function ran is this function then delete the trigger
                        if (trigger.getHandlerFunction() == 'extendedExportFilingsToParseDocuments') {
                            ScriptApp.deleteTrigger(trigger);
                        }
                    });
                    //make the new trigger to run after 1 minute
                    ScriptApp.newTrigger('extendedExportFilingsToParseDocuments').timeBased().after(90000).create();
                    //return to finish execution
                    return;
                }
                //continue down the rows in export sheet and parse the documents
                else {
                    var currentRow = filingsToParse[row];
                    Logger.log(filingsToParse.length);
                    Logger.log(row);
                    Logger.log(filingsToParse[row]);

                    //currentRow -> [UniqueId, EIN, Year, URL]
                    //get all the uniqueIDs that have already been parsed
                    var currentParsed = parseSSs[currentRow[2]]['rawKeys'].getRange(1, 1, ((parseSSs[currentRow[2]]['rawKeys'].getLastRow() > 0) ? parseSSs[currentRow[2]]['rawKeys'].getLastRow() : 1), 1).getValues().map(function(row2) {
                        return row2[0]
                    });
                    //if the doc hasn't been parsed yet parse it and append as row
                    if (currentParsed.indexOf(currentRow[0]) == -1) {
                        var parsedInfo = parseXML(currentRow[3]);
                        parsedInfo[0].unshift(currentRow[0]);
                        parsedInfo[1].unshift(currentRow[0]);
                        Logger.log(currentRow[2]);
                        Logger.log(parsedInfo[0].length);
                        parseSSs[currentRow[2]]['rawKeys'].appendRow(parsedInfo[0]);
                        parseSSs[currentRow[2]]['rawValues'].appendRow(parsedInfo[1]);
                        parseCounter = parseCounter + 1;
                    }

                }
            }
            //once for loop is done send email saying the script has finished
            MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Execution Notice', 'Parsing the XML files from ' + ss.getName() + ' has finished');
            PropertiesService.getScriptProperties().deleteProperty('export_start_row');
        } catch (err) {
            MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Execution Notice', 'Parsing the XML files from ' + ss.getName() + ' has failed due to \n' + err);
            PropertiesService.getScriptProperties().deleteProperty('export_start_row');
        }

    } else {
        ss.toast('Cannot find "' + scriptVariables.exportSheet + '" Sheet in Spreadsheet');
    }
}

/*
  Function to delete the folder in Drive for the parse documents
*/
function cleanParsingThenParse(){
  var parsingFolderIter = DriveApp.getFoldersByName('Parsed 990 Filings');
  if(parsingFolderIter.hasNext() == false){
    extendedExportFilingsToParseDocuments();
    return;
  }
  
  while(parsingFolderIter.hasNext()){
    var folder = parsingFolderIter.next();
    DriveApp.removeFolder(folder);
  }
  extendedExportFilingsToParseDocuments();
}

/*
  insert sheet to convert the Unique Ids to URLS for parsing the XML files
*/
function insertUniqueToId(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var staticVariables = getStaticVariables();
  
  //remove the old sheet
  if(ss.getSheetByName(staticVariables.exportSheet)){
    ss.deleteSheet(ss.getSheetByName(staticVariables.exportSheet));
  }
  //insert the new sheet
  var newSheet = ss.insertSheet(staticVariables.exportSheet);
  //add the first row with descriptions in the new sheet
  newSheet.appendRow(['Unique ID', 'EIN', 'Year', 'URL']);
  
  //make the formula to be inserted into the id to url sheet
  //=uniqueIDToURL('Fetched API Data'!B2:J,'Fetched API Data'!A2:A,'Fetched API Data'!B1:1)
  // parameter meaning - Unique Id cells, EIN values, Years
  var insertFormula = "=uniqueIDToURL('" + staticVariables.dataSheetName + "'!B2:" + incrementChar("B",getIndexYears().length - 1) + ",'" + staticVariables.dataSheetName + "'!A2:A, '" + staticVariables.dataSheetName + "'!B1:1)";
  newSheet.getRange(2,1).setFormula(insertFormula);
  addFullCustomMenu();
  
}

/*
  Static store of reused variables
*/
function getStaticVariables(){
  var returnObj = {
    urlBeginning:'https://s3.amazonaws.com/irs-form-990/index_',
    fileType:'.csv',
    dataSheetName: 'Fetched API Data',
    rangePropKey: 'EIN_RANGE',
    exportSheet: 'Unique Id to URL'
  }
  return returnObj;
}

/*
  On open trigger to add the custom menu to user client
*/
function onOpen(){
  var variables = getStaticVariables();
  if(SpreadsheetApp.getActive().getSheetByName(variables.dataSheetName)){
    addFullCustomMenu();
  }
  else{
    addBasicCustomMenu();
  }
}

/*
  Function to add a custom menu to active spreadsheet. Menu only has option to create the data 
  sheet for fetching documents associated with EIN numbers.
*/
function addBasicCustomMenu(){
  var ui = SpreadsheetApp.getUi();
  var customMenu = ui.createMenu('Custom Menu');
  var dataSheetSubMenu = ui.createMenu('API Data').addItem('Add Data Sheet', 'addDataSheet');
  customMenu.addSubMenu(dataSheetSubMenu).addToUi();
  
}

/*
  Function to add the full custom menu to active spreadsheet. Options include adding the data sheet,
  refreshing the api calls, and updating the data sheet with a new range for EIN values
*/
function addFullCustomMenu(){
  var ui = SpreadsheetApp.getUi();
  var customMenu = ui.createMenu('Custom Menu');
  var dataSheetSubMenu = ui.createMenu('API Data').addItem('Add Data Sheet', 'addDataSheet').addItem('Manual API Fetch Refresh', 'refreshCalls').addItem('Update EIN Range','updateEINRange');
  var ExportSubMenu = ui.createMenu('Export Menu');
  //if Unique to ID sheet exists then give options to remake the sheet, full export, regular export
  if(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getStaticVariables().exportSheet)){
    ExportSubMenu.addItem('Make ID to URL Sheet','insertUniqueToId').addItem('Hard Parse all URLs', 'cleanParsingThenParse').addItem('Soft Parse URLs ', 'extendedExportFilingsToParseDocuments').addItem('Aggregate Xml Files', 'aggregateXmlFiles').addItem('Aggregate Range','rangeAggregateXmlFiles');
  }
  //if not then only give option to create the sheet
  else{
    ExportSubMenu.addItem('Make ID to URL Sheet','insertUniqueToId');
  }
  customMenu.addSubMenu(dataSheetSubMenu).addSubMenu(ExportSubMenu).addToUi();
}

/*
  Function to return a list of years starting with the first index of filing and ending with the current year
*/
function getIndexYears(){
  var firstYearInDataBase = 2011;
  var currentYear = new Date().getFullYear();
  var years = [];
  for(var i = firstYearInDataBase; i <= currentYear; i++){
    years.push(i);
  }
  return years;
}

/*
  Helper function to increment letter by {increment} amount to create A1 Notation
*/
function incrementChar(c,increment) {
    return String.fromCharCode(c.charCodeAt(0) + increment);
}

/*
  Function to create the sheet to display the fetched filing data
  User is prompted to enter a range in A1 Notation for the EIN values
  Custom functions are placed in sheet to seperate fetch calls
*/
function addDataSheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var range = ui.prompt("Enter range for EIN numbers in A1 notation").getResponseText();
  var variables = getStaticVariables();
  try{
    ss.getRange(range);
  }
  catch(e){
    ui.alert("Range is not valid");
    return;
  }
  var props = PropertiesService.getDocumentProperties();
  props.setProperty(variables.rangePropKey, range);
  if(ss.getSheetByName(variables.dataSheetName) == null){
    var sheet = ss.insertSheet(variables.dataSheetName);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);
    var years = getIndexYears();
    var firstRow = ['EIN Numbers'];
    years.forEach(function(element) {firstRow.push(element);});
    Logger.log(firstRow);
    sheet.getRange('A1:' + incrementChar('A',firstRow.length - 1) + '1').setValues([firstRow]);
    sheet.getRange('A2').setValue('=ARRAYFORMULA(UNIQUE(' + range + '))');
    var yearsToFunction = years.map(function(element){ return '=fetchFilings(' + element + ',' + 'UNIQUE(' + props.getProperty(variables.rangePropKey) + '))'});
    sheet.getRange(2,2,1,years.length).setValues([yearsToFunction]);
    sheet.getRange('B:' + incrementChar('B',sheet.getLastColumn() - 1)).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    addFullCustomMenu();
    
  }
}

/*
  Custom function to be used in cell to fetch filing data,
  Parameters include the year for the filings and the range 
  in A1 Notation for the EIN numbers,
*/
function fetchFilings(year, einRange){
  var einList = einRange.map(function(row) {return parseInt(row[0]);});
  var variables = getStaticVariables();
  var response = UrlFetchApp.fetch(variables.urlBeginning + year + variables.fileType);
  var returnList = [];
  var cvsObj = {};
  if(response.getResponseCode() == 200){

    var cvsString = response.getContentText();
    delete response;
    cvsString.split('\n').forEach(function(row) {
      var record = row.split(',');
      if(einList.indexOf(parseInt(record[2])) != -1){
        if(cvsObj[parseInt(record[2])]){
          if(Array.isArray(cvsObj[parseInt(record[2])])){
            cvsObj[parseInt(record[2])].push(record[record.length - 1]);
          }
          else{
            cvsObj[parseInt(record[2])] = [cvsObj[parseInt(record[2])], record[record.length - 1]];
          }
        }
        else{
          cvsObj[parseInt(record[2])] = record[record.length - 1];
        }
      }
    });
    einList.forEach(function(element){
      if(Array.isArray(cvsObj[element])){
        returnList.push(cvsObj[element].join());
      }
      else{
        returnList.push(cvsObj[element]);
      }
    });
  }
  else{
    return [[response.getResponseCode() +' response code']];
    
  }
  return returnList.map(function(element){ return [element]; });
}

/*
  Function to manually refresh the api calls by resetting the formulas
*/
function refreshCalls(){
  var ss = SpreadsheetApp.getActive();
  var variables = getStaticVariables();
  var sheet = ss.getSheetByName(variables.dataSheetName);
  var ui = SpreadsheetApp.getUi();
  if(sheet){
    var functionCallRange = sheet.getRange(2, 2, 1, sheet.getLastColumn() - 1);
    functionCallRange.clearContent();
    SpreadsheetApp.flush();
    var years = sheet.getRange(1,2,1,sheet.getLastColumn() - 1).getValues()[0];
    var props = PropertiesService.getDocumentProperties();
    var einRange = props.getProperty(variables.rangePropKey);
    var formulasToSet = []
    years.forEach(function(year){formulasToSet.push('=fetchFilings('+ year + ',' + einRange +')');});
    functionCallRange.setFormulas([formulasToSet]);
  }
  else{
    addBasicCustomMenu();
    ui.alert('No Data Sheet to manually refresh fetched data for.');
    
  }
}

/*
  Function to prompt the user to input a new range for EIN values
  Replaces the stored properties range and puts in formulas for new range
  into data sheet
*/
function updateEINRange(){
 var ss = SpreadsheetApp.getActive();
 var ui = SpreadsheetApp.getUi();
 var variables = getStaticVariables();
 if(ss.getSheetByName(variables.dataSheetName)){
   var userRange = ui.prompt('Enter range for EIN numbers in A1 notation').getResponseText();
   var sheet = ss.getSheetByName(variables.dataSheetName);
   try{
     sheet.getRange(userRange);
   }
   catch(e){
     ui.alert('Range is not valid');
     return;
   }
   var props = PropertiesService.getDocumentProperties();
   props.setProperty(variables.rangePropKey, userRange);
   var formulaRange = sheet.getRange(2,1,1,sheet.getLastColumn());
   var years = sheet.getRange(1,2,1,sheet.getLastColumn() - 1).getValues()[0];
   var formulasToSet = [];
   formulasToSet.push('=ARRAYFORMULA('+ userRange +')');
   years.forEach(function(year){
     formulasToSet.push('=fetchFilings('+ year + ',' + userRange +')');
   });
   formulaRange.setFormulas([formulasToSet]);
 }
 else{
   addBasicCustomMenu();
   ui.alert('No Data Sheet to update range for.');
 }
}

function makeURL(uniqueID){
  //https://s3.amazonaws.com/irs-form-990/{UNIQUEID}_public.xml
  return 'https://s3.amazonaws.com/irs-form-990/' + uniqueID + '_public.xml';
}

function einTest(values){
  var valueList = values.map(function(element){return element[0];});
  var countObj = {};
  var retValue = [];
  valueList.forEach(function(element){
    if(countObj[element]){
      countObj[element] = countObj[element] + 1;
    }
    else{
      countObj[element] = 1;
    }
  });
  valueList.forEach(function(element){
    retValue.push([element, countObj[element]]);
  });
  return retValue;
}

function uniqueIDToURL(idRange, einNumbers, years){
  var einList = einNumbers.filter(function (element) {
    if(element.length == 0 || element[0] == ''){
      return false;
    }
    else{
      return true;
    }
  }).map(function(element){
    return element[0];
  });
  var yearList = years[0].filter(function(element){
    if(element == ''){
      return false;
    }
    else{
      return true;
    }
  });
  var dataRange = idRange.slice(0,einList.length);
  var returnList = [];
  for(var i = 0; i < dataRange.length; i++){
    for(var j = 0; j < dataRange[i].length; j++){
    var currentCell  = dataRange[i][j];
      if(currentCell != ''){
        currentCell.split(',').forEach(function(element) {
          returnList.push([element, einList[i], yearList[j],makeURL(element)]);
        });
      }
    }
  }
  
  return returnList;
}


