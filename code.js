// JIRA Fields Global readonly Metadata
let storyIdField = "";
let podField = "";
let assigneeField = "";
let storyPointsField = "";
let actualEndDateColumnField = "";
let plannedEndDateColumnField = "";
let originalEstimateField = "";
let timeSpentField = "";
let remainingTimeField = "";
let storyStatusField = "";

//Non-JIRA Global readonly metadata 
let completeStatusString = "";
let incompleteStatusString = "";
let devNotStartedString = "";

/**
 * Function to be executed when sheet opens
*/
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('JIRA Menubar')
   .addItem('Click to Retrieve', 'logJiraInfo')
   .addToUi();
}

function initializeMetaData(){
  let storiesMetaDataSheet = SpreadsheetApp.getActive().getSheetByName('Metadata');
  let metadata = storiesMetaDataSheet.getDataRange().getValues();
        
  for(let rownum=1; rownum < metadata.length; rownum++){
    if((metadata[rownum][0]) === "storyIdField"){
      storyIdField = metadata[rownum][1];

    } else if((metadata[rownum][0]) === "podField"){
      podField = metadata[rownum][1];

    } else if((metadata[rownum][0]) === "assigneeField"){
      assigneeField = metadata[rownum][1];

    } else if((metadata[rownum][0]) === "storyPointsField"){
      storyPointsField = metadata[rownum][1];

    } else if((metadata[rownum][0]) === "actualEndDateColumnField"){
      actualEndDateColumnField = metadata[rownum][1];

    } else if((metadata[rownum][0]) === "plannedEndDateColumnField"){
      plannedEndDateColumnField = metadata[rownum][1];

    } else if((metadata[rownum][0]) === "originalEstimateField"){
      originalEstimateField = metadata[rownum][1];

    } else if((metadata[rownum][0]) === "timeSpentField"){
      timeSpentField = metadata[rownum][1];

    } else if((metadata[rownum][0]) === "remainingTimeField"){
      remainingTimeField = metadata[rownum][1];

    } else if((metadata[rownum][0]) === "storyStatusField"){
      storyStatusField = metadata[rownum][1];
    } else if((metadata[rownum][0]) === "completeStatusString"){
      completeStatusString = metadata[rownum][1];
    } else if((metadata[rownum][0]) === "incompleteStatusString"){
      incompleteStatusString = metadata[rownum][1];
    } else if((metadata[rownum][0]) === "devNotStartedString"){
      devNotStartedString = metadata[rownum][1];
    } 
  }
}

function roundToTwo(num) {
  return +(Math.round(num + "e+2") + "e-2");
}

/**
* This function is to calculate the non working days between two dates
* @param {*} startDate 
* @param {*} endDate 
*/
function getNonWorkingDaysBtwTwoDates(startDate, endDate) {
  let counter = 0;
  let datesSheet = SpreadsheetApp.getActive().getSheetByName('Calendar');
  let datesData = datesSheet.getDataRange().getValues();
  let nonWorkingDays = [];

  for (let rownum = 1; rownum < datesData.length; rownum++) {
      if (datesData[rownum][0] !== ""){
        nonWorkingDays.push(new Date(datesData[rownum][0]));
      }
  }

  nonWorkingDays.forEach(nonWorkingDay => {
    if ((nonWorkingDay >= startDate) && (nonWorkingDay <= endDate)) {
      counter++;
    }
  });
  return counter;
}

/**
* This function returns the total working days in a sprint
*/
function getSprintWorkingDays(){
let datesSheet = SpreadsheetApp.getActive().getSheetByName('Calendar');
let datesData = datesSheet.getDataRange().getValues();
let totalWorkingDays = [];

for (let rownum = 1; rownum < datesData.length; rownum++) {
  if (datesData[rownum][1] !== ""){
    totalWorkingDays.push(new Date(datesData[rownum][0]));
  }
}
return totalWorkingDays;
}

/**
* This function is used to log into google sheets information from JIRA
*
*/
function logJiraInfo() {
  initializeMetaData();
  let storiesDumpSheet = SpreadsheetApp.getActive().getSheetByName('StoriesDump');
  let storiesAnalysisSheet = SpreadsheetApp.getActive().getSheetByName('StoriesChart');
 
  let data = storiesDumpSheet.getDataRange().getValues();
  //First get the indexes right for the Actual Start Date, Actual End Date, Planned Start Date and Due Date
  let storyIdIndex = data[0].indexOf(storyIdField);
  let podIndex = data[0].indexOf(podField);
  let assigneeIndex = data[0].indexOf(assigneeField);
  let storyPointsIndex = data[0].indexOf(storyPointsField);
  let actualEndDateColumnIndex = data[0].indexOf(actualEndDateColumnField);
  let plannedEndDateColumnIndex = data[0].indexOf(plannedEndDateColumnField);
  let originalEstimateIndex = data[0].indexOf(originalEstimateField);
  let timeSpentIndex = data[0].indexOf(timeSpentField);
  let remainingTimeIndex = data[0].indexOf(remainingTimeField);
  let storyStatusIndex = data[0].indexOf(storyStatusField);

  let today = new Date();

  storiesAnalysisSheet.clear();
  storiesAnalysisSheet.appendRow(["Key", "<ProjectName> POD", "Delay in Start - Not yet started", "Delay In Progress - Due date based", "ETC - Not yet Started", "ETC - In Progress Stories", "Delay - Completed", "Story Status", "Original Estimate", "Remaining Effort", "Time Already Spent", "Assignee", "JIRA Status", "Story Points", "Planned Due Date", "Projected Due Date"]);

  let finalStoriesArray = [];

  for (let rownum = 1; rownum < data.length; rownum++) {

      let storyKey = data[rownum][storyIdIndex];
      let storyPoints = data[rownum][storyPointsIndex];
      if (storyPoints === "") {
          storyPoints = 0;
      }
      let jiraStatus = data[rownum][storyStatusIndex];

      let effortRemaining = data[rownum][remainingTimeIndex] / 28800;
      let plannedDueDate = data[rownum][plannedEndDateColumnIndex];

      let projectedDueDate = new Date();

      let storyStatusCategory;

      //If the story does not have a planned date, skip it
      if (plannedDueDate === "") {
          continue;
      } else {
          plannedDueDate = new Date(data[rownum][plannedEndDateColumnIndex]);
          plannedDueDate.setHours(23, 59, 59);
      }

      //Status for not started Story
      let yetToStart_delayInStart = 0;
      let yetToStart_ETC = roundToTwo((data[rownum][originalEstimateIndex]) / 28800);;

      if ((devNotStartedString.indexOf(data[rownum][storyStatusIndex])) !== -1) {

          let plannedStartDate = new Date(plannedDueDate.getTime() - effortRemaining * (24 * 3600 * 1000));

          if ((plannedStartDate.getDay() > 5) || (plannedStartDate.getDay() < 1)) {
              plannedStartDate.setDate(plannedStartDate.getDate() - 2);
          }
          if (plannedStartDate < today) {
              yetToStart_delayInStart = roundToTwo((today.getTime() - plannedStartDate.getTime()) / (24 * 3600 * 1000)) - getNonWorkingDaysBtwTwoDates(plannedStartDate, today);
          }

          if (yetToStart_delayInStart === 0) {
              projectedDueDate = plannedDueDate;
          } else {
              projectedDueDate.setDate(plannedDueDate.getDate() + roundToTwo(yetToStart_ETC));
              projectedDueDate.setDate(projectedDueDate.getDate() + getNonWorkingDaysBtwTwoDates(plannedDueDate, projectedDueDate));
          }

          storyStatusCategory = "Not yet Started";
      }

      //Status for in-progress Story
      let delayInProgress_DueDateBased = 0;
      let delayInProgress_ETC = 0;

      if ((incompleteStatusString.indexOf(data[rownum][storyStatusIndex])) !== -1) {

          let estimatedCompletionDate = new Date(today.getTime() + effortRemaining * (24 * 3600 * 1000));
          if ((estimatedCompletionDate.getDay() > 5) || (estimatedCompletionDate.getDay() < 1)) {
              estimatedCompletionDate.setDate(estimatedCompletionDate.getDate() + 2);
          }

          let nonWorkingDays_InProgress;

          if (estimatedCompletionDate < plannedDueDate) {
              nonWorkingDays_InProgress = getNonWorkingDaysBtwTwoDates(estimatedCompletionDate, plannedDueDate);
              delayInProgress_DueDateBased = -1.0 * (roundToTwo(((plannedDueDate.getTime() - estimatedCompletionDate.getTime()) / (24 * 3600 * 1000))) - nonWorkingDays_InProgress);
              console.log("Story id: " + storyKey + "::::" + plannedDueDate + "::::" + estimatedCompletionDate);
          } else {
              nonWorkingDays_InProgress = getNonWorkingDaysBtwTwoDates(plannedDueDate, estimatedCompletionDate);
              delayInProgress_DueDateBased = roundToTwo(((estimatedCompletionDate.getTime() - plannedDueDate.getTime()) / (24 * 3600 * 1000))) - nonWorkingDays_InProgress;
              console.log("Story id: " + storyKey + "::::" + plannedDueDate + "::::" + estimatedCompletionDate + "::::" + delayInProgress_DueDateBased);
          }

          projectedDueDate = estimatedCompletionDate;
          delayInProgress_ETC = roundToTwo((data[rownum][remainingTimeIndex]) / 28800);
          storyStatusCategory = "In Progress";
      }

      //Status for completed Story
      let delayInCompletion_Completed = 0;
      if ((completeStatusString.indexOf(data[rownum][storyStatusIndex])) !== -1) {
          let actualCompletionDate = data[rownum][actualEndDateColumnIndex];
          storyStatusCategory = "Dev Complete";
          if (actualCompletionDate !== "") {
              // Need to account for non-working days - remove them
              delayInCompletion_Completed = roundToTwo((actualCompletionDate.getTime() - plannedDueDate.getTime()) / (24 * 3600 * 1000));
          }
      }
     
     finalStoriesArray.push([data[rownum][storyIdIndex], data[rownum][podIndex], yetToStart_delayInStart, delayInProgress_DueDateBased, yetToStart_ETC, delayInProgress_ETC, delayInCompletion_Completed, storyStatusCategory, data[rownum][originalEstimateIndex], data[rownum][remainingTimeIndex], data[rownum][timeSpentIndex], data[rownum][assigneeIndex], jiraStatus, storyPoints, plannedDueDate, projectedDueDate]);
  }
  // Batch Update
  storiesAnalysisSheet.getRange(storiesAnalysisSheet.getLastRow() + 1, 1, finalStoriesArray.length, finalStoriesArray[0].length).setValues(finalStoriesArray);
  getPODStatusGeneric();
}

function getPODStatusGeneric(){
  let podNames = getPODNames();
  let consolidatedPODStatus = {};
  
  let storiesAnalysisSheet = SpreadsheetApp.getActive().getSheetByName('StoriesChart');
  let storiesCartData = storiesAnalysisSheet.getDataRange().getValues();
  let sprintWorkingDays = getSprintWorkingDays();
  
  for(let rownum = 0; rownum<podNames.length; rownum++){
    let podName = podNames[rownum].trim();
    consolidatedPODStatus[podName]={};
    consolidatedPODStatus[podName].OverallSP = 0;
    consolidatedPODStatus[podName].SPAchieved = 0;
    consolidatedPODStatus[podName].SPExpected = 0;
    consolidatedPODStatus[podName].SPExpectedTillYesterday = 0;
    consolidatedPODStatus[podName].OriginalEstimates = 0;
    consolidatedPODStatus[podName].ETCRemaining = 0;
    consolidatedPODStatus[podName].ETCExpected = 0;
  }
  
  let podIndex = storiesCartData[0].indexOf("<ProjectName> POD");
  let storyPointsIndex = storiesCartData[0].indexOf("Story Points");
  let completionStatusIndex = storiesCartData[0].indexOf("Story Status");
  let plannedEndDateColumnIndex = storiesCartData[0].indexOf("Planned Due Date");
  let originalEstimatesIndex = storiesCartData[0].indexOf("Original Estimate");
  let etcRemainingIndex = storiesCartData[0].indexOf("Remaining Effort");
  
  let eodToday = new Date();
  eodToday.setHours(23, 59, 59);

  let eodYesterday = new Date();
  eodYesterday.setHours(23, 59, 59);
  eodYesterday.setDate(eodYesterday.getDate() - 1);
  
  for (let rownum = 1; rownum < storiesCartData.length; rownum++) {
    let podName = (storiesCartData[rownum][podIndex]).trim().toUpperCase();
    let value = consolidatedPODStatus[podName];
   
    try{
     consolidatedPODStatus[podName].OverallSP += storiesCartData[rownum][storyPointsIndex];
    
      // SP Status
      if (storiesCartData[rownum][completionStatusIndex] === "Dev Complete") {
        consolidatedPODStatus[podName].SPAchieved += storiesCartData[rownum][storyPointsIndex];
      }
      if (storiesCartData[rownum][plannedEndDateColumnIndex] <= eodToday) {
        consolidatedPODStatus[podName].SPExpected += storiesCartData[rownum][storyPointsIndex];
      }
     
      if (storiesCartData[rownum][plannedEndDateColumnIndex] <= eodYesterday) {
        consolidatedPODStatus[podName].SPExpectedTillYesterday += storiesCartData[rownum][storyPointsIndex];
      }
      
      // ETC Status
      if (storiesCartData[rownum][originalEstimatesIndex] !== "") {
        consolidatedPODStatus[podName].OriginalEstimates += storiesCartData[rownum][originalEstimatesIndex];
      }
      if (storiesCartData[rownum][etcRemainingIndex] !== "") {
        consolidatedPODStatus[podName].ETCRemaining += storiesCartData[rownum][etcRemainingIndex];
      }
       
      let index = 0;
      for (let i = 0; i < sprintWorkingDays.length; i++) {
        if (sprintWorkingDays[i] > new Date()) {
          index = i + 1;
          break;
        }
      }
      consolidatedPODStatus[podName].ETCExpected = consolidatedPODStatus[podName].OriginalEstimates * (index / sprintWorkingDays.length);
    
    }catch(err){
    }
    //console.log("Hello");
  }
  console.log(consolidatedPODStatus);

  //Add data to visualization sheet
  let podStatusSheet = SpreadsheetApp.getActive().getSheetByName('POD Status');
  podStatusSheet.clear();
  podStatusSheet.appendRow(["Responsive POD", "SP - Total", "SP - Expected by EOD", "SP - Achieved", "Original Estimates", "ETC - Expected", "ETC - Pending", "SP - Efficiency", "ETC - Overflow", "SP - Expected by Yesterday"]);
  
  let podStatusKeys = Object.keys(consolidatedPODStatus);
  for(let rownum = 0; rownum <podStatusKeys.length; rownum++){
    let key = podStatusKeys[rownum];
    podStatusSheet.appendRow([key,Math.round(consolidatedPODStatus[key].OverallSP), Math.round(consolidatedPODStatus[key].SPExpected), Math.round(consolidatedPODStatus[key].SPAchieved), roundToTwo(consolidatedPODStatus[key].OriginalEstimates / 28800), roundToTwo(consolidatedPODStatus[key].ETCExpected / 28800), roundToTwo(consolidatedPODStatus[key].ETCRemaining / 28800), roundToTwo(consolidatedPODStatus[key].SPAchieved / consolidatedPODStatus[key].SPExpected), roundToTwo(consolidatedPODStatus[key].ETCRemaining / consolidatedPODStatus[key].ETCExpected) - 1, Math.round(consolidatedPODStatus[key].SPExpectedTillYesterday)]);
  }
}

function getPODNames(){
  let podNames = [];
  let cleanedPODNames = []
  let storiesMetaDataSheet = SpreadsheetApp.getActive().getSheetByName('Metadata');
  let metadata = storiesMetaDataSheet.getDataRange().getValues();
  let rowData;
  for(rowData in metadata){
    if(metadata[rowData][0] === "podNamesField"){
      podNames = metadata[rowData][1].split(";");
    }
  }
  podNames.forEach(entry => cleanedPODNames.push(entry.trim().toUpperCase()));
  console.log(podNames);
  return cleanedPODNames;
}