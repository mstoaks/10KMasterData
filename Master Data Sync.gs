/*
File: Master Data Sync
Author: Max Stoaks
Purpose: This script populates a google sheet with project data from our DS account at 10000ft.com
It gets basic project information as well as custom fields and tags from projects
Note - the target sheet must already created with the column headers and filters defined. This is because apps
script cannot create the filters, they must be manually created.

TODO:
*
-rename operational prioritization to DS prioritization etc
-remove bene impact from prio calc - DONE
*/


//*****************************************************************************************
//******************************************************************************************
//Main function to populate spreadsheet with data
//This function is called from the custom menu added in onOpen()
//******************************************************************************************
//******************************************************************************************/
function get10KProjectData() {

  //make sure we're on the right tab
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.setActiveSheet(ss.getSheetByName("Projects"));

  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("10K Integration", "Fetch project data from 10K? This will delete the data the active sheet and repopulate it with the latest data from 10K", ui.ButtonSet.OK_CANCEL);
  if (response != ui.Button.OK)
  {
    return;
  };

  //take care of setting the global params var in this method first
  props = PropertiesService.getScriptProperties();
  auth = props.getProperty('10KAuth');
  params = {
    'contentType':'application/json',
    'method': 'get',
    'headers': {
      'auth': auth,
      'Content-Type': 'application/json'
    }
  }

  //set up
  Logger.clear();
  clearActiveSheet(); //clear sheet

  //start the timer
  var startTime = Date.now();

  //get resources from 10K for lookup later
  getAllResources();

  //get all projects and populate with custom fields and labels
  getAllProjects();



  //create the array that will hold each project as an array of the project's attribute values
  var rangeValuesArray = new Array(allProjects.length);
  var x = 0;

  //loop through the new array and add each array-ified project to it
  for (y = 0; y < allProjects.length; y++) {
    //var tempObjArray = new Array();
    var p = allProjects[y];
    var tempObjArray=[p.name, p.link10K, p.description, p.programName, p.projectCode,
      p.miciti, p.strategicObjective, p.state, p.status, p.phase,
      p.externalStatus, p.activePhases, p.notes, p.recordMgr, p.effort,
      p.nonCompCost, p.forceGsbPrioritization, p.hiBeneImpact, p.gsbPrioStatus, p.gsbPrioFlag,
      p.operPrioStatus, p.operPrioFlag, p.inconsistencyMessage,
      p.director,
      p.projectManager, p.architectPartner, p.client, p.primaryContact, p.startDate,
      p.endDate,  p.team, p.beneficiaries, p.folder];

    rangeValuesArray[y] = tempObjArray;
  }

  var range = activeSheet.getRange(2, 1, allProjects.length, 33);
  range.setValues(rangeValuesArray);

  var endTime = Date.now();
  var totalTime = endTime - startTime;

  //update the sheet row colors
  colorActiveSheetRows();

  //finally update the timing message
  var rangeAll = activeSheet.getDataRange();
  var numCols = rangeAll.getNumColumns();
  var numRows = rangeAll.getNumRows();
  var user = Session.getActiveUser().getEmail();
  var updateTimeCell = activeSheet.getRange(1,numCols);
  updateTimeCell.setValue("Last data sync : " + new Date().toString() + " ---  milliseconds: " + totalTime + " by:" + user);


}



/************************************************************
//get all projects from 10K and process into local objects
************************************************************/
function getAllProjects() {
  //get all the projects (up to 500)
  allProjects = new Array();
  var response = UrlFetchApp.fetch('https://api.10000ft.com/api/v1/projects?per_page=500&fields=custom_field_values,tags,phase_count', params);

  var jsonProjects = JSON.parse(response.getContentText()); //JSONified response


  //now we'll loop over each element in the jsonProject.data structure - each one is a project
  jsonProjects.data.forEach(function(element){ //this is also a good place to limit what to process. i.e. if element.state != internal
    var proj = new Object();
    proj.id = element.id;
    proj.link10K = "=HYPERLINK(" + "\"https://app.10000ft.com/viewproject?id=" + element.id +"\"" + "," + "\"Link\")";
    proj.name = element.name;
    proj.startDate = element.starts_at;
    proj.endDate = element.ends_at;
    proj.description = element.description;
    proj.projectCode = element.project_code;
    proj.client = element.client;
    proj.state = element.project_state;
    proj.phaseCount = element.phase_count;
    proj.beneficiaries = "";
    //get the custom fields
    element.custom_field_values.data.forEach(function(custFieldVal) {
      switch (custFieldVal.custom_field_id) {
        case PROJ_PHASE_ID: proj.phase = custFieldVal.value;
        break;
        case PROJ_MGR_ID: proj.projectManager = custFieldVal.value;
        break;
        case PROJ_STRAT_OBJ_ID: proj.strategicObjective = custFieldVal.value;
        break;
        case PROJ_ARCH_PARTNER_ID: proj.architectPartner = custFieldVal.value;
        break;
        case PROJ_STATUS_ID: proj.status = custFieldVal.value;
        break;
        case PROJ_BENEFICIARY: proj.beneficiaries += custFieldVal.value + " ";
        break;
        case PROJ_PARENT_PGM: proj.programID = custFieldVal.value;
        break;
        case PROJ_DIRECTOR: proj.director = custFieldVal.value;
        break;
        case PROJ_GSB_PRIO: proj.gsbPrioStatus = custFieldVal.value;
        break;
        case PROJ_EFFORT: proj.effort = custFieldVal.value;
        break;
        case PROJ_VALUE: proj.valueToSchool = custFieldVal.value;
        break;
        case PROJ_MICITI: proj.miciti = custFieldVal.value;
        break;
        case PROJ_NOTES: proj.notes = custFieldVal.value;
        break;
        case PROJ_PRMY_CONTACT: proj.primaryContact = custFieldVal.value;
        break;
        case PROJ_NONCOMP_COST: proj.nonCompCost = custFieldVal.value;
        break;
        case PROJ_HI_BENE_IMPACT: proj.hiBeneImpact = custFieldVal.value;
        break;
        case PROJ_FORCE_PRIO: proj.forceGsbPrioritization = custFieldVal.value;
        break;
        case PROJ_FOLDER:
          //proj.folder = custFieldVal.value;
          proj.folder = "=HYPERLINK(" + "\"" + custFieldVal.value + "\"" + ",\"Link\")";
        break;
        case PROJ_PMO_DECK: proj.pmoDeck = custFieldVal.value;
        break;
        case PROJ_RECORD_MGR: proj.recordMgr = custFieldVal.value;
        break;
        case PROJ_OP_PRIO: proj.operPrioStatus = custFieldVal.value;
        break;
        default:Logger.log("default reached in custom field line 147 or therabouts");
      };
    });

    //get the project tags
    proj.tags = "";
    element.tags.data.forEach(function(element) {
      proj.tags += element.value;
    });

    //derive the external status
    if (proj.phase == "Pre-concept" || proj.phase == "Concept") {proj.externalStatus = "Exploration"}
    else if (proj.phase == "High Level Design" || proj.phase == "Pitch") {proj.externalStatus = "Queued"}
    else {proj.externalStatus = "Active"};

    /////////////////START PRIORITIZATION STUFF////////////////////

    // Set some defaults
    proj.gsbPrioFlag = false;
    proj.operPrioFlag = false;
    proj.inconsistencyErrors = [];




    if (proj.state == "Internal") {
      if (isPrioritized(proj.gsbPrioStatus) || isPrioritized(proj.operPrioStatus)) {
        proj.inconsistencyErrors.push("Project is internal so it shouldn't be GSB or DS prioritized.");
      }
    }
    // Derive prioritization
    else {
      //derive if project should get gsb level prioritization
      if (proj.forceGsbPrioritization == "Yes" || levelHigherThan(proj.effort, 'Low') || levelHigherThan(proj.nonCompCost, 'Low')){proj.gsbPrioFlag = true}

      // Determine inconsistency

      // If it is currently not set to be prioritized.
      if (proj.gsbPrioStatus == "No" || proj.gsbPrioStatus == "TBD") {
        // Check Effort
        if (proj.effort == "Medium" || proj.effort == "High") {
          proj.inconsistencyErrors.push("Effort is greater than Low, but it is not set to be GSB prioritized.");
        }

        // Check Non-Comp Cost
        if (proj.nonCompCost == "Medium" || proj.nonCompCost == "High") {
          proj.inconsistencyErrors.push("Non-Comp Cost is greater than Low, but it is not set to be GSB prioritized.");
        }

        // Check Force Prioritization.
        if (proj.forceGsbPrioritization == "Yes") {
          proj.inconsistencyErrors.push("Force prioritization is selected, but it is not set to be GSB prioritized.");
        }
      }
      // If we calculate that it shouldn't be prioritized but it is.
      else if (!proj.gsbPrioFlag) {
        proj.inconsistencyErrors.push("Nothing warrants it to be prioritized, but it is set to be GSB prioritized.");
      }

      //derive DS prioritization
      if (!proj.gsbPrioFlag) {
        if (proj.effort == "High" || proj.effort == "Medium" || proj.effort == "Low" || proj.nonCompCost == "High" || proj.nonCompCost == "Medium" || proj.nonCompCost == "Low"){proj.operPrioFlag = true}

        // Check for inconsistency
        if (proj.operPrioStatus == "No" || proj.operPrioStatus == "TBD") {
          // Check Effort
          if (proj.effort == "Low") {
            proj.inconsistencyErrors.push("Effort is Low, but it is not set to be DS prioritized.");
          }

          // Check Non-Comp Cost
          if (proj.nonCompCost == "Low") {
            proj.inconsistencyErrors.push("Non-Comp Cost is Low, but it is not set to be DS prioritized.");
          }
        }
      }
      else {
        // Check for consistency
        if (proj.operPrioStatus != "No") {
          proj.inconsistencyErrors.push("It is GSB prioritized so DS prioritization should be No.")
        }
      }
    }


    // If there are no errors set the message to none.
    // Otherwise join the errors together with new lines and * as bullet points.
    if (!proj.inconsistencyErrors.length) {
      proj.inconsistencyMessage = '-None-';
    }
    else {
      proj.inconsistencyMessage = "* " + proj.inconsistencyErrors.join("\n\n* ");
    }

    //////////////////////END PRIORITIZATION STUFF//////////////////////////

    //get current resources
    proj.team = getStringifiedProjectResources(proj.id);


    //if the project has phases then get the current ones (if any)
    if (proj.phaseCount > 0) {
      proj.activePhases = getStringifiedCurrentPhases(proj.id);
    }
    else {
      //get the phase from the DS phase mapping
      proj.activePhases = "None";
    }
    //add the project to the list
    allProjects.push(proj);
  });

  //order the projectDetailedObjects in alpha order of element.name
  allProjects.sort(function(a,b) {
    var nameA = a.name.toUpperCase();
    var nameB = b.name.toUpperCase();
    return (nameA < nameB) ? -1 : (nameA > nameB) ? 1 : 0;
  });

  //now that we have all the projects with names in an array we can populate the programName if any
  allProjects.forEach(function(p){
    if (p.programID) {
        p.programName = getProjectName(p.programID);
      }
   else {
        p.programName = "None";
      }
  })


}


/************************************************************
//get *current* phases for a project
************************************************************/
function getStringifiedCurrentPhases(projectID) {


  date= getTodayString(); //might need to look from yesterday until tomorrow instead of just today

  var response = UrlFetchApp.fetch('https://api.10000ft.com/api/v1/projects/' + projectID + '/phases?per_page=100&from='+date+'&to='+date, params);

  var jsonPhases = JSON.parse(response.getContentText()); //JSON response

  if (jsonPhases.data.length == 0) {
    //no current 10k phases so derive what it is
    return "None";
  }
  else {
  //have at least one current 10k phase
    var stringifiedPhases = "";

    jsonPhases.data.forEach(function(element){
      stringifiedPhases = stringifiedPhases + element.phase_name + " ";
    });

    //should I sort these? not sure, we'll see how it looks and ask DLS

    return stringifiedPhases;
  }
}


/**************************************************************************************************
//return a sorted string of folks *currently* assigned to a project
**************************************************************************************************/
function getStringifiedProjectResources(projectID){

  date= getTodayString(); //might need to look from yesterday until tomorrow instead of just today

  var response = UrlFetchApp.fetch('https://api.10000ft.com/api/v1/projects/' + projectID + '/assignments?per_page=100&from='+date+'&to='+date+'&with_phases=true', params);

  var jsonResources = JSON.parse(response.getContentText()); //JSON response


  var projectResources = new Array();


  jsonResources.data.forEach(function(element){
    var user= new Object();
    user.id = element.user_id;
    user.percent = element.percent;

    var person = getPerson(user.id);
    /*see if we actually got a person back. in 10K archived people are not returned from the /users API even though they are still in the assignments.
    so we check */
    if (person) {
      user.first = person.first;
      user.last = person.last;
      projectResources.push(user);
    }
  });

  //sort em by first name
  projectResources.sort(function(a, b) {
    var nameA = a.first.toUpperCase(); // ignore upper and lowercase
    var nameB = b.first.toUpperCase(); // ignore upper and lowercase
    if (nameA < nameB) {
     return -1;
    }
    if (nameA > nameB) {
      return 1;
    }
    // names must be equal
     return 0;
  });

  var stringifiedResources= "";
  projectResources.forEach(function(element) {
    stringifiedResources = stringifiedResources + element.first + " " + element.last + " [" + element.percent*100 +"%]" +  ", ";
  });
  //stringifiedUsers.trim(); //trim() doesn't seem to be supported in GAS
  stringifiedResources = stringifiedResources.slice(0, -2);
  return stringifiedResources;
}


/*********************************************************************************************************
//function to get all resources in order to get assigned team members
*********************************************************************************************************/

function getAllResources(){
  //Logger.log("in getAllResources()");
  allResources = new Array();

  var response = UrlFetchApp.fetch('https://api.10000ft.com/api/v1/users?per_page=500&fields=custom_field_values,tags', params);

  allRes = JSON.parse(response.getContentText()); //JSON response


  //array to hold each javascript project object
  allRes.data.forEach(function(element){
    var user= new Object();
    user.id = element.id;
    user.first = element.first_name;
    user.last = element.last_name;
    allResources.push(user);
  });
}



/*************************************************************************************************
//function to get a person from the allResources array - Array.find() not support in GAS
**************************************************************************************************/
function getPerson(pid) {
  for (i = 0; i< allResources.length; i++) {
    tempPerson = allResources[i];
    if (tempPerson.id == pid) {
      return tempPerson;
    }
  }
}


/*****************************************************************************************************
// function to get a project's name out of the allProjects array - GAS does not support Array.find()
******************************************************************************************************/
function getProjectName(pid) {
  for (var i = 0; i < allProjects.length; i++) {
    if (allProjects[i].id == pid) {
      return allProjects[i].name;
    }
  }
  //throw "There's a problem on deck " + pid + " Mr Sulu. Tried to get a project out of allProjects with a project ID and didn't find it. That shouldn't happen. Ever. function getProjectName(). Did someone enter a program ID that doesn't exist?";
  Logger.log("Could not find a project for PID = " + pid + " in getProjectName(pid). Check that projects have correct parent ids setup.");
  return "Unknown Program";
}


/*************************************************************************************************
//function to get a date string in a format 10K api can understand
**************************************************************************************************/
function getTodayString() {
  var d = new Date();
  var month = '' + (d.getMonth() + 1);
  var day = '' + d.getDate();
  var year = d.getFullYear();

  if (month.length < 2) {month = '0' + month};
  if (day.length < 2) {day = '0' + day};

  return([year, month, day].join('-'));
}

/*********************************************
//utility function to clear spreadsheet
*********************************************/
function clearActiveSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.setActiveSheet(ss.getSheetByName("Projects"));

  /*
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("10K Integration", "Clear data? This will delete all data under the headers (10K unaffected)", ui.ButtonSet.OK_CANCEL);
  if (response != ui.Button.OK)
  {
    return;
  }
  */
  //clear out the old data, leaving the header row so that we don't lose the filter
  rangeAll = activeSheet.getDataRange();
  numCols = rangeAll.getNumColumns();
  numRows = rangeAll.getNumRows();

  if (numRows > 1) {
    rangeToClear = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(2,1,numRows-1, numCols);
    rangeToClear.clearContent();
  }

}

/***********************************************************************
//utility function to change background colors of alternating rows
************************************************************************/
function colorActiveSheetRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.setActiveSheet(ss.getSheetByName("Projects"));

  var rangeAll = activeSheet.getDataRange();
  var numCols = rangeAll.getNumColumns();
  var numRows = rangeAll.getNumRows();

  if (numRows > 1) {
    var rangeToColor = activeSheet.getRange(2,1,numRows-1, numCols);
    rangeToColor.setBackground("#FFFFFF");
    for (var i=3; i <= rangeAll.getNumRows(); i+=2){
      activeSheet.getRange(i, 1, 1, numCols).setBackground("#F3F3F3");
    }
  }
}


/**********************************************
//utility function to see if it is prioritized.
***********************************************/
function isPrioritized(status) {
  return (status == "Prioritize - Approved" || status == "Prioritized - Deferred" || status == "Prioritized - Denied" || status == "To Be Prioritized");
}

/**********************************************
//utility function to convert level to a number
***********************************************/
function levelNumber(level) {
  levelNumber = 0;
  switch(level) {
    case "Very Low":
      levelNumber = 1;
      break;
    case "Low":
      levelNumber = 2;
      break;
    case "Medium":
      levelNumber = 3;
      break;
    case "High":
      levelNumber = 4;
      break;
  }

  return levelNumber;
}

function levelHigherThan(level, wantedLevel) {
  isHigher = false;
  switch (wantedLevel) {
    case "Very Low":
      isHigher = (level == "Low" || level == "Medium" || level == "High") ? true : false;
      break;
    case "Low":
      isHigher = (level == "Medium" || level == "High") ? true : false;
      break;
    case "Medium":
      isHigher = (level == "High") ? true : false;
      break;
  }

  return isHigher;
}
