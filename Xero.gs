function setEnvironmentVariable(key, value) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(key, value);
}

function getEnvironmentVariable(key) {
  const scriptProperties = PropertiesService.getScriptProperties();
  try {
    return scriptProperties.getProperty(key);
  } catch (err) {
    Logger.log("Error retrieving " + key + " from script properties: " + err.message);
  }
}

function refreshToken() {
  let url = "https://identity.xero.com/connect/token?="
  
  let formData = {
    'grant_type': 'refresh_token',
    'refresh_token': getEnvironmentVariable('refresh_token'),
    'client_id': getEnvironmentVariable('client_id'),
    'client_secret': getEnvironmentVariable('client_secret')
  }

  let res = UrlFetchApp.fetch(url, {'method': 'post', 'payload': formData})

  let data = JSON.parse(res.getContentText())
  setEnvironmentVariable("access_token", data.access_token)
  setEnvironmentVariable("refresh_token", data.refresh_token)
}

function getProjects() {
  let headers = { 
    "Authorization": "Bearer " + getEnvironmentVariable('access_token'), 
    "Content-Type": "application/json", "Accept": "application/json",
    "xero-tenant-id": getEnvironmentVariable("xero-tenant-id"),
  }
  let res = UrlFetchApp.fetch("https://api.xero.com/projects.xro/2.0/projects", {headers: headers, redirect: 'follow'})
  let projects = JSON.parse(res.getContentText())
  return projects
}

function getClientName(contactId) {
  let url = "https://api.xero.com/api.xro/2.0/Contacts/" + contactId
  let headers = { 
    "Authorization": "Bearer " + getEnvironmentVariable('access_token'), 
    "Content-Type": "application/json", "Accept": "application/json",
    "xero-tenant-id": getEnvironmentVariable("xero-tenant-id"),
  }
  let res = UrlFetchApp.fetch(url, {headers: headers, redirect: 'follow'})
  let contact = JSON.parse(res.getContentText())
  return contact.Contacts[0]["Name"]
}

function updateProject(projectId, existingName, deadlineUtc, estimateAmount, contactId) {
  // get the new OE number
  var sheet = SpreadsheetApp.getActive().getSheetByName("OE IDs");
  var dispData = sheet.getDataRange().getDisplayValues()
  var rawData = sheet.getDataRange().getValues()
  let oeId;
  for (var i = 1; i < dispData.length; i++) {
    if (projectId == dispData[i][6]) {
      oeId = dispData[i][0];
      // TODO: Check that name doesn't already include an OE number
      sheet.getRange(`B${i+1}`).setValue(oeId + " | " + existingName);
      // persist the OE number in sheets, respecting the one below it
      sheet.getRange(`A${i+1}`).setValue(rawData[i][0]);
      // get the client name
      let client = getClientName(contactId)
      sheet.getRange(`C${i+1}`).setValue(client);
    }
  }
  Logger.log(oeId)
  // tell xero the new name
  updateProjectNameInXero(projectId, oeId + " | " + existingName, deadlineUtc, estimateAmount)
}

function updateProjectNameInXero(projectId, updatedName, deadlineUtc, estimateAmount) {
  let url = "https://api.xero.com/projects.xro/2.0/projects/" + projectId
  var raw = {
    "name": updatedName,
    "deadlineUtc": deadlineUtc,
    "estimateAmount": estimateAmount
  };
  let headers = { 
    "Authorization": "Bearer " + getEnvironmentVariable('access_token'), 
    "Content-Type": "application/json", "Accept": "application/json",
    "xero-tenant-id": getEnvironmentVariable("xero-tenant-id"),
    "muteHttpExceptions": true,
  }
  let res = UrlFetchApp.fetch(url, {headers: headers, redirect: 'follow', payload: JSON.stringify(raw), method: 'put'})
  Logger.log(res)
}

function trigger_FindNewProjects() {
  refreshToken();
  let projects = getProjects();

  projects.items.forEach((project) => {
    if (isProjectKnown(project.projectId) == false) {
      addProject(project.projectId, project.name, "", project.status, project.deadlineUtc, project.estimate.value, project.projectId)

      updateProjectNameInXero(project.projectId, project.name, project.deadlineUtc, project.estimate.value, project.contactId)
    }
  })
}

function isProjectKnown(projectId) {
  var sheet = SpreadsheetApp.getActive().getSheetByName("OE IDs");
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (projectId == data[i][6]) {
      return true;
    }
  }
  return false;
}

function addProject(projectId, name, client, status, deadlineUtc, estimateAmount) {
  let sheet = SpreadsheetApp.getActive().getSheetByName("OE IDs");
  sheet.appendRow(["=INDIRECT(ADDRESS(ROW()-1,COLUMN()))+1",name,client,status,deadlineUtc,estimateAmount,projectId])
}
