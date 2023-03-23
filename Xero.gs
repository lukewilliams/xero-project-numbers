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

function hasProjectNumber(projectName) {
  var expr = /^OE\s(\d{4})\s\|\s/;
  var result = projectName.match(expr);
  if (result) {
    return true
  }
  return false
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

function highestKnownProjectNumber() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("OE IDs");
  var data = sheet.getDataRange().getValues();

  let result = 0

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] > result) {
      result = data[i][0]
    }
  }

  return result
}

function newTrigger() {
  refreshToken()
  let projects = getProjects();

  // get highest number in the Sheet first
  let highestKnown = highestKnownProjectNumber()

  // Loop through all unknown projects:
  projects.items.filter(project => isProjectKnown(project.projectId) == false).forEach(project => {
    // First, mark it as discovered:
    project.discovered = true
    // Then get the client name.
    project.client = getClientName(project.contactId)
    let projectNumber
    if (hasProjectNumber(project.name)) {
      try {
        projectNumber = parseInt(project.name.match(/^OE\s(\d{4})\s\|\s/)[1])
      } catch(err) {
        Logger.log(err.message)
      }
      project.projectNumber = projectNumber
      project.needsProjectNumber = false
    } else {
      project.needsProjectNumber = true
    }
    if (projectNumber) {
      if (projectNumber > highestKnown) {
        highestKnown = projectNumber
      }
    }
  })

  let sheet = SpreadsheetApp.getActive().getSheetByName("OE IDs");

  // now we know the default project number to start on, and can add to the google sheet
  let startingNumber = highestKnown + 1

  projects.items.forEach((project) => {
    if (project.discovered) {
      if (project.needsProjectNumber) {
        project.projectNumber = startingNumber
        startingNumber++
      }

      sheet.appendRow([project.projectNumber, project.name, project.client, project.status, project.deadlineUtc, project.estimate.value, project.projectId])

    }
  })

  // now we go through and get the display name from A to prepend to the project.name for any project with the needsProjectNumber flag.
  // This will get sent to Xero. 

  projects.items.forEach((project) => {
    if (project.discovered) {
      let dispData = sheet.getDataRange().getDisplayValues()
      for (let i=1; i<dispData.length; i++) {
        if (project.projectId == dispData[i][6]) {
          if (project.needsProjectNumber) {
              let updatedName = dispData[i][0] + " | " + project.name 
              project.updatedName = updatedName
              // update Xero
              xero_updateProjectName(project)
              // update column B
              sheet.getRange(`B${i+1}`).setValue(updatedName)
          }
        }
      }
    }
  })
}

function xero_updateProjectName(projectData) {
  Logger.log("update xero: " + `${projectData.updatedName}, ${projectData.deadlineUtc}, ${projectData.estimate.value}`)

  let url = "https://api.xero.com/projects.xro/2.0/projects/" + projectData.projectId
  var raw = {
    "name": projectData.updatedName,
    "deadlineUtc": projectData.deadlineUtc,
    "estimateAmount": projectData.estimate.value
  };
  let headers = { 
    "Authorization": "Bearer " + getEnvironmentVariable('access_token'), 
    "Content-Type": "application/json", "Accept": "application/json",
    "xero-tenant-id": getEnvironmentVariable("xero-tenant-id"),
    "muteHttpExceptions": true,
  }
  try {
    let res = UrlFetchApp.fetch(url, {headers: headers, redirect: 'follow', payload: JSON.stringify(raw), method: 'put'})
    Logger.log(res)
  } catch(err) {
    Logger.log("Failed to update project name in xero for " + projectData.updatedName, err.message)
  }
}
