const msalConfig = {
    auth: {
      clientId: "afa1af36-1e13-4a8b-bb5a-0d0bfba72bb2", 
      authority: "https://login.microsoftonline.com/059227af-dbff-4759-b2c3-8e7e7945fcdb", 
      redirectUri: "https://ittridentsqa.onmicrosoft.com" 
    }
  };
  
  // Create an instance of MSAL PublicClientApplication
  const msalInstance = new msal.PublicClientApplication(msalConfig);
  
  // Login request configuration
  const loginRequest = {
    scopes: ["https://ittridentsqa.sharepoint.com/.default"] // The scope for SharePoint access
  };
  
  // Function to get an access token using MSAL.js
  function getAccessToken() {
    return msalInstance.acquireTokenSilent(loginRequest).then(response => {
      return response.accessToken;
    }).catch(error => {
      // If silent token acquisition fails, initiate login flow
      return msalInstance.loginPopup(loginRequest).then(response => {
        return response.accessToken;
      }).catch(loginError => {
        console.error("Login failed: ", loginError);
      });
    });
  }
  
  // Function to fetch SharePoint list items and populate the dropdown
  function populateDropdown() {
    var projectDropdown = document.getElementById("projectDropdown");
  
    // Get access token before calling SharePoint REST API
    getAccessToken().then(accessToken => {
      var sharePointListUrl = "https://ittridentsqa.sharepoint.com/sites/TridentSQAM365InternalSolution/_api/web/lists/getbytitle('Project_Dropdown')/items";
  
      // Use Fetch API to get list items
      fetch(sharePointListUrl, {
        method: "GET",
        headers: {
          "Accept": "application/json;odata=verbose",
          "Authorization": "Bearer " + accessToken 
        }
      })
        .then(response => response.json())
        .then(data => {
          while (projectDropdown.options.length > 1) {
            projectDropdown.remove(1);
          }
  
          
          data.d.results.forEach(item => {
            var option = document.createElement("option");
            option.value = item.Title;  
            option.text = item.Title;
            projectDropdown.add(option);
          });
        })
        .catch(error => console.error("Error fetching SharePoint list: ", error));
    });
  }
  

  function addProjectToSubject() {
    const project = document.getElementById("projectDropdown").value;
  

    if (project && project !== "Select a project") {
      Office.context.mailbox.item.subject.setAsync(project, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Subject updated successfully.");
        } else {
          console.error(asyncResult.error.message);
        }
      });
    } else {
      console.log("Please select a valid project.");
    }
  }
  
 
  Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      populateDropdown();
      document.getElementById("projectDropdown").onchange = addProjectToSubject;
    }
  });

  function validateSubjectAndCC(event) {
    shouldChangeSubjectOnSend(event);
}

function shouldChangeSubjectOnSend(event) {
    mailboxItem.subject.getAsync(
        { asyncContext: event },
        function (asyncResult) {
            addCCOnSend(asyncResult.asyncContext);
            //console.log(asyncResult.value);
            // Match string.
            var checkSubject = (new RegExp(/\[Checked\]/)).test(asyncResult.value)
            // Add [Checked]: to subject line.
            subject = '[Checked]: ' + asyncResult.value;

            // Check if a string is blank, null or undefined.
            // If yes, block send and display information bar to notify sender to add a subject.
            if (asyncResult.value === null || (/^\s*$/).test(asyncResult.value)) {
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Please enter a subject for this email.' });
                asyncResult.asyncContext.completed({ allowEvent: false });
            }
            else {
                // If can't find a [Checked]: string match in subject, call subjectOnSendChange function.
                if (!checkSubject) {
                    subjectOnSendChange(subject, asyncResult.asyncContext);
                    //console.log(checkSubject);
                }
                else {
                    // Allow send.
                    asyncResult.asyncContext.completed({ allowEvent: true });
                }
            }

        }
      )
}

function subjectOnSendChange(subject, event) {
    mailboxItem.subject.setAsync(
        subject,
        { asyncContext: event },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Unable to set the subject.' });

                // Block send.
                asyncResult.asyncContext.completed({ allowEvent: false });
            }
            else {
                // Allow send.
                asyncResult.asyncContext.completed({ allowEvent: true });
            }

        });
}
  