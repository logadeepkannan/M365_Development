const msalConfig = {
    auth: {
      clientId: "afa1af36-1e13-4a8b-bb5a-0d0bfba72bb2", 
      authority: "https://login.microsoftonline.com/059227af-dbff-4759-b2c3-8e7e7945fcdb", 
      redirectUri: "https://SharepointAPI.azurewebsite.net" 
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
  