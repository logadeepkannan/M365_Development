Office.onReady((info)=> {
            if (info.host === Office.HostType.Outlook) {
                // Fetch and populate project values from SharePoint
                fetchProjectsFromSharePoint();

                // Add event listener to the dropdown
                document.getElementById("Placeholder").onchange = addProjectToSubject;
            }
        });


 // Wait for the iframe to load
 document.getElementById('myIframe').onload = function() {
    // Access the iframe's document
    var iframeDoc = document.getElementById('myIframe').contentWindow.document;
    // Get the placeholder value of an input element inside the iframe
    var placeholderValue = iframeDoc.querySelector('input').getAttribute('placeholder');
    console.log(placeholderValue);
};



    
       
        // Function to update the email subject based on selected project
        function addProjectToSubject() {
            const project = document.getElementById("Placeholder").value;

            // Ensure a valid project is selected before updating the subject
            if (project) {
                Office.context.mailbox.item.subject.setAsync(project, (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        console.log("Subject updated successfully.");
                    } else {
                        console.error(asyncResult.error.message);
                    }
                });
            } else {
                alert("Please select a project before sending the email.");
            }
        }