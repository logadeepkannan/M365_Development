Office.initialize = function () {
  // Initialization logic, if needed
};

// This function is triggered before the email is sent
function beforeSendHandler(event) {
  const item = Office.context.mailbox.item;

  // Get the subject of the email
  item.subject.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
          const subjectText = result.value.toLowerCase(); // Convert to lowercase for case-insensitive check

          // Check if the subject contains the word 'project'
          if (!subjectText.includes("Project")) {
              // Block the Send operation and notify the user
              event.completed({
                  allowEvent: false // Block the send action
              });
              showNotification("Send Blocked", "The email cannot be sent because the subject does not contain the word 'project'.");
          } else {
              // Allow the Send operation
              event.completed({
                  allowEvent: true // Allow the email to be sent
              });
          }
      } else {
          // If unable to get the subject, allow sending by default
          event.completed({
              allowEvent: true
          });
      }
  });
}

// Function to show notification to the user
function showNotification(title, message) {
  Office.context.mailbox.item.notificationMessages.addAsync("restrictionNotice", {
      type: "informationalMessage",
      message: message,
      icon: "icon_16x16",
      persistent: false
  });
}