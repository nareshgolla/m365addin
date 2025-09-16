Office.onReady((info) => {
  if (info.host === Office.HostType.Excel || info.host === Office.HostType.Word) {
    console.log("Add-in loaded in", info.host);
  }
});

function logoutUser() {
  console.log("Logging out...");
  // Add logout logic here
}

function closeFile() {
  if (Office.context.ui && Office.context.ui.closeContainer) {
    Office.context.ui.closeContainer();
  } else {
    console.log("closeContainer not available in this context.");
  } 
}

function onDocumentOpen(event) {
  Office.addin.showAsTaskpane();
  event.completed();
}
