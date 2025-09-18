var customProps = {};

Office.onReady((info) => {
  Word.run(async (context) => {
    const documentProps = context.document.properties.customProperties;
    documentProps.load("items");
        context.document.save(); // saves current doc

    await context.sync();

    customProps = documentProps.items.reduce((map, p) => {
      map[p.key] = p.value;
      return map;
    }, {});

    onLoad();
  });
});

function deleteAllCookies() {
  const cookies = document.cookie.split("; ");
  cookies.forEach(cookie => {
    const [key] = cookie.split("=");
    document.cookie = `${key}=; path=/; SameSite=None; Secure; expires=Thu, 01 Jan 1970 00:00:00 GMT`;
  });
}

function onLoad() {
  invokeEndPoint(getEndPoint("T21_Open"), "Document ready for editing.", "Error occurred while opening document. Please login to the app and try again.", function(response) {
      deleteAllCookies();
    
      if (response && typeof response === "object") {
      Object.entries(response).forEach(([key, value]) => {
        document.cookie = `${key}=${value}; path=/; SameSite=None; Secure`;
      });
    }
  });
}

function closeFile() {
  invokeEndPoint(getEndPoint("T21_Close"), "Document closed.", "Error occurred. Please login to the app and try again.", function(response) {
    if (response && typeof response === "object") {
      Object.entries(response).forEach(([key, value]) => {
        document.cookie = `${key}=${value}; path=/; SameSite=None; Secure`;
      });
    }
  });
}

function uploadFile() {

    console.log("Hello There From Title21");

    //const url = getEndPoint("T21_Upload");
    //if (!url) {
    //    console.error("Upload URL not found in custom properties");
    //    return null;
    //}
    //Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, function (result) {
    //    if (result.status === Office.AsyncResultStatus.Succeeded) {
    //        var file = result.value;
    //        var sliceCount = file.sliceCount;
    //        var slicesReceived = 0, gotAllSlices = true, docdataSlices = [];

    //        for (var i = 0; i < sliceCount && gotAllSlices; i++) {
    //            file.getSliceAsync(i, function (result) {
    //                if (result.status === Office.AsyncResultStatus.Succeeded) {
    //                    docdataSlices.push(result.value.data);
    //                    slicesReceived++;

    //                    if (slicesReceived === sliceCount) {
    //                        // File reading complete, make AJAX request to upload file
    //                        var fileContent = new Uint8Array(docdataSlices[0].length * sliceCount);
    //                        for (var j = 0; j < docdataSlices.length; j++) {
    //                            fileContent.set(docdataSlices[j], j * docdataSlices[0].length);
    //                        }

    //                        var file = new File([fileContent], "document.docx", { type: getOfficeMimeType() });
    //                        uploadFileToServer(file, url, callback);
    //                    }
    //                } else {
    //                    gotAllSlices = false;
    //                }
    //            });
    //        }

    //        // Close the file handle
    //        file.closeAsync(function (closeResult) {
    //            if (closeResult.status === Office.AsyncResultStatus.Succeeded) {
    //                console.log("File handle closed successfully");
    //            } else {
    //                console.log("Error closing file handle");
    //            }
    //        });
    //    }
    //    else {
    //        showOfficePopup("Error reading file. Please login to the app and try again.");
    //    }
    //});
}

function uploadFileToServer(file, url, callback) {
  var formData = new FormData();
  formData.append("upload", file);

  fetch(url, {
    method: "POST",
    body: formData,
    credentials: "include"
  })
  .then(response => response.json())
  .then(data => {
    showOfficePopup("Upload successful.");
    callback(true);
  })
  .catch(error => {
      //onLoad();
      //uploadFile();
      const cookies = document.cookie.split("; ").reduce((acc, cookie) => {
  const [key, value] = cookie.split("=");
  acc[key] = value;
  return acc;
}, {});

console.log(cookies);
    showOfficePopup("Upload failed. Please login to the app, refresh the document and try again.");
    callback(false);
  });
}

function callback(status) {
    // Placeholder for upload callback
}

function callbackAfterInvokeEndPoint(status) {
    // Placeholder for fetc callback
}

function getEndPoint(key) {
  if (!customProps[key] || !customProps["T21_BaseUrl"]) {
    console.error("Missing custom property:", key);
    return "";
  }
  return customProps["T21_BaseUrl"] + customProps[key];
}

function getOfficeExtension() {
  switch (Office.context.host) {
    case Office.HostType.Word: return ".docx";
    case Office.HostType.Excel: return ".xlsx";
    case Office.HostType.PowerPoint: return ".pptx";
    default: return ".bin";
  }
}

function getOfficeMimeType() {
  switch (Office.context.host) {
    case Office.HostType.Word:
      return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    case Office.HostType.Excel:
      return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    case Office.HostType.PowerPoint:
      return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
    default:
      return "application/octet-stream";
  }
}

function invokeEndPoint(url, successMessage, failureMessage, callback) {
  fetch(url, { method: "GET" })
  .then(response => response.text())
  .then(text => {
    const data = tryParseJSON(text);
    if (data) {
      showOfficePopup(successMessage);
      callback(data);
    } else {
      showOfficePopup(successMessage);
      console.log("Not JSON:", text);
      callback(text);
    }
  })
  .catch(error => {
    showOfficePopup(failureMessage);
    console.log(error);
    callback(null);
  });
}

function tryParseJSON(str) {
  try {
    return JSON.parse(str);
  } catch {
    return null;
  }
}

function showOfficePopup(message) {
  return new Promise((resolve) => {
    const popup = document.getElementById("officePopup");
    const overlay = document.getElementById("officePopupOverlay");
    const msgElem = document.getElementById("officePopupMessage");
    const button = document.getElementById("officePopupButton");

    msgElem.innerText = message;
    popup.style.display = "block";
    overlay.style.display = "block";

    function closeHandler() {
      popup.style.display = "none";
      overlay.style.display = "none";
      button.removeEventListener("click", closeHandler);
      resolve();
    }

    button.addEventListener("click", closeHandler);
  });
}
