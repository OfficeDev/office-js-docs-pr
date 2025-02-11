---
title: Get the whole document from an add-in for PowerPoint or Word
description: Learn to get the whole document from a PowerPoint or Word add-in.
ms.date: 02/12/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Get the whole document from an add-in for PowerPoint or Word

You can create an Office Add-in to send or publish a PowerPoint presentation or Word document to a remote location. This article demonstrates how to build a simple task pane add-in for PowerPoint or Word that gets all of the presentation or document as a data object and sends that data to a web server via an HTTP request.

## Prerequisites for creating an add-in for PowerPoint or Word

This article assumes that you are using a text editor to create the task pane add-in for PowerPoint or Word. To create the task pane add-in, you must create the following files.

- On a shared network folder or on a web server, you need the following files.

  - An HTML file (**GetDoc_App.html**) that contains the user interface plus links to the JavaScript files (including Office.js and application-specific .js files) and Cascading Style Sheet (CSS) files.

  - A JavaScript file (**GetDoc_App.js**) to contain the programming logic of the add-in.

  - A CSS file (**Program.css**) to contain the styles and formatting for the add-in.

- A manifest file (**GetDoc_App.xml** or **GetDoc_App.json**) for the add-in, available on a shared network folder or add-in catalog. The manifest file must point to the location of the HTML file mentioned previously.

Alternatively, you can create an add-in for your Office application using one of the following options. You won't have to create new files as the equivalent of each required file will be available for you to update. For example, the Yeoman generator options include **./src/taskpane/taskpane.html**, **./src/taskpane/taskpane.js**, **./src/taskpane/taskpane.css**, and **./manifest.xml**.

- PowerPoint
  - [Visual Studio](../quickstarts/powerpoint-quickstart-vs.md)
  - [Yeoman generator for Office Add-ins](../quickstarts/powerpoint-quickstart-yo.md)
- Word
  - [Visual Studio](../quickstarts/word-quickstart-vs.md)
  - [Yeoman generator for Office Add-ins](../quickstarts/word-quickstart-yo.md)

### Core concepts to know for creating a task pane add-in

Before you begin creating this add-in for PowerPoint or Word, you should be familiar with building Office Add-ins and working with HTTP requests. This article doesn't discuss how to decode Base64-encoded text from an HTTP request on a web server.

## Create the manifest for the add-in

The manifest file for an Office Add-in provides important information about the add-in: what applications can host it, the location of the HTML file, the add-in title and description, and many other characteristics.

In a text editor, add the following code to the manifest file. If you're using a Visual Studio project, select the "Add-in only manifest" option.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

> [!NOTE]
> The unified manifest is generally available for production Outlook add-ins. It's available only for preview in Excel, PowerPoint, and Word add-ins.

```json
{
    "$schema": "https://developer.microsoft.com/json-schemas/teams/vDevPreview/MicrosoftTeams.schema.json#",
    "manifestVersion": "devPreview",
    "version": "1.0.0.0",
    "id": "[Replace_With_Your_GUID]",
    "localizationInfo": {
        "defaultLanguageTag": "en-us"
    },
    "developer": {
        "name": "[Provider Name e.g., Contoso]",
        "websiteUrl": "[Insert the URL for the app e.g., https://www.contoso.com]",
        "privacyUrl": "[Insert the URL of a page that provides privacy information for the app e.g., https://www.contoso.com/privacy]",
        "termsOfUseUrl": "[Insert the URL of a page that provides terms of use for the app e.g., https://www.contoso.com/servicesagreement]"
    },
    "name": {
        "short": "Get Doc add-in",
        "full": "Get Doc add-in"
    },
    "description": {
        "short": "My get PowerPoint or Word document add-in.",
        "full": "My get PowerPoint or Word document add-in."
    },
    "icons": {
        "outline": "_layouts/images/general/office_logo.jpg",
        "color": "_layouts/images/general/office_logo.jpg"
    },
    "accentColor": "#230201",
    "validDomains": [
        "https://www.contoso.com"
    ],
    "showLoadingIndicator": false,
    "isFullScreen": false,
    "defaultBlockUntilAdminAction": false,
    "authorization": {
        "permissions": {
            "resourceSpecific": [
                {
                    "name": "Document.ReadWrite.User",
                    "type": "Delegated"
                }
            ]
        }
    },
    "extensions": [
        {
            "requirements": {
                "scopes": [
                    "document",
                    "presentation"
                ]
            },
            "alternates": [
                {
                    "alternateIcons": {
                        "icon": {
                            "size": 32,
                            "url": "http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg"
                        },
                        "highResolutionIcon": {
                            "size": 64,
                            "url": "http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg"
                        }
                    }
                }
            ]
        }
    ]
}
```

# [Add-in only manifest](#tab/xmlmanifest)

```xml
<?xml version="1.0" encoding="utf-8" ?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
xsi:type="TaskPaneApp">
    <Id>[Replace_With_Your_GUID]</Id>
    <Version>1.0</Version>
    <ProviderName>[Provider Name]</ProviderName>
    <DefaultLocale>EN-US</DefaultLocale>
    <DisplayName DefaultValue="Get Doc add-in" />
    <Description DefaultValue="My get PowerPoint or Word document add-in." />
    <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" />
    <HighResolutionIconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" />
    <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
    <Hosts>
        <Host Name="Document" />
        <Host Name="Presentation" />
    </Hosts>
    <DefaultSettings>
        <SourceLocation DefaultValue="[Network location of app]/GetDoc_App.html" />
    </DefaultSettings>
    <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

Save the file as **GetDoc_App.xml** using UTF-8 encoding to a network location or to an add-in catalog.

---

## Create the user interface for the add-in

For the user interface of the add-in, you can use HTML written directly into the **GetDoc_App.html** file. The programming logic and functionality of the add-in must be contained in a JavaScript file (for example, **GetDoc_App.js**).

Use the following procedure to create a simple user interface for the add-in that includes a heading and a single button.

1. In a new file in the text editor, add the HTML for your selected Office application.

    ### [PowerPoint](#tab/powerpoint)

    ```html
    <!DOCTYPE html>
    <html>
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
            <title>Publish presentation</title>
            <link rel="stylesheet" type="text/css" href="Program.css" />
            <script src="https://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js" type="text/javascript"></script>
            <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
            <script src="GetDoc_App.js"></script>
        </head>
        <body>
            <form>
                <h1>Publish presentation</h1>
                <br />
                <div><input id='submit' type="button" value="Submit" /></div>
                <br />
                <div><h2>Status</h2>
                    <div id="status"></div>
                </div>
            </form>
        </body>
    </html>
    ```

    ### [Word](#tab/word)

    ```html
    <!DOCTYPE html>
    <html>
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
            <title>Publish document</title>
            <link rel="stylesheet" type="text/css" href="Program.css" />
            <script src="https://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js" type="text/javascript"></script>
            <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
            <script src="GetDoc_App.js"></script>
        </head>
        <body>
            <form>
                <h1>Publish document</h1>
                <br />
                <div><input id='submit' type="button" value="Submit" /></div>
                <br />
                <div><h2>Status</h2>
                    <div id="status"></div>
                </div>
            </form>
        </body>
    </html>
    ```

    ---

1. Save the file as **GetDoc_App.html** using UTF-8 encoding to a network location or to a web server.

    > [!NOTE]
    > Be sure that the **head** tags of the add-in contains a **script** tag with a valid link to the Office.js file.

1. We'll use some CSS to give the add-in a simple yet modern and professional appearance. Use the following CSS to define the style of the add-in.

    In a new file in the text editor, add the following CSS.

    ```css  
    body
    {
        font-family: "Segoe UI Light","Segoe UI",Tahoma,sans-serif;
    }
    h1,h2
    {
        text-decoration-color:#4ec724;
    }
    input [type="submit"], input[type="button"]
    {
        height:24px;
        padding-left:1em;
        padding-right:1em;
        background-color:white;
        border:1px solid grey;
        border-color: #dedfe0 #b9b9b9 #b9b9b9 #dedfe0;
        cursor:pointer;
    }
    ```

1. Save the file as **Program.css** using UTF-8 encoding to the network location or to the web server where the **GetDoc_App.html** file is located.

## Add the JavaScript to get the document

In the code for the add-in, a handler to the [Office.initialize](/javascript/api/office#office-office-initialize-function(1)) event adds a handler to the click event of the **Submit** button on the form and informs the user that the add-in is ready.

The following code example shows the event handler for the `Office.initialize` event along with a helper function, `updateStatus`, for writing to the status div.

```js
// The initialize or onReady function is required for all add-ins.
Office.initialize = function (reason) {

    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {

        // Run sendFile when Submit is clicked.
        $('#submit').on("click", function () {
            sendFile();
        });

        // Update status.
        updateStatus("Ready to send file.");
    });
}

// Create a function for writing to the status div.
function updateStatus(message) {
    var statusInfo = $('#status');
    statusInfo[0].innerHTML += message + "<br/>";
}
```

When you choose the **Submit** button in the UI, the add-in calls the `sendFile` function, which contains a call to the [Document.getFileAsync](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) method. The `getFileAsync` method uses the asynchronous pattern, similar to other methods in the Office JavaScript API. It has one required parameter, _fileType_, and two optional parameters,  _options_ and _callback_.

The _fileType_ parameter expects one of three constants from the [FileType](/javascript/api/office/office.filetype) enumeration: `Office.FileType.Compressed` ("compressed"), `Office.FileType.PDF` ("pdf"), or `Office.FileType.Text` ("text"). The current file type support for each platform is listed under the [Document.getFileType](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)) remarks. When you pass in **Compressed** for the _fileType_ parameter, the `getFileAsync` method returns the current document as a PowerPoint presentation file (\*.pptx) or Word document file (\*.docx) by creating a temporary copy of the file on the local computer.

The `getFileAsync` method returns a reference to the file as a [File](/javascript/api/office/office.file) object. The `File` object exposes the following four members.

- [size](/javascript/api/office/office.file#office-office-file-size-member) property
- [sliceCount](/javascript/api/office/office.file#office-office-file-slicecount-member) property
- [getSliceAsync](/javascript/api/office/office.file#office-office-file-getsliceasync-member(1)) method
- [closeAsync](/javascript/api/office/office.file#office-office-file-closeasync-member(1)) method

The `size` property returns the number of bytes in the file. The `sliceCount` returns the number of [Slice](/javascript/api/office/office.slice) objects (discussed later in this article) in the file.

Use the following code to get the current PowerPoint or Word document as a `File` object using the `Document.getFileAsync` method and then make a call to the locally defined `getSlice` function. Note that the `File` object, a counter variable, and the total number of slices in the file are passed along in the call to `getSlice` in an anonymous object.

```js
// Get all of the content from a PowerPoint or Word document in 100-KB chunks of text.
function sendFile() {
    Office.context.document.getFileAsync("compressed",
        { sliceSize: 100000 },
        function (result) {

            if (result.status === Office.AsyncResultStatus.Succeeded) {

                // Get the File object from the result.
                var myFile = result.value;
                var state = {
                    file: myFile,
                    counter: 0,
                    sliceCount: myFile.sliceCount
                };

                updateStatus("Getting file of " + myFile.size + " bytes");
                getSlice(state);
            } else {
                updateStatus(result.status);
            }
        });
}
```

The local function `getSlice` makes a call to the `File.getSliceAsync` method to retrieve a slice from the `File` object. The `getSliceAsync` method returns a `Slice` object from the collection of slices. It has two required parameters, _sliceIndex_ and _callback_. The  _sliceIndex_ parameter takes an integer as an indexer into the collection of slices. Like other methods in the Office JavaScript API, the `getSliceAsync` method also takes a callback function as a parameter to handle the results from the method call.

The `Slice` object gives you access to the data contained in the file. Unless otherwise specified in the _options_ parameter of the `getFileAsync` method, the `Slice` object is 4 MB in size. The `Slice` object exposes three properties: [size](/javascript/api/office/office.slice#office-office-slice-size-member), [data](/javascript/api/office/office.slice#office-office-slice-data-member), and [index](/javascript/api/office/office.slice#office-office-slice-index-member). The `size` property gets the size, in bytes, of the slice. The `index` property gets an integer that represents the slice's position in the collection of slices.

```js
// Get a slice from the file and then call sendSlice.
function getSlice(state) {
    state.file.getSliceAsync(state.counter, function (result) {
        if (result.status == Office.AsyncResultStatus.Succeeded) {
            updateStatus("Sending piece " + (state.counter + 1) + " of " + state.sliceCount);
            sendSlice(result.value, state);
        } else {
            updateStatus(result.status);
        }
    });
}
```

The `Slice.data` property returns the raw data of the file as a byte array. If the data is in text format (that is, XML or plain text), the slice contains the raw text. If you pass in **Office.FileType.Compressed** for the _fileType_ parameter of `Document.getFileAsync`, the slice contains the binary data of the file as a byte array. In the case of a PowerPoint or Word file, the slices contain byte arrays.

You must implement your own function (or use an available library) to convert byte array data to a Base64-encoded string. For information about Base64 encoding with JavaScript, see [Base64 encoding and decoding](https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding).

Once you've converted the data to Base64, you can then transmit it to a web server in several ways, including as the body of an HTTP POST request.

Add the following code to send a slice to a web service.

> [!NOTE]
> This code sends a PowerPoint or Word file to the web server in multiple slices. The web server or service must append each individual slice into a single file, and then save it as a .pptx or .docx file before you can perform any manipulations on it.

```js
function sendSlice(slice, state) {
    var data = slice.data;

    // If the slice contains data, create an HTTP request.
    if (data) {

        // Encode the slice data, a byte array, as a Base64 string.
        // NOTE: The implementation of myEncodeBase64(input) function isn't
        // included with this example. For information about Base64 encoding with
        // JavaScript, see https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding.
        var fileData = myEncodeBase64(data);

        // Create a new HTTP request. You need to send the request
        // to a webpage that can receive a post.
        var request = new XMLHttpRequest();

        // Create a handler function to update the status
        // when the request has been sent.
        request.onreadystatechange = function () {
            if (request.readyState == 4) {

                updateStatus("Sent " + slice.size + " bytes.");
                state.counter++;

                if (state.counter < state.sliceCount) {
                    getSlice(state);
                } else {
                    closeFile(state);
                }
            }
        }

        request.open("POST", "[Your receiving page or service]");
        request.setRequestHeader("Slice-Number", slice.index);

        // Send the file as the body of an HTTP POST
        // request to the web server.
        request.send(fileData);
    }
}
```

As the name implies, the `File.closeAsync` method closes the connection to the document and frees up resources. Although the Office Add-ins sandbox garbage collects out-of-scope references to files, it's still a best practice to explicitly close files once your code is done with them. The `closeAsync` method has a single parameter, _callback_, that specifies the function to call on the completion of the call.

```js
function closeFile(state) {
    // Close the file when you're done with it.
    state.file.closeAsync(function (result) {

        // If the result returns as a success, the
        // file has been successfully closed.
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            updateStatus("File closed.");
        } else {
            updateStatus("File couldn't be closed.");
        }
    });
}
```

The final JavaScript file could look like the following:

```js
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// The initialize or onReady function is required for all add-ins.
Office.initialize = function (reason) {

    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {

        // Run sendFile when Submit is clicked.
        $('#submit').on("click", function () {
            sendFile();
        });

        // Update status.
        updateStatus("Ready to send file.");
    });
}

// Create a function for writing to the status div.
function updateStatus(message) {
    var statusInfo = $('#status');
    statusInfo[0].innerHTML += message + "<br/>";
}

// Get all of the content from a PowerPoint or Word document in 100-KB chunks of text.
function sendFile() {
    Office.context.document.getFileAsync("compressed",
        { sliceSize: 100000 },
        function (result) {

            if (result.status === Office.AsyncResultStatus.Succeeded) {

                // Get the File object from the result.
                var myFile = result.value;
                var state = {
                    file: myFile,
                    counter: 0,
                    sliceCount: myFile.sliceCount
                };

                updateStatus("Getting file of " + myFile.size + " bytes");
                getSlice(state);
            } else {
                updateStatus(result.status);
            }
        });
}

// Get a slice from the file and then call sendSlice.
function getSlice(state) {
    state.file.getSliceAsync(state.counter, function (result) {
        if (result.status == Office.AsyncResultStatus.Succeeded) {
            updateStatus("Sending piece " + (state.counter + 1) + " of " + state.sliceCount);
            sendSlice(result.value, state);
        } else {
            updateStatus(result.status);
        }
    });
}

function sendSlice(slice, state) {
    var data = slice.data;

    // If the slice contains data, create an HTTP request.
    if (data) {

        // Encode the slice data, a byte array, as a Base64 string.
        // NOTE: The implementation of myEncodeBase64(input) function isn't
        // included with this example. For information about Base64 encoding with
        // JavaScript, see https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding.
        var fileData = myEncodeBase64(data);

        // Create a new HTTP request. You need to send the request
        // to a webpage that can receive a post.
        var request = new XMLHttpRequest();

        // Create a handler function to update the status
        // when the request has been sent.
        request.onreadystatechange = function () {
            if (request.readyState == 4) {

                updateStatus("Sent " + slice.size + " bytes.");
                state.counter++;

                if (state.counter < state.sliceCount) {
                    getSlice(state);
                } else {
                    closeFile(state);
                }
            }
        }

        request.open("POST", "[Your receiving page or service]");
        request.setRequestHeader("Slice-Number", slice.index);

        // Send the file as the body of an HTTP POST
        // request to the web server.
        request.send(fileData);
    }
}

function closeFile(state) {
    // Close the file when you're done with it.
    state.file.closeAsync(function (result) {

        // If the result returns as a success, the
        // file has been successfully closed.
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            updateStatus("File closed.");
        } else {
            updateStatus("File couldn't be closed.");
        }
    });
}
```
