
# Document.getFileAsync method
Returns the entire document file in slices of up to 4194304 bytes (4MB) or for add-ins for iOS up to 65536 (64KB).

|||
|:-----|:-----|
|**Hosts:**|Excel, PowerPoint, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|File|
|**Last changed in File**|1.1|

```js
Office.context.document.getFileAsync(fileType [, options], callback);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _fileType_|[FileType](../../reference/shared/filetype-enumeration.md)|Specifies the format in which the file will be returned. Required.<br/><table><tr><th>Host</th><th>Supported fileType</th></tr><tr><td>Excel Online</td><td>Office.FileType.Compressed</td></tr><tr><td>PowerPoint on Windows desktop</td><td>Office.FileType.CompressedOffice.FileType.Pdf</td></tr><tr><td>Word on Windows desktop and iPad</td><td>Office.FileType.CompressedOffice.FileType.PdfOffice.FileType.Text</td></tr><tr><td>Word Online</td><td>Office.FileType.CompressedOffice.FileType.PdfOffice.FileType.Text</td></tr><tr><td>PowerPoint Online</td><td>Office.FileType.CompressedOffice.FileType.Pdf</td></tr></table>|**Changed in** 1.1, see [Support history](#bk_history)|
| _options_|**object**|Specifies any of the following [optional parameters](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _sliceSize_|**number**|Specifies the desired slice size (in bytes) up to 4194304 bytes (4MB). If not specified, a default slice size of 4194304 bytes (4MB) will be used. ||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **getFileAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Access the [File](../../reference/shared/file.md) object.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determine the success or failure of the operation.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Access an [Error](../../reference/shared/error.md) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Remarks

For add-ins running in Office host applications other than Office for iOS, the  **getFileAsync** method supports getting files in slices of up to 4194304 bytes (4MB). For add-ins running in Office of iOS apps, the **getFileAsync** method supports getting files in slices of up to 65536 (64KB).

The  _fileType_ parameter can be specified using the following enumerations or text values.


**FileType enumeration**


|**Enumeration**|**Value**|**Description**|
|:-----|:-----|:-----|
|Office.FileType.Compressed|"compressed"|Returns the entire document (.docx, .pptx , or .xslx) in Office Open XML (OOXML) format as a byte array.|
|Office.FileType.Pdf|"pdf"|Returns the entire document in PDF format as a byte array.|
|Office.FileType.Text|"text"|Returns only the text of the document as a  **string**. |
No more than two documents are allowed to be in memory; otherwise the  **getFileAsync** operation will fail. Use the [File.closeAsync](../../reference/shared/file.closeasync.md) method to close the file when you are finished working with it.


## Example - Get a document in Office Open XML ("compressed") format

The following example gets the document in Office Open XML ("compressed") format in 65536 bytes (64KB) slices. Note: The implementation of  `app.showNotification` in this example is from the Visual Studio template for Office Add-ins.


```js
function getDocumentAsCompressed() {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 /*64 KB*/ }, 
        function (result) {
            if (result.status == "succeeded") {
            // If the getFileAsync call succeeded, then
            // result.value will return a valid File Object.
            var myFile = result.value;
            var sliceCount = myFile.sliceCount;
            var slicesReceived = 0, gotAllSlices = true, docdataSlices = [];
            app.showNotification("File size:" + myFile.size + " #Slices: " + sliceCount);

            // Get the file slices.
            getSliceAsync(myFile, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
            }
            else {
            app.showNotification("Error:", result.error.message);
            }
    });
}

function getSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived) {
    file.getSliceAsync(nextSlice, function (sliceResult) {
        if (sliceResult.status == "succeeded") {
            if (!gotAllSlices) { // Failed to get all slices, no need to continue.
                return;
            }

            // Got one slice, store it in a temporary array.
            // (Or you can do something else, such as
            // send it to a third-party server.)
            docdataSlices[sliceResult.value.index] = sliceResult.value.data;
            if (++slicesReceived == sliceCount) {
               // All slices have been received.
               file.closeAsync();
               onGotAllSlices(docdataSlices);
            }
            else {
                getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
            }
        }
            else {
                gotAllSlices = false;
                file.closeAsync();
                app.showNotification("getSliceAsync Error:", sliceResult.error.message);
            }
    });
}

function onGotAllSlices(docdataSlices) {
    var docdata = [];
    for (var i = 0; i < docdataSlices.length; i++) {
        docdata = docdata.concat(docdataSlices[i]);
    }

    var fileContent = new String();
    for (var j = 0; j < docdata.length; j++) {
        fileContent += String.fromCharCode(docdata[j]);
    }

    // Now all the file content is stored in 'fileContent' variable,
    // you can do something with it, such as print, fax...
}

```


## Example - Get a document in PDF format

The following example gets the document in PDF format.


```js
Office.context.document.getFileAsync(Office.FileType.Pdf,
    function(result) {
        if (result.status == "succeeded") {
            var myFile = result.value;
            var sliceCount = myFile.sliceCount;
            app.showNotification("File size:" + myFile.size + " #Slices: " + sliceCount);
            // Now, you can call getSliceAsync to download the files, as described in the previous code segment (compressed format).
            
            myFile.closeAsync();
        }
        else {
            app.showNotification("Error:", result.error.message);
        }
}
);


```


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||Y||
|**PowerPoint**|Y||Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|File|
|**Minimum permission level**|[ReadAllDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history


|**Version**|**Changes**|
|:-----|:-----|
|1.1| In PowerPoint Online, added support for **Office.FileType.Pdf** as the _fileType_ parameter.|
|1.1| In PowerPoint Online, added support for **Office.FileType.Compressed** as the _fileType_ parameter.|
|1.1| In Word Online, added support for **Office.FileType.Text** as the _fileType_ parameter.|
|1.1| In Excel Online, added support for **Office.FileType.Compressed** as the _fileType_ parameter.|
|1.1| In Word Online, added support for **Office.FileType.Compressed** and **Office.FileType.Pdf** as the _fileType_ parameter.|
|1.1|In PowerPointWord on Office for iPad, added support for all  **FileType** values as the _fileType_ parameter.|
|1.1|In Word and PowerPoint on Windows desktop, added support for  **Office.FileType.Pdf** as the _fileType_ parameter..|
|1.0|Introduced|
