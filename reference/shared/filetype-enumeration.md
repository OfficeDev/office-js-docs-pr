
# FileType enumeration
Specifies the format in which to return the document.

|||
|:-----|:-----|
|**Hosts:**|PowerPoint, Word|
|**Last changed in**|1.1|

```js
Office.FileType
```


## Members


**Values**


|**Enumeration**|**Value**|**Description**|
|:-----|:-----|:-----|
|Office.FileType.Compressed|"compressed"|Returns the entire document (.pptx or .docx) in Office Open XML (OOXML) format as a byte array.|
|Office.FileType.Pdf|"pdf"|Returns the entire document in PDF format as a byte array.|
|Office.FileType.Text|"text"|Returns only the text of the document as a  **string**. (Word only)|

## Support details


A capital Y in the following matrix indicates that this enumeration is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this enumeration.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for PowerPoint and Word in Office for iPad.|
|1.1|Added support for saving as PDF.|
|1.0|Introduced|
