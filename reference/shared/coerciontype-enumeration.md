
# CoercionType enumeration
Specifies how to coerce data returned or set by the invoked method.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Last changed in Mailbox**|1.1|

```js
Office.CoercionType
```

[![Try out this call in the interactive API Tutorial for Excel](../../images/819b84bf-151c-4a12-80c3-d6f8d7c03251.png)](https://officeapitutorial.azurewebsites.net/TryItOut.html)

## Members


**Values**


|**Enumeration**|**Value**|**Description**|
|:-----|:-----|:-----|
|Office.CoercionType.Html|"html"|Return or set data as HTML.<br/><br/> **Note**  Only applies to data in add-ins for Word and Outlook add-ins for Outlook (compose mode).|
|Office.CoercionType.Matrix|"matrix"|Return or set data as tabular data with no headers. Data is returned or set as an array of arrays containing one-dimensional runs of characters. For example, three rows of  **string** values in two columns would be: ` [["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`.<br/><br/> **Note**  Only applies to data in Excel and Word.|
|Office.CoercionType.Ooxml|"ooxml"|Return or set data as Office Open XML.<br/><br/> **Note**  Only applies to data in Word.|
|Office.CoercionType.SlideRange|"slideRange"|Return a JSON object that contains an array of the ids, titles, and indexes of the selected slides.For example,  `{"slides":[{"id":257,"title":"Slide 2","index":2},{"id":256,"title":"Slide 1","index":1}]}` for a selection of two slides.<br/><br/> **Note**  Only applies to data in PowerPoint when calling the [Document.getSelectedData](../../reference/shared/document.getselecteddataasync.md) method to get the current slide or selected range of slides.|
|Office.CoercionType.Table|"table"|Return or set data as tabular data with optional headers. Data is returned or set as an array of arrays with optional headers.<br/><br/> **Note**  Only applies to data in Access, Excel and Word.|
|Office.CoercionType.Text|"text"|Return or set data as text ( **string**).Data is returned or set as a one-dimensional run of characters.|
|Office.CoercionType.Image|"image"|Data is returned or set as an image stream.<br/><br/> **Note**  Only applies to data in Excel, Word and PowerPoint.|
PowerPoint supports only  **Office.CoercionType.Text**,  **Office.CoercionType.Image**, and  **Office.CoercionType.SlideRange**.

Project supports only  **Office.CoercionType.Text**.


## Support details


A capital Y in the following matrix indicates that this enumeration is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this enumeration.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|**OWA for Devices**|**Office for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**|Y|||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**Add-in types**|Content, Outlook (compose mode), task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Word Online.|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Added support for add-ins for Access.|
|1.1|Added support for [compose mode Outlook add-ins](../../docs/outlook/compose-scenario.md).|
|1.0|Introduced|
