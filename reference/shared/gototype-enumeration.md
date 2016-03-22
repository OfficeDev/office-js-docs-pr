
# GoToType enumeration
Specifies the type of place or object to navigate to.

|||
|:-----|:-----|
|**Hosts:**|Excel, PowerPoint, Word|
|**Added in**|1.1|

```js
Office.GoToType
```


## Members


**Values**


|**Enumeration**|**Value**|**Description**|**Supported clients**|
|:-----|:-----|:-----|:-----|
|Office.GoToType.Binding|"binding"|Goes to a binding object using the specified binding id.|Excel</br>Word|
|Office.GoToType.NamedItem|"namedItem"|Goes to an item using that item's name, such as the name assigned to a table or range.In Excel, you can use any structured reference for a named range or table: "Worksheet2!Table1"|Excel|
|Office.GoToType.Slide|"slide"|Goes to a slide using the specified id.|PowerPoint|
|Office.GoToType.Index|"index"|Goes to the specified index by slide number or enum:</br>**Office.Index.First**</br>**Office.Index.Last**</br>**Office.Index.Next**</br>**Office.Index.Previous**|PowerPoint|

## Support details


A capital Y in the following matrix indicates that this enumeration is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this enumeration.


For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
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
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Introduced|
