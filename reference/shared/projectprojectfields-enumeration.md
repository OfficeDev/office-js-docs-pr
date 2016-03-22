
# ProjectProjectFields enumeration
Specifies the project fields that are available as a parameter for the  **[getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)** method.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Added in**|1.0|

```
ProjectProjectFields={
    CurrencyDigits: 0, 
    CurrencySymbol: 1, 
    CurrencySymbolPosition: 2, 
    GUID: 3, 
    Finish: 4, 
    Start: 5, 
    ReadOnly: 7, 
    VERSION: 8, 
    WorkUnits: 9, 
    ProjectServerUrl: 10, 
    WSSUrl: 11, 
    WSSList: 12
}
```


## Members


****


|**Member**|**Description**|
|:-----|:-----|
|**CurrencyDigits**|The number of digits after the decimal for the currency.|
|**CurrencySymbol**|The currency symbol.|
|**CurrencySymbolPosition**|The placement of the currency symbol: Not specified = -1; Before the value with no space ($0) = 0; After the value with no space (0$) = 1; Before the value with a space ($ 0) = 2; After the value with a space (0 $) = 3.|
|**GUID**|The GUID of the project.|
|**Finish**|The project finish date.|
|**Start**|The project start date.|
|**ReadOnly**|Specifies whether the project is read-only.|
|**VERSION**|The project version.|
|**WorkUnits**|The work units of the project, such as days or hours.|
|**ProjectServerUrl**|The Project Web App URL, for projects that are stored in Project Server.|
|**WSSUrl**|The SharePoint URL, for projects that are synchronized with a SharePoint list.|
|**WSSList**|The name of the SharePoint list, for projects that are synchronized with a tasks list.|

## Remarks

A  **ProjectProjectFields** constant can be used as a parameter of the **[getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)** method.


## Support details


A capital Y in the following matrix indicates that this enumeration is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this enumeration.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.0|Introduced|

## See also



#### Other resources


[getProjectFieldAsync method](../../reference/shared/projectdocument.getprojectfieldasync.md)
