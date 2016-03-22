
# CustomXmlParts object
Represents a collection of [CustomXMLPart](../../reference/shared/customxmlpart.customxmlpart.md) objects.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Last changed in**|1.1|

```
Office.context.document.customXmlParts
```


## Members


**Methods**


|**Name**|**Description**|
|:-----|:-----|
|[addAsync](../../reference/shared/customxmlparts.addasync.md)|Asynchronously adds a new custom XML part to a file.|
|[getByIdAsync](../../reference/shared/customxmlparts.getbyidasync.md)|Asynchronously gets a custom XML part by its ID.|
|[getByNamespaceAsync](../../reference/shared/customxmlparts.getbynamespaceasync.md)|Asynchronously gets an array of custom XML parts that match the specified namespace.|

## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Available in requirement sets**|CustomXmlParts|
|**Minimum permission level**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Word in Office for iPad.|
|1.0|Introduced|
