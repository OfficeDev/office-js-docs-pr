
# CustomXmlPrefixMappings object
Represents a collection of custom namespace prefix mappings.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Last changed in**|1.1|

```
CustomXmlPrefixMappings
```


## Members


**Methods**


|**Name**|**Description**|
|:-----|:-----|
|[addNamespaceAsync](../../reference/shared/customxmlprefixmappings.addnamespaceasync.md)|Asynchronously adds a prefix to namespace mapping to use when querying an item.|
|[getNamespaceAsync](../../reference/shared/customxmlprefixmappings.getnamespaceasync.md)|Asynchronously gets the namespace mapped to the specified prefix.|
|[getPrefixAsync](../../reference/shared/customxmlprefixmappings.getprefixasync.md)|Asynchronously gets the prefix for the specified namespace.|

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
