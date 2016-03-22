
# Binding object
An abstract class that represents a binding to a section of the document.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**Available in [Requirement sets](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBinding, TableBinding, TextBinding|
|**Last changed in TableBinding**|1.1|

```js
Office.context.document.bindings.getByIdAsync(id);
```

[![Try out this call in the interactive API Tutorial for Excel](../../images/819b84bf-151c-4a12-80c3-d6f8d7c03251.png)](http://officeapitutorial.azurewebsites.net/Redirect.html?scenario=Update+a+Row+in+a+Table)

## Members


**Objects**


|**Name**|**Description**|
|:-----|:-----|
|[MatrixBinding](../../reference/shared/binding.matrixbinding.md)|Represents a binding in two dimensions of rows and columns.|
|[TableBinding](../../reference/shared/binding.tablebinding.md)|Represents a binding in two dimensions of rows and columns, optionally with headers.|
|[TextBinding](../../reference/shared/binding.textbinding.md)|Represents a bound text selection in the document.|

**Properties**


|**Name**|**Description**|
|:-----|:-----|
|[document](../../reference/shared/binding.document.md)|Get the  **Document** object associated with the binding.|
|[id](../../reference/shared/binding.id.md)|Gets the identifier of the object.|
|[type](../../reference/shared/binding.type.md)|Gets the type of the binding.|

**Methods**


|**Name**|**Description**|
|:-----|:-----|
|[addHandlerAsync](../../reference/shared/binding.addhandlerasync.md)|Adds a handler to the binding for the specified event type.|
|[getDataAsync](../../reference/shared/binding.getdataasync.md)|Returns the data contained within the binding.|
|[removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md)|Removes the specified handler from the binding for the specified event type.|
|[setDataAsync](../../reference/shared/binding.setdataasync.md)|Writes data to the bound section of the document represented by the specified binding object.|
|[TableBinding.setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md)|Sets or updates formatting on specified items and data in the bound table.|

**Events**


|**Name**|**Description**|
|:-----|:-----|
|[bindingDataChanged](../../reference/shared/binding.bindingdatachangedevent.md)|Occurs when data within the binding is changed.|
|[bindingSelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md)|Occurs when the selection is changed within the binding.|

## Remarks

The  **Binding** object exposes the functionality possessed by all bindings regardless of type.

The  **Binding** object is never called directly. It is the abstract parent class of the objects that represent each type of binding: [MatrixBinding](../../reference/shared/binding.matrixbinding.md), [TableBinding](../../reference/shared/binding.tablebinding.md), or [TextBinding](../../reference/shared/binding.textbinding.md). All three of these objects inherit the  **getDataAsync** and **setDataAsync** methods from the **Binding** object that enable to you interact with the data in the binding. They also inherit the **id** and **type** properties for querying those property values. Additionally, the **MatrixBinding** and **TableBinding** objects expose additional methods for matrix- and table-specific features, such as counting the number of rows and columns.


## Support details


Support for each API member of the  **Binding** object differs across Office host applications. See the "Support details" section of each member's topic for host support information.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


|||
|:-----|:-----|
|**Available in requirement sets**|MatrixBinding, TableBinding, TextBinding|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|
