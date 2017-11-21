
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

## Members


**Objects**


|**Name**|**Description**|
|:-----|:-----|
|[MatrixBinding](https://dev.office.com/reference/add-ins/shared/binding.matrixbinding)|Represents a binding in two dimensions of rows and columns.|
|[TableBinding](https://dev.office.com/reference/add-ins/shared/binding.tablebinding)|Represents a binding in two dimensions of rows and columns, optionally with headers.|
|[TextBinding](https://dev.office.com/reference/add-ins/shared/binding.textbinding)|Represents a bound text selection in the document.|

**Properties**


|**Name**|**Description**|
|:-----|:-----|
|[document](https://dev.office.com/reference/add-ins/shared/binding.document)|Get the  **Document** object associated with the binding.|
|[id](https://dev.office.com/reference/add-ins/shared/binding.id)|Gets the identifier of the object.|
|[type](https://dev.office.com/reference/add-ins/shared/binding.type)|Gets the type of the binding.|

**Methods**


|**Name**|**Description**|
|:-----|:-----|
|[addHandlerAsync](https://dev.office.com/reference/add-ins/shared/binding.addhandlerasync)|Adds a handler to the binding for the specified event type.|
|[getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync)|Returns the data contained within the binding.|
|[removeHandlerAsync](https://dev.office.com/reference/add-ins/shared/binding.removehandlerasync)|Removes the specified handler from the binding for the specified event type.|
|[setDataAsync](https://dev.office.com/reference/add-ins/shared/binding.setdataasync)|Writes data to the bound section of the document represented by the specified binding object.|
|[TableBinding.setFormatsAsync](https://dev.office.com/reference/add-ins/shared/binding.tablebinding.setformatsasync)|Sets or updates formatting on specified items and data in the bound table.|

**Events**


|**Name**|**Description**|
|:-----|:-----|
|[bindingDataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent)|Occurs when data within the binding is changed.|
|[bindingSelectionChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingselectionchangedevent)|Occurs when the selection is changed within the binding.|

## Remarks

The  **Binding** object exposes the functionality possessed by all bindings regardless of type.

The  **Binding** object is never called directly. It is the abstract parent class of the objects that represent each type of binding: [MatrixBinding](https://dev.office.com/reference/add-ins/shared/binding.matrixbinding), [TableBinding](https://dev.office.com/reference/add-ins/shared/binding.tablebinding), or [TextBinding](https://dev.office.com/reference/add-ins/shared/binding.textbinding). All three of these objects inherit the  **getDataAsync** and **setDataAsync** methods from the **Binding** object that enable to you interact with the data in the binding. They also inherit the **id** and **type** properties for querying those property values. Additionally, the **MatrixBinding** and **TableBinding** objects expose additional methods for matrix- and table-specific features, such as counting the number of rows and columns.


## Support details


Support for each API member of the  **Binding** object differs across Office host applications. See the "Support details" section of each member's topic for host support information.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


|||
|:-----|:-----|
|**Available in requirement sets**|MatrixBinding, TableBinding, TextBinding|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|
