# Action Element
 Specifies the action to perform when the user selects a  [Button](./button-control.md) or [Menu](./menu-control.md) controls.
 
 ## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Yes  | Action type to take|


## Child elements

|  Element |  Description  |
|:-----|:-----|
|  [FunctionName](#functionname) |    Specifies the name of the function to execute. |
|  [SourceLocation](#sourcelocation) |    Specifies the source file location for this action. |
  

## xsi:type
This attribute specifies the kind of action performed when the user selects the button. It can be one of the following
- ExecuteFunction
- ShowTaskpane

## FunctionName
Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](./functionfile.md) element.

```xml
<Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
</Action>
```

## SourceLocation
Required element when  **xsi:type** is "ShowTaskpane". Specifies the source file location for this action. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the [Urls](./resources.md#urls) element in the [Resources](./resources.md) element.

```xml
 <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
```  