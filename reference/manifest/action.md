# Action element
Specifies the action to perform when the user selects a  [Button](./control.md#button-control) or [Menu](./control.md#menu-dropdown-button-controls) controls.
 
## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Yes  | Action type to take|


## Child elements

|  Element |  Description  |
|:-----|:-----|
|  [FunctionName](#functionname) |    Specifies the name of the function to execute. |
|  [SourceLocation](#sourcelocation) |    Specifies the source file location for this action. |
|  [TaskpaneId](#taskpaneid) | Specifies the ID of the task pane container.|
|  [Title](#title) | Specifies the custom title for the task pane.|
|  [SupportsPinning](#supportspinning) | Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.|
  

## xsi:type
This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:

- `ExecuteFunction`
- `ShowTaskpane`

## FunctionName

Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](./functionfile.md) element.

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
.</Action>
```

## SourceLocation
Required element when  **xsi:type** is "ShowTaskpane". Specifies the source file location for this action. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the [Urls](./resources.md#urls) element in the [Resources](./resources.md) element.

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## TaskpaneId
Optional element when  **xsi:type** is "ShowTaskpane". Specifies the ID of the task pane container. When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each. Use the same **TaskpaneId** for  different actions that share the same pane. When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation". 

>**Note:** This element is not supported in Outlook.

The following example shows two Actions that share the same **TaskpaneId**. 


```xml
<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="aTaskPaneUrl" />
</Action>

<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="anotherTaskPaneUrl" />
</Action>
```  

## Title
Optional element when  **xsi:type** is "ShowTaskpane". Specifies the custom title for the task pane for this action. 

The following examples show two different Actions that use the **Title** element. To see the examples in their entirety, see [Script Lab manifest](https://github.com/OfficeDev/script-lab/blob/master/manifests/script-lab-local.xml)

```xml
<Action xsi:type="ShowTaskpane">
<TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
<SourceLocation resid="PG.Code.Url" />
<Title resid="PG.CodeCommand.Title" />
</Action>
``` 

```xml
<Action xsi:type="ShowTaskpane">
<SourceLocation resid="PG.Run.Url" />
<Title resid="PG.RunCommand.Title" />
</Action>
``` 

```xml
<bt:Urls>
<bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
<bt:Url id="PG.Run.Url" DefaultValue="https://localhost:3000/run.html" />
</bt:Urls>
``` 
      
## SupportsPinning

Optional element when **xsi:type** is "ShowTaskpane". The containing [VersionOverrides](./versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`. Include this element with a value of `true` to support taskpane pinning. The user will be able to "pin" the taskpane, causing it to stay open when changing the selection. For more information, see [Implement a pinnable taskpane in Outlook](../../docs/outlook/manifests/pinnable-taskpane).

> **Note**: SupportsPinning currently only supported by Outlook 2016 for Windows (build 7628.1000 or later).

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```


