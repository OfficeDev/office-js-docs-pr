#Button control

A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest. 

## Child elements
|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Label](#label)     | Yes |  The text for the button.         |
|  [Supertip](./supertip.md)  | Yes |  The supertip for the button.    |
|  [Icon](./icon.md)      | Yes |  An image for the button.         |
|  [Action](./action.md)    | Yes |  Specifies the action to perform  |

## Label
Required. The text for the button. The  **resid** attribute must be set to the value of the **id** attribute 
of a **String** element in the [ShortStrings](./resources.md#shortstrings) element in the [Resources](./resources.md)  element.

## Supertip
Required. See [Supertip](./supertip.md).
 
## Icon
Required. See [Icon](./icon.md).

## Action
Required. See [Action](./action.md).

## ExecuteFunction button example
```xml
<Control xsi:type="Button" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

## ShowTaskpane button example
```xml
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```