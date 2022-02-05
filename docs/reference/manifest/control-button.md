---
title: Control element of type Button in the manifest file
description: Defines a button that executes an action or launches a task pane.
ms.date: 02/04/2022
ms.localizationpriority: medium
---

# Control element of type Button

Defines a button that executes an action or launches a task pane.

> [!NOTE]
> This article assumes familiarity with the basic [Control reference article](control.md) which contains important information about the element's attributes.

A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` attribute value that is unique among all **Control** elements in the manifest.

> [!IMPORTANT]
> "Button" type controls are ignored on mobile platforms. To support mobile platforms, you must also have a control of type "MobileButton" for every control of type "Button".

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Label](#label)     | Yes |  The text for the button. |
|  **ToolTip**    |No|The tooltip for the button. The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.|
|  [Supertip](supertip.md)  | Yes |  The supertip for the button.    |
|  [Icon](icon.md)      | Yes |  An image for the button.         |
|  [Action](action.md)    | Yes |  Specifies the action to perform. There can be only one **Action** child of a **Control** element. |
|  [Enabled](enabled.md)    | No |  Specifies whether the control is enabled when the add-in launches.  |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | No |  Specifies whether the button should appear on application and platform combinations that support custom contextual tabs. If used, it must be the *first* child element. |

### Label

Specifies the text for the button by means of its only attribute, **resid**, which can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** child of the [Resources](resources.md) element.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) when the parent **VersionOverrides** is type Taskpane 1.0.
- [Mailbox 1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) when the parent **VersionOverrides** is type Mail 1.0.
- [Mailbox 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) when the parent **VersionOverrides** is type Mail 1.1.

## Examples

In the following example, the button executes a function. It's also configured to be disabled when the add-in launches. It can be programmatically enabled. For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).

```xml
<Control xsi:type="Button" id="Contoso.msgReadFunctionButton">
  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
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
  <Enabled>false</Enabled>
</Control>
```

In the following example, the button displays a task pane.

```xml
<Control xsi:type="Button" id="Contoso.msgReadOpenPaneButton">
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
