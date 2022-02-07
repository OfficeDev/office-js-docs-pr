---
title: Control element of type Menu in the manifest file
description: Defines a menu whose items can execute actions or launch task panes.
ms.date: 02/04/2022
ms.localizationpriority: medium
---


# Control element of type Menu

A menu defines a list of options. Each menu item either executes a function or shows a task pane.

> [!NOTE]
> This article assumes familiarity with the basic [Control reference article](control.md) which contains important information about the element's attributes.

The menu control defines:

- A root-level menu control.
- A list of menu items.

When used with the **PrimaryCommandSurface** [extension point](extensionpoint.md), the root menu item displays as a button on the ribbon. When the button is selected, the menu displays as a dropdown list. Submenus are not supported.

When used with the **ContextMenu** [extension point](extensionpoint.md), a root menu item displays on the context menu. When the root item is selected, the menu items display as a submenu. None of the items can itself be a submenu because only one level of submenus is supported.

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Label](#label)     | Yes |  The text for the menu. |
|  **ToolTip**    |No|The tooltip for the menu. The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.|
|  [Supertip](supertip.md)  | Yes |  The supertip for this menu.    |
|  [Icon](icon.md)      | Yes |  An image for the menu.         |
|  **Items**     | Yes |  A collection of items to display within the menu. Contains the **Item** element for each item. |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | No |  Specifies whether the menu should appear on application and platform combinations that support custom contextual tabs. If used, it must be the *first* child element. |

### Label

Specifies the text for the menu name by means of its only attribute, **resid**, which can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** child of the [Resources](resources.md) element.

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

In the following example, the menu has two items. The first displays a task pane. The second executes a function. The menu has been configured to *not* be visible when the add-in is running on a platform that supports contextual tabs. For more information, please read [Implement an alternate UI experience when custom contextual tabs are not supported](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

```xml
<Control xsi:type="Menu" id="Contoso.TestMenu2">
  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="ShowMainTaskPane">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="GetData">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getData</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

In the following example, the menu's second item is configured to *not* be visible when the add-in is running on a platform that supports contextual tabs. For more information, please read [Implement an alternate UI experience when custom contextual tabs are not supported](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

```xml
<Control xsi:type="Menu" id="Contoso.msgReadMenuButton">
  <Label resid="menuReadButtonLabel" />
  <Supertip>
    <Title resid="menuReadSuperTipTitle" />
    <Description resid="menuReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="red-icon-16" />
    <bt:Image size="32" resid="red-icon-32" />
    <bt:Image size="80" resid="red-icon-80" />
  </Icon>
  <Items>
    <Item id="ShowMainTaskPane">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="msgReadMenuItem1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <Label resid="menuItem1ReadLabel" />
      <Supertip>
        <Title resid="menuItem1ReadLabel" />
        <Description resid="menuItem1ReadTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getItemClass</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```
