---
title: Custom contextual tabs in Office Add-ins
description: 'Learn how to add custom contextual tabs to your Office Add-in.'
ms.date: 11/16/2020
localization_priority: Normal
---

# Custom contextual tabs in Office Add-ins (preview)

A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when an object in the Office document, such as an image or a table has focus; for example, the **Table Design** tab that appears on the Excel ribbon when a table is selected. You can include custom contextual tabs in your Office add-in and specify when they are visible or hidden.

> [!NOTE]
> This article assumes that you are familiar with the following documentation. Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.
>
> - [Basic concepts for Add-in Commands](add-in-commands.md)

> [!IMPORTANT]
> Custom contextual tabs are in preview. Please experiment with them in a development or testing environment but don't add them to a production add-in.
>
> Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:
>
>* Excel on Windows (Microsoft 365 only, not perpetual license): Version ???? (Build ?????.?????) Your Microsoft 365 subscription may need to be on the [Current Channel (Preview)](https://insider.office.com/en-us/join/windows) formerly called "Monthly Channel (Targeted)" or "Insider Slow".

> [!NOTE]
> Custom contextual tabs work only on platforms that support the following requirement sets. For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).
>
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## Behavior of custom contextual tabs



## Major steps for including a contextual tab in an add-in

The following are the major steps for including a custom contextual tab in an add-in:

1. Configure the add-in to use a shared runtime.
1. Define the groups and controls that appear on the tab.
1. Define control strings, such as button names and tooltips in the add-in's manifest.
1. Register the contextual tab with Office.
1. Specify the circumstances when the tab will be visible.

## Configure the add-in to use a shared runtime

Adding custom contextual tabs requires your add-in to use the shared runtime. For more information, [Configure an add-in to use a shared runtime](../excel/configure-your-add-in-to-use-a-shared-runtime.md).

## Define the groups and controls that appear on the tab

Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are mostly defined at runtime with a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object. (Some strings in a contextual tab, such as the titles of buttons on the tab, are defined in the [Resources](../reference/manifest/resources.md) section of the manifest.

> [NOTE!]
> The structure of the RibbonUpdaterData object's properties and subproperties (and the property names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.

We'll construct an example of a **RibbonUpdaterData** object step-by-step. (The complete example is below.)

1. Begin by creating an object with two array properties named `actions` and `tabs`. The `actions` array is a specification of all of the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs.

    ```javascript
    const ribbonUpdater = {
      actions: [

      ],
      tabs: [

      ]
    ```

1. This simple example of a contextual tab will have only a single button and, thus, only a single action. Add the following as the only member of the `actions` array. About this code, note:

    - The `id` and `type` properties are mandatory.
    - The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".
    - The `functionName` property is only used when the value of `type` is "ExecuteFunction`. It is the name of a function defined in the FunctionFile. For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).

    ```javascript
    {
      id: "executeWriteData",
      type: "ExecuteFunction",
      functionName: "writeData"
    }
   ```

1. Add the following as the only member of the `tabs` array. About this code, note:

    - The `id` property is required. Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.
    - The `label` property is required, but the value is *not* the label that the tab will have. Instead, it is a resid (resource ID) that is defined in the [Resources](../reference/manifest/resources.md) section of the manifest. In a later step, you add a resource with this ID to the manifest and assign it a user-friendly string to serve as the label of the contextual tab, such as "Data".
    - The `visible` property is optional and defaults to `false` when not present. You typically want it to be `false` when you are using a **RibbonUpdaterData** object to define a contextual tab when the add-in starts up. You typically set it to true when you are using a **RibbonUpdaterData** object to make the tab visible in response to some event, such as the user selecting an entity of some type in the document.
    - The `groups` property is required. It defines the groups of controls that will appear on the tab. It must have at least one member.

    ```javascript
    {
      id: "CtxTab1",
      label: "CtxTab1_label",
      visible: true,
      groups: [

      ]
    }
    ```

1. In the simple ongoing example, the contextual tab has only a single group. Add the following as the only member of the `groups` array. About this code, note:

    - The `id` property is required. Use a brief, descriptive ID that is unique among all groups in the tab.
    - The `label` property is required, but the value is *not* the label that the group will have. Instead, it is a resid (resource ID) that is defined in the [Resources](../reference/manifest/resources.md) section of the manifest. In a later step, you add a resource with this ID to the manifest and assign it a user-friendly string to serve as the label of the group, such as "Insertion".
    - The `icon` property is required. Its value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.
    - The `controls` property is required. Its value is an array of objects that specify the buttons and other controls in the group.

    ```javascript
    {
        id: "CustomGroup111",
        label: "Group11Title",
        icon: [

        ],
        controls: [

        ]
    }
    ```

1. Every group must have an icon of at least three sizes, 16x16 px, 32x32 px, and 80x80 px. Office decides which icon to use based on the size of the ribbon and Office application window. Add the following objects to the icon array. (If the window and ribbon sizes are large enough for at least one of the controls on the group to appear, then no group icon at all appears. For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this code, note:

    - The `size` property is required. The unit of measure is pixels.
    - The `sourceLocation` property is required. It specifies the full URL to the icon.

    > [IMPORTANT!]
    > Just as you typically must change the URLs in a the add-in's manifest when you move from developing to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JavaScript.

    ```javascript
    {
        size: 16,
        sourceLocation: "https://cdn.contoso.com/addins/datainsertion/Images/Group16x16.png"
    },
    {
        size: 32,
        sourceLocation: "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
    },
    {
        size: 80,
        sourceLocation: "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
    }
    ```

1. In the simple ongoing example, the group has only a single button. Add the following object as the only member of the `controls` array. About this code, note:

    - All the properties are required except `enabled` and `tooltip`.
    - `type` specifies the type of control. The values can be "Button", "Menu", or "MobileButton".
    - `id` can be up to 125 characters. 
    - `actionId` must be the ID of an action defined in the `actions` array.
    - `enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up. The default if not present is `true`.
    - `label` refers to the button label. However, the value is *not* the label that the group will have. Instead, it is a resid (resource ID) that is defined in the [Resources](../reference/manifest/resources.md) section of the manifest. In a later step, you add a resource with this ID to the manifest and assign it a user-friendly string to serve as the label of the button, such as "Write Data".
    - `superTip` represents a rich form of tool tip. Both the `title` and `description` properties of the `superTip` object are resids that are defined in the [Resources](../reference/manifest/resources.md) section of the manifest. In a later step, you add a resource with this ID to the manifest and assign it a user-friendly string.
    - `icon` specifies the icons for the button. The remarks above about the group icon apply here too. 

    ```javascript
    {
        type: "Button",
        id: "CtxBt112",
        actionId: "executeWriteData",
        enabled: false,
        label: "ExeFunc_CtxBt112",
        toolTip: "Btn112ToolTip",
        superTip: {
            title: "Btn112SuperTipTitle",
            description: "Btn112SuperTipDesc"
        },
        icon: [
            {
                size: 16,
                sourceLocation: "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton16x16.png"
            },
            {
                size: 32,
                sourceLocation: "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
            },
            {
                size: 80,
                sourceLocation: "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
            }
        ]
    }
    ```
 
The following is the complete example **RibbonUpdaterData** object:

```javascript
const ribbonUpdater = {
  actions: [
    {
      id: "executeWriteData",
      type: "ExecuteFunction",
      functionName: "writeData"
    }
  ],
  tabs: [
    {
      id: "CtxTab1",
      label: "CtxTab1_label",
      visible: true,
      groups: [
        {
          id: "CustomGroup111",
          label: "Group11Title",
          icon: [
            {
                size: 16,
                sourceLocation: "https://cdn.contoso.com/addins/datainsertion/Images/Group16x16.png"
            },
            {
                size: 32,
                sourceLocation: "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
            },
            {
                size: 80,
                sourceLocation: "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
            }
          ],
          controls: [
            {
                type: "Button",
                id: "CtxBt112",
                actionId: "executeWriteData",
                enabled: false,
                label: "ExeFunc_CtxBt112",
                toolTip: "Btn112ToolTip",
                superTip: {
                    title: "Btn112SuperTipTitle",
                    description: "Btn112SuperTipDesc"
                },
                icon: [
                    {
                        size: 16,
                        sourceLocation: "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton16x16.png"
                    },
                    {
                        size: 32,
                        sourceLocation: "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
                    },
                    {
                        size: 80,
                        sourceLocation: "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
                    }
                ]
            }
          ]
        }
      ]
    }
  ]
}
```



## Register the contextual tab with Office

## Specify the circumstances when the tab will be visible

## Best practice: Test for control status errors

In some circumstances, the ribbon does not repaint after `requestUpdate` is called, so the control's clickable status does not change. For this reason it is a best practice for the add-in to keep track of the status of its controls. The add-in should conform to these rules:

1. Whenever `requestUpdate` is called, the code should record the intended state of the custom buttons and menu items.
2. When a custom control is clicked, the first code in the handler, should check to see if the button should have been clickable. If shouldn't have been, the code should report or log an error and try again to set the buttons to the intended state.

The following example shows a function that disables a button and records the button's status. Note that `chartFormatButtonEnabled` is a global boolean variable that is initialized to the same value as the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest.

```javascript
function disableChartFormat() {
    var button = {id: "ChartFormatButton", enabled: false};
    var parentTab = {id: "CustomChartTab", controls: [button]};
    var ribbonUpdater = {tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);

    chartFormatButtonEnabled = false;
}
```

The following example shows how the button's handler tests for an incorrect state of the button. Note that `reportError` is a function that shows or logs an error.

```javascript
function chartFormatButtonHandler() {
    if (chartFormatButtonEnabled) {

        // Do work here

    } else {
        // Report the error and try again to disable.
        reportError("That action is not possible at this time.");
        disableChartFormat();
    }
}
```

## Error handling

In some scenarios, Office is unable to update the ribbon and will return an error. For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened. Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`. The following is an example of how to handle this error. In this case, the `reportError` method displays the error to the user.

```javascript
function disableChartFormat() {
    try {
        var button = {id: "ChartFormatButton", enabled: false};
        var parentTab = {id: "CustomChartTab", controls: [button]};
        var ribbonUpdater = {tabs: [parentTab]};
        await Office.ribbon.requestUpdate(ribbonUpdater);

        chartFormatButtonEnabled = false;
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, close the Office application, and restart it.");
        }
    }
}
```
