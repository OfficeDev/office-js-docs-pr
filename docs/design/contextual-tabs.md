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

1. 

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
              sourceLocation: "https://officedev.github.io/custom-functions/addins/cfsample2/Images/Button16x16.png"
            },
            {
              size: 32,
              sourceLocation: "https://officedev.github.io/custom-functions/addins/cfsample2/Images/Button32x32.png"
            },
            {
              size: 80,
              sourceLocation: "https://officedev.github.io/custom-functions/addins/cfsample2/Images/Button80x80.png"
            }
          ],
          controls: [
            {
              type: "Button",
              id: "CtxBt112",
              enabled: false,
              icon: [
                {
                  size: 16,
                  sourceLocation: "https://officedev.github.io/custom-functions/addins/cfsample2/Images/Button16x16.png"
                },
                {
                  size: 32,
                  sourceLocation: "https://officedev.github.io/custom-functions/addins/cfsample2/Images/Button32x32.png"
                },
                {
                  size: 80,
                  sourceLocation: "https://officedev.github.io/custom-functions/addins/cfsample2/Images/Button80x80.png"
                }
              ],
              labe: "ExeFunc_CtxBt112",
              toolTip: "Btn112ToolTip",
              superTip: {
                title: "Btn112SuperTipTitle",
                description: "Btn112SuperTipDesc"
              },
              actionId: "executeWriteData"
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
