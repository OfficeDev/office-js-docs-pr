---
title: Custom contextual tabs in Office Add-ins
description: 'Learn how to add custom contextual tabs to your Office Add-in.'
ms.date: 11/20/2020
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
1. Define the tab and the groups and controls that appear on it.
1. Register the contextual tab with Office.
1. Specify the circumstances when the tab will be visible.

## Configure the add-in to use a shared runtime

Adding custom contextual tabs requires your add-in to use the shared runtime. For more information, [Configure an add-in to use a shared runtime](../excel/configure-your-add-in-to-use-a-shared-runtime.md).

## Define the groups and controls that appear on the tab

Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob. With code that runs when the add-in starts, you parse the blob into a JavaScript object, and then pass the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestcreateconrtols-input-) method. 

> [NOTE!]
> The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.

We'll construct an example of a contextual tabs JSON blob step-by-step. (The full schema for the contextual tab JSON is at: .)

1. Begin by creating a JSON string with two array properties named `actions` and `tabs`. The `actions` array is a specification of all of the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, up to a maximum of 10.

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. This simple example of a contextual tab will have only a single button and, thus, only a single action. Add the following as the only member of the `actions` array. About this markup, note:

    - The `id` and `type` properties are mandatory.
    - The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".
    - The `functionName` property is only used when the value of `type` is "ExecuteFunction`. It is the name of a function defined in the FunctionFile. For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).
    - In a later step, you will map this action to a button on the contextual tab.

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. Add the following as the only member of the `tabs` array. About this markup, note:

    - The `id` property is required. Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.
    - The `label` property is required. It is a user-friendly string to serve as the label of the contextual tab.
    - The `groups` property is required. It defines the groups of controls that will appear on the tab. It must have at least one member.
    
    > [NOTE!]
    > The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up. Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present. In a later section, we show how to set the property to `true` in response to an event.

    ```json
    {
      "id": "CtxTab1",
      "label": "Data",
      "groups": [

      ]
    }
    ```

1. In the simple ongoing example, the contextual tab has only a single group. Add the following as the only member of the `groups` array. About this markup, note:

    - All the properties are required.
    - The `id` property must be unique among all the groups in the tab. Use a brief, descriptive ID.
    - The `label` is a user-friendly string to serve as the label of the group.
    - The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.
    - The `controls` property's value is an array of objects that specify the buttons and other controls in the group.

    ```json
    {
        "id": "CustomGroup111",
        "label": "Insertion",
        "icon": [

        ],
        "controls": [

        ]
    }
    ```

1. Every group must have an icon of at least three sizes, 16x16 px, 32x32 px, and 80x80 px. Office decides which icon to use based on the size of the ribbon and Office application window. Add the following objects to the icon array. (If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears. For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:

    - Both the properties are required.
    = The `size` property unit of measure is pixels. Icons are always square, so the number is both the height and the width.
    - The `sourceLocation` property specifies the full URL to the icon.

    > [IMPORTANT!]
    > Just as you typically must change the URLs in a the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.

    ```json
    {
        "size": 16,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group16x16.png"
    },
    {
        "size": 32,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
    },
    {
        "size": 80,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
    }
    ```

1. In the simple ongoing example, the group has only a single button. Add the following object as the only member of the `controls` array. About this markup, note:

    - All the properties, except `enabled` are required.
    - `type` specifies the type of control. The values can be "Button", "Menu", or "MobileButton".
    - `id` can be up to 125 characters. 
    - `actionId` must be the ID of an action defined in the `actions` array. (See above.)
    - `label` is a user-friendly string to serve as the label of the button.
    - `superTip` represents a rich form of tool tip. Both the `title` and `description` properties are required.
    - `icon` specifies the icons for the button. The remarks above about the group icon apply here too.
    - `enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up. The default if not present is `true`. 


    ```json
    {
        "type": "Button",
        "id": "CtxBt112",
        "actionId": "executeWriteData",
        "enabled": false,
        "label": "Write Data",
        "superTip": {
            "title": "Data Insertion",
            "description": "Use this button to insert data into the document."
        },
        "icon": [
            {
                "size": 16,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton16x16.png"
            },
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
            }
        ]
    }
    ```
 
The following is the complete example of the JSON blob:

```json
'{
  "actions": [
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
  ],
  "tabs": [
    {
      "id": "CtxTab1",
      "label": "Data",
      "groups": [
        {
          "id": "CustomGroup111",
          "label": "Insertion",
          "icon": [
            {
                "size": 16,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group16x16.png"
            },
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
            }
          ],
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "executeWriteData",
                "enabled": false,
                "label": "Write Data",
                "superTip": {
                    "title": "Data Insertion",
                    "description": "Use this button to insert data into the document."
                },
                "icon": [
                    {
                        "size": 16,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton16x16.png"
                    },
                    {
                        "size": 32,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
                    },
                    {
                        "size": 80,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
                    }
                ]
            }
          ]
        }
      ]
    }
  ]
}'
```

## Register the contextual tab with Office with requestCreateControls

The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestcreateconrtols-input-) method in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method. For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md).

The following is an example. Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` methd before it can be passed to a JavaScript function.

```javascript
const contextualTabJSON = ' ... ' // Assign the JSON string such as the one at the end of the preceding section.
const contextualTab = JSON.parse(contextualTabJSON);

Office.onReady(async () => {
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## Specify the circumstances when the tab will be visible with requestUpdate

Typically, a contextual tab should appear when a user-initiated event changes the add-in context. Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated. 

Begin by assigning handlers. This is commonly done in the **Office.onReady** method as in the following example which assigns handlers (created in a later step) to the **onActivated** and **onDeactivated** events of all the charts in the worksheet.

```javascript
Office.onReady(async () => {
    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(showDataTab);
        charts.onDeactivated.add(hideDataTab);
        return context.sync();
    });
});
```

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
