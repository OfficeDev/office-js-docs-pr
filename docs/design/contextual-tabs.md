---
title: Create custom contextual tabs in Office Add-ins
description: 'Learn how to add custom contextual tabs to your Office Add-in.'
ms.date: 11/20/2020
localization_priority: Normal
---

# Create custom contextual tabs in Office Add-ins (preview)

A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when a specified event occurs in the Office document. For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected. You can include custom contextual tabs in your Office add-in and specify when they are visible or hidden, by creating event handlers that change the visibility. (However, custom contextual tabs do not respond to focus changes.)

> [!NOTE]
> This article assumes that you are familiar with the following documentation. Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.
>
> - [Basic concepts for Add-in Commands](add-in-commands.md)

> [!IMPORTANT]
> Custom contextual tabs are in preview. Please experiment with them in a development or testing environment but don't add them to a production add-in.
>
> Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:
>
> - Excel on Windows (Microsoft 365 only, not perpetual license): Version 2011 (Build 13426.20274). Your Microsoft 365 subscription may need to be on the [Current Channel (Preview)](https://insider.office.com/join/windows) formerly called "Monthly Channel (Targeted)" or "Insider Slow".

> [!NOTE]
> Custom contextual tabs work only on platforms that support the following requirement sets. For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).
>
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## Behavior of custom contextual tabs

The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs. The following are the basic principles for the placement custom contextual tabs:

- When a custom contextual tab is visible, it appears on the right end of the ribbon.
- If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.
- If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in. (The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.
- If more than one add-in has a contextual tab that is visible in a specific context, then they appear in the order in which the add-ins were installed.
- Custom *contextual* tabs, unlike custom core tabs, are not added permanently to the Office application's ribbon. They are present only in Office documents on which your add-in has run.

## Major steps for including a contextual tab in an add-in

The following are the major steps for including a custom contextual tab in an add-in:

1. Configure the add-in to use a shared runtime.
1. Define the tab and the groups and controls that appear on it.
1. Register the contextual tab with Office.
1. Specify the circumstances when the tab will be visible.

## Configure the add-in to use a shared runtime

Adding custom contextual tabs requires your add-in to use the shared runtime. For more information, see [Configure an add-in to use a shared runtime](../excel/configure-your-add-in-to-use-a-shared-runtime.md).

## Define the groups and controls that appear on the tab

Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob. Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestcreatecontrols-input-) method. Custom contextual tabs are only present in documents on which your add-in has run. This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened. Also, the `requestCreateControls` method can be run only once on a specific document.

> [!NOTE]
> The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.

We'll construct an example of a contextual tabs JSON blob step-by-step. (The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json). This link may not be working in the early preview period for contextual tabs. If the link is not working, you can find the latest draft of the schema at [draft dynamic-ribbon.schema.json](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json).) If you are working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON. For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).


1. Begin by creating a JSON string with two array properties named `actions` and `tabs`. The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, *up to a maximum of 10*.

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
    - The `functionName` property is only used when the value of `type` is `ExecuteFunction`. It is the name of a function defined in the FunctionFile. For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).
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
    - The `groups` property is required. It defines the groups of controls that will appear on the tab. It must have at least one member *and no more than 20*. (There are also limits on the number of controls that you can have on a custom contextual tab and that will also constrain how many groups that you have. See the next step for more information.)

    > [!NOTE]
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
    - The `controls` property's value is an array of objects that specify the buttons and other controls in the group. There must be at least one and *no more than 6 in a group*.

    > [!IMPORTANT]
    > *The total number of controls on the whole tab can be no more than 20.* For example, you could have 3 groups with 6 controls each, and a fourth group with 2 controls, but you cannot have 4 groups with 6 controls each.  

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

1. Every group must have an icon of at least two sizes, 32x32 px and 80x80 px. Optionally, you can also have icons of sizes 16x16, 20x20, 24x24, 40x40, 48x48 and 64x64. Office decides which icon to use based on the size of the ribbon and Office application window. Add the following objects to the icon array. (If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears. For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:

    - Both the properties are required.
    - The `size` property unit of measure is pixels. Icons are always square, so the number is both the height and the width.
    - The `sourceLocation` property specifies the full URL to the icon.

    > [!IMPORTANT]
    > Just as you typically must change the URLs in the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.

    ```json
    {
        "size": 32,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
    },
    {
        "size": 80,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
    }
    ```

1. In our simple ongoing example, the group has only a single button. Add the following object as the only member of the `controls` array. About this markup, note:

    - All the properties, except `enabled`, are required.
    - `type` specifies the type of control. The values can be "Button", "Menu", or "MobileButton".
    - `id` can be up to 125 characters. 
    - `actionId` must be the ID of an action defined in the `actions` array. (See step 1 of this section.)
    - `label` is a user-friendly string to serve as the label of the button.
    - `superTip` represents a rich form of tool tip. Both the `title` and `description` properties are required.
    - `icon` specifies the icons for the button. The previous remarks about the group icon apply here too.
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

The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestcreatecontrols-input-) method. This is typically done in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method. For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md). You can, however, call the method anytime after initialization.

> [!IMPORTANT]
> The `requestCreateControls` method can be called only once on a given document from a specific add-in. 

The following is an example. Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ' ... '; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## Specify the contexts when the tab will be visible with requestUpdate

Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context. Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.

Begin by assigning handlers. This is commonly done in the `Office.onReady` method as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ' ... '; // Assign the JSON string.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);

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

Next, define the handlers. The following is a simple example of a `showDataTab`, but see [Error Handling](#error-handling) later in this article for a more robust version of the function. About this code, note:

- Office controls when it updates the state of the ribbon. The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update. The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.
- The parameter for the `requestUpdate` method is a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the tab by its ID *exactly as specified in the JSON* and (2) specifies visibility of the tab.
- If you have more than one custom contextual tab that should be visible in the same context, you simply add additional tab objects to the `tabs` array.

```javascript
async function showDataTab() {
    await Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            }
        ]});
}
```

The handler to hide the tab is nearly identical, except that it sets the `visible` property back to `false`.

The Office JavaScript library also provides several interfaces (types) to make it easier to construct the`RibbonUpdateData` object. The following is the `showDataTab` function in TypeScript and it makes use of these types.

```typescript
const showDataTab = async () => {
    const myContextualTab: Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### Toggle tab visibility and the enabled status of a button at the same time

The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Enable and Disable Add-in Commands](disable-add-in-commands.md). There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time. You can do this with a single call of `requestUpdate`. The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            },
            {
                id: "OfficeAppTab1",
                controls: [
                {
                    id: "MyButton",
                    enabled: true
                }
            ]}
        ]});
}
```

In the following example, the button that is enabled is on the very same contextual tab that is being made visible.

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
                controls: [
                    {
                        id: "MyButton",
                        enabled: true
                    }
                ]
            }
        ]});
}
```

## Error handling

In some scenarios, Office is unable to update the ribbon and will return an error. For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened. Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`. The following is an example of how to handle this error. In this case, the `reportError` method displays the error to the user.

```javascript
function showDataTab() {
    try {
        await Office.ribbon.requestUpdate({
            tabs: [
                {
                    id: "CtxTab1",
                    visible: true
                }
            ]});
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, then close and reopen the Office application.");
        }
    }
}
```
