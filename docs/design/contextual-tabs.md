---
title: Create custom contextual tabs in Office Add-ins
description: Learn how to add custom contextual tabs to your Office Add-in.
ms.date: 06/30/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Create custom contextual tabs in Office Add-ins

A contextual tab is a hidden tab control in the Office ribbon that's displayed in the tab row when a specified event occurs in the Office document. For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected. You include custom contextual tabs in your Office Add-in and specify when they're visible or hidden, by creating event handlers that change the visibility. (However, custom contextual tabs don't respond to focus changes.)

> [!NOTE]
> This article assumes that you're familiar with [Basic concepts for add-in commands](add-in-commands.md). Please review it if you haven't worked with add-in commands (custom menu items and ribbon buttons) recently.

## Prerequisites

Custom contextual tabs are currently only supported on **Excel** and only on the following platforms and builds.

- Excel on the web
- Excel on Windows: Version 2102 (Build 13801.20294) and later.
- Excel on Mac: Version 16.53 (21080600) and later.

Additionally, custom contextual tabs only work on platforms that support the following requirement sets. For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).

- [RibbonApi 1.2](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets)
- [SharedRuntime 1.1](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)

> [!TIP]
> Use the runtime checks in your code to test whether the user's host and platform combination supports these requirement sets as described in [Check for API availability at runtime](../develop/specify-api-requirements-runtime.md). (The technique of specifying the requirement sets in the manifest, which is also described in that article, doesn't currently work for RibbonApi 1.2.) Alternatively, you can [implement an alternate UI experience when custom contextual tabs aren't supported](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-arent-supported).

## Behavior of custom contextual tabs

The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs. The following are the basic principles for the placement custom contextual tabs.

- When a custom contextual tab is visible, it appears on the right end of the ribbon.
- If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.
- If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in. (The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.
- If more than one add-in has a contextual tab that's visible in a specific context, then they appear in the order in which the add-ins were launched.
- Custom *contextual* tabs, unlike custom core tabs, aren't added permanently to the Office application's ribbon. They're present only in Office documents on which your add-in is running.

## Major steps for including a contextual tab in an add-in

The following are the major steps for including a custom contextual tab in an add-in.

1. [Configure the add-in to use a shared runtime](#configure-the-add-in-to-use-a-shared-runtime).
1. [Specify the icons for your contextual tab](#specify-the-icons-for-your-contextual-tab).
1. [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab).
1. [Register the contextual tab with Office](#register-the-contextual-tab-with-office-with-requestcreatecontrols).
1. [Specify the circumstances when the tab will be visible](#specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate).

## Configure the add-in to use a shared runtime

Adding custom contextual tabs requires your add-in to use the [shared runtime](../testing/runtimes.md#shared-runtime). For more information, see [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## Specify the icons for your contextual tab

Before you can customize your contextual tab, you must first specify any icons that will appear on it in your add-in's manifest. Each icon must have at least three sizes: 16x16 px, 32x32 px, and 80x80 px. Select the tab for the type of manifest your add-in uses.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

In the [`"extensions.ribbons.tabs.groups.icons"`](/microsoft-365/extensibility/schema/extension-ribbons-custom-tab-groups-item#icons) array, specify the icons for the group of contextual tab controls that will be displayed on the host's ribbon. For icons that will be used by the tab's buttons and menus, specify these in the `"icons"` property of the [`"extensions.ribbons.tabs.groups.controls"`](/microsoft-365/extensibility/schema/extension-ribbons-custom-tab-groups-item#controls) object.

Because the contextual tab will only be shown when a certain event occurs, you must also set the [`"extensions.ribbons.tabs.groups.controls.overriddenByRibbonApi"`](/microsoft-365/extensibility/schema/extension-common-custom-group-controls-item#overriddenbyribbonapi) property to `true`.

The following is an example.

```json
"ribbons": [
    {
        ...
        "tabs": [
            "groups": [
                {
                    "id": "ContosoGroup",
                    ...
                    "icons": [
                        {
                            "size": 16,
                            "url": "https://cdn.contoso.com/addins/datainsertion/Images/Group16x16.png"
                        },
                        {
                            "size": 32,
                            "url": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
                        },
                        {
                            "size": 80,
                            "url": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
                        }
                    ],
                    "controls": [
                        {
                            "id": "WriteDataButton",
                            ...
                            "icons": [
                                {
                                    "size": 16,
                                    "url": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton16x16.png"
                                },
                                {
                                    "size": 32,
                                    "url": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton16x16.png"
                                },
                                {
                                    "size": 80,
                                    "url": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton16x16.png"
                                }
                            ],
                            ...
                            "overriddenByRibbonApi": true
                        },
                        ...
                    ]
                }
            ]
        ]
    }
],
```

# [Add-in only manifest](#tab/xmlmanifest)

Use the [Image](/javascript/api/manifest/image) element in the [Resources](/javascript/api/manifest/resources) section of your add-in's manifest to specify the icons for your contextual tab.

The following is an example.

```xml
<Resources>
    <bt:Images>
        <bt:Image id="contextual-tab-icon-16" DefaultValue="https://cdn.contoso.com/addins/datainsertion/Images/Group16x16.png"/>
        <bt:Image id="contextual-tab-icon-32" DefaultValue="https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"/>
        <bt:Image id="contextual-tab-icon-80" DefaultValue="https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"/>
        <bt:Image id="contextual-button-icon-16" DefaultValue="https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton16x16.png"/>
        <bt:Image id="contextual-button-icon-32" DefaultValue="https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"/>
        <bt:Image id="contextual-button-icon-80" DefaultValue="https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"/>
    </bt:Images>
    ...
</Resources>
```

---

> [!IMPORTANT]
> When you move your add-in from development to staging or production, remember to update the URLs in your manifest as needed (such as changing the domain from `localhost` to `contoso.com`).

## Define the groups and controls that appear on the tab

Unlike custom core tabs, which are defined in the manifest, custom contextual tabs are defined at runtime with a JSON blob. Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)) method. Custom contextual tabs are only present in documents on which your add-in is currently running. This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened. Also, the `requestCreateControls` method may be run only once in a session of your add-in. If it's called again, an error is thrown.

We'll construct an example of a contextual tabs JSON blob step-by-step. The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json). If you're working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON. For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).

1. Begin by creating a JSON string with two array properties named `actions` and `tabs`. The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs.

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
    - The value of `type` can be either `"ExecuteFunction"` or `"ShowTaskpane"`.
    - The `functionName` property is only used when the value of `type` is `ExecuteFunction`. It's the name of a function defined in the FunctionFile. For more information about the FunctionFile, see [Basic concepts for add-in commands](add-in-commands.md).
    - In a later step, you'll map this action to a button on the contextual tab.

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
    ```

1. Add the following as the only member of the `tabs` array. About this markup, note:

    - The `id` property is required. Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.
    - The `label` property is required. It's a user-friendly string to serve as the label of the contextual tab.
    - The `groups` property is required. It defines the groups of controls that will appear on the tab. It must have at least one member.

    > [!NOTE]
    > The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up. Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present. In a later section, we show how to set the property to `true` in response to an event.

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. In the simple ongoing example, the contextual tab has only a single group. Add the following as the only member of the `groups` array. About this markup, note:

    - All the properties are required.
    - The `id` property must be unique among all the groups in the manifest. Use a brief, descriptive ID, of up to 125 characters.
    - The `label` is a user-friendly string to serve as the label of the group.
    - The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.
    - The `controls` property's value is an array of objects that specify the buttons and menus in the group. There must be at least one.

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

1. Every group must have an icon of at least three sizes: 16x16 px, 32x32 px, and 80x80 px. Optionally, you can also have icons of sizes 20x20 px, 24x24 px, 40x40 px, 48x48 px, and 64x64 px. Office decides which icon to use based on the size of the ribbon and Office application window. Add the following objects to the icon array. (If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears. For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:

    - Both the properties are required.
    - The `size` property unit of measure is pixels. Icons are always square, so the number is both the height and the width.
    - The `sourceLocation` property specifies the full URL to the icon. Its value must match the URL specified in the `<Image>` element of the `<Resources>` section of your manifest (see [Specify the icons for your contextual tab](#specify-the-icons-for-your-contextual-tab)).

    > [!IMPORTANT]
    > Just as you typically must change the URLs in the add-in's manifest when you move from development to production, you must also change the URLs in your contextual tabs JSON.

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

1. In our simple ongoing example, the group has only a single button. Add the following object as the only member of the `controls` array. About this markup, note:

    - All the properties, except `enabled`, are required.
    - `type` specifies the type of control. The values can be `"Button"`, `"Menu"`, or `"MobileButton"`.
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

The following is the complete example of the JSON blob.

```json
`{
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
      "label": "Contoso Data",
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
}`
```

## Register the contextual tab with Office with requestCreateControls

The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)) method. This is typically done in either the function that's assigned to `Office.initialize` or with the `Office.onReady` function. For more about these functions and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md). You can, however, call the method anytime after initialization.

> [!IMPORTANT]
> The `requestCreateControls` method may be called only once in a given session of an add-in. An error is thrown if it's called again.

The following is an example. Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## Specify the contexts when the tab will be visible with requestUpdate

Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context. Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.

Begin by assigning handlers. This is commonly done in the `Office.onReady` function as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);

    await Excel.run(context => {
        const charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(showDataTab);
        charts.onDeactivated.add(hideDataTab);
        return context.sync();
    });
});
```

Next, define the handlers. The following is a simple example of a `showDataTab`, but see [Handling the HostRestartNeeded error](#handle-the-hostrestartneeded-error) later in this article for a more robust version of the function. About this code, note:

- Office controls when it updates the state of the ribbon. The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestupdate-member(1)) method queues a request to update. The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.
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
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### Toggle tab visibility and the enabled status of a button at the same time

The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Change the availability of add-in commands](disable-add-in-commands.md). There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time. You do this with a single call of `requestUpdate`. The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.

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
                groups: [
                    {
                        id: "CustomGroup111",
                        controls: [
                            {
                                id: "MyButton",
                                enabled: true
                            }
                        ]
                    }
                ]
            ]}
        ]
    });
}
```

In the following example, the button that's enabled is on the very same contextual tab that is being made visible.

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
                groups: [
                    {
                        id: "CustomGroup111",
                        controls: [
                            {
                                id: "MyButton",
                                enabled: true
                           }
                       ]
                   }
               ]
            }
        ]
    });
}
```

## Open a task pane from contextual tabs

To open your task pane from a button on a custom contextual tab, create an action in the JSON with a `type` of `ShowTaskpane`. Then define a button with the `actionId` property set to the `id` of the action. This opens the default task pane specified in your manifest.

```json
`{
  "actions": [
    {
      "id": "openChartsTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Charts",
      "supportPinning": false
    }
  ],
  "tabs": [
    {
      // some tab properties omitted
      "groups": [
        {
          // some group properties omitted
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "openChartsTaskpane",
                "enabled": false,
                "label": "Open Charts Taskpane",
                // some control properties omitted
            }
          ]
        }
      ]
    }
  ]
}`
```

To open any task pane that's not the default task pane, specify a `sourceLocation` property in the definition of the action. In the following example, a second task pane is opened from a different button.

> [!IMPORTANT]
>
> - When a `sourceLocation` is specified for the action, then the task pane does *not* use the shared runtime. It runs in a new separate runtime.
> - No more than one task pane can use the shared runtime, so no more than one action of type `ShowTaskpane` can omit the `sourceLocation` property.

```json
`{
  "actions": [
    {
      "id": "openChartsTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Charts",
      "supportPinning": false
    },
    {
      "id": "openTablesTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Tables",
      "supportPinning": false
      "sourceLocation": "https://MyDomain.com/myPage.html"
    }
  ],
  "tabs": [
    {
      // some tab properties omitted
      "groups": [
        {
          // some group properties omitted
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "openChartsTaskpane",
                "enabled": false,
                "label": "Open Charts Taskpane",
                // some control properties omitted
            },
            {
                "type": "Button",
                "id": "CtxBt113",
                "actionId": "openTablesTaskpane",
                "enabled": false,
                "label": "Open Tables Taskpane",
                // some control properties omitted
            }
          ]
        }
      ]
    }
  ]
}`
```

## Localize the JSON text

The JSON blob that's passed to `requestCreateControls` isn't localized the same way that the manifest markup for custom core tabs is localized (which is described at [Control localization from the manifest](../develop/localization.md#control-localization-from-the-manifest)). Instead, the localization must occur at runtime using distinct JSON blobs for each locale. We suggest that you use a `switch` statement that tests the [Office.context.displayLanguage](/javascript/api/office/office.context#office-office-context-displaylanguage-member) property. The following is an example.

```javascript
function GetContextualTabsJsonSupportedLocale () {
    const displayLanguage = Office.context.displayLanguage;

        switch (displayLanguage) {
            case 'en-US':
                return `{
                    "actions": [
                        // actions omitted
                     ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Data",
                          "groups": [
                              // groups omitted
                          ]
                        }
                    ]
                }`;

            case 'fr-FR':
                return `{
                    "actions": [
                        // actions omitted 
                    ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Données",
                          "groups": [
                              // groups omitted
                          ]
                       }
                    ]
               }`;

            // Other cases omitted
       }
}
```

Then your code calls the function to get the localized blob that's passed to `requestCreateControls`, as in the following example.

```javascript
const contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## Best practices for custom contextual tabs

### Implement an alternate UI experience when custom contextual tabs aren't supported

Some combinations of platform, Office application, and Office build don't support `requestCreateControls`. Your add-in should be designed to provide an alternate experience to users who are running the add-in on one of those combinations. The following sections describe two ways of providing a fallback experience.

#### Use noncontextual tabs or controls

The add-in's manifest provides a way to create a fallback experience in an add-in that implements custom contextual tabs when the add-in is running on an application or platform that doesn't support custom contextual tabs. The strategy is to define a custom core tab (that is, *noncontextual* custom tab) in the manifest that duplicates the ribbon customizations of the custom contextual tabs in your add-in. Then you use special manifest markup to enable the custom core tab to be visible all the time on platform and version combinations that don't support custom contextual tabs. The process depends on which type of manifest your add-in uses.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

[!include[Unified manifest host application support note](../includes/unified-manifest-support-note.md)]

Begin by defining a custom core tab (that is, *noncontextual* custom tab) in the manifest that duplicates the ribbon customizations of the custom contextual tabs in your add-in. Then, mark any control groups, or individual controls, or menu items that shouldn't be visible on platforms that support contextual tabs. You mark a group, control, or menu item object by adding an `"overriddenByRibbonApi"` property to it and setting its value to `true`. The effect of doing so is the following:

- If the add-in runs on an application and platform that support custom contextual tabs, then the marked custom groups, controls, and menu items won't appear on the ribbon. Instead, the custom contextual tab will be created when the add-in calls the `requestCreateControls` method.
- If the add-in runs on an application or platform that *doesn't* support `requestCreateControls`, then the groups, controls, and menu items do appear on the custom core tab.

The following is an example. Note that "Contoso.MyButton1" will appear on the custom core tab only when custom contextual tabs aren't supported. However, the parent group (with "ContosoButton2") and the custom core tab will appear regardless of whether custom contextual tabs are supported.

```json
"extensions": [
    ...
    {
        ...
        "ribbons": [
            ...
            {
                ...
                "tabs": [
                    {
                        "id": "MyTab",
                        "groups": [
                            {
                                ...
                                "controls": [
                                    {
                                        "id": "Contoso.MyButton1",
                                        ...
                                        "overriddenByRibbonApi": true
                                    },
                                    {
                                        "id": "Contoso.MyButton2",
                                        ...
                                    }
                                ]
                            }
                        ]
                    }
                ]
            }
        ]
    }
]
```

The following is another example. Note that "MyControlGroup" will appear on the custom core tab only when custom contextual tabs aren't supported. However, the parent custom core tab (with unmarked groups) will appear regardless of whether custom contextual tabs are supported.

```json
"extensions": [
    ...
    {
        ...
        "ribbons": [
            ...
            {
                ...
                "tabs": [
                    {
                        "id": "MyTab",
                        "groups": [
                            {
                                "id": "MyControlGroup",
                                "overriddenByRibbonApi": true
                                ...
                                "controls": [
                                    {
                                        "id": "Contoso.MyButton1",
                                        ...
                                    }
                                ]
                            },
                            ... other groups configured here
                        ]
                    }
                ]
            }
        ]
    }
]
```

When a parent menu control is marked with `"overriddenByRibbonApi": true`, then it isn't visible, and all of its child markup is ignored when custom contextual tabs aren't supported. So, it doesn't matter if any of those child menu items have the `"overriddenByRibbonApi"` property or what its value is. The implication of this is that if a menu item must be visible in all contexts, then not only should it not be marked with `"overriddenByRibbonApi": true`, but *its ancestor menu control must also not be marked this way*. A similar point applies to ribbon controls. If a control must be visible in all contexts, then not only should it not be marked with `"overriddenByRibbonApi": true`, but its parent group must also not be marked this way.

> [!IMPORTANT]
> Don't mark *all* of the child items of a menu with `"overriddenByRibbonApi": true`. This is pointless if the parent element is marked with `"overriddenByRibbonApi": true` for reasons given in the preceding paragraph. Moreover, if you leave out the `"overriddenByRibbonApi"` property on the parent menu control (or set it to `false`), then the parent will appear regardless of whether custom contextual tabs are supported, but it will be empty when they are supported. So, if all the child elements shouldn't appear when custom contextual tabs are supported, mark the *parent* menu control with `"overriddenByRibbonApi": true`.
>
> A parallel point applies to groups and controls, don't mark all of the controls in a group with `"overriddenByRibbonApi": true`. This is pointless if the parent group is marked with `"overriddenByRibbonApi": true`. Moreover, if you leave out the `"overriddenByRibbonApi"` property on the parent group (or set it to `false`), then the group will appear regardless of whether custom contextual tabs are supported, but it will have no controls in it when they are supported. So, if all the controls shouldn't appear when custom contextual tabs are supported, mark the parent group with `"overriddenByRibbonApi": true`.

# [Add-in only manifest](#tab/xmlmanifest)

Begin by defining a custom core tab (that is, *noncontextual* custom tab) in the manifest that duplicates the ribbon customizations of the custom contextual tabs in your add-in. But add an [OverriddenByRibbonApi](/javascript/api/manifest/overriddenbyribbonapi) element as the first child element of the duplicate [Group](/javascript/api/manifest/group), [Control](/javascript/api/manifest/control), and menu `<Item>` elements on the custom core tabs. The effect of doing so is the following:

- If the add-in runs on an application and platform that support custom contextual tabs, then the custom core groups and controls won't appear on the ribbon. Instead, the custom contextual tab will be created when the add-in calls the `requestCreateControls` method.
- If the add-in runs on an application or platform that *doesn't* support `requestCreateControls`, then the elements do appear on the custom core tab.

The following is an example. Note that "MyButton" will appear on the custom core tab only when custom contextual tabs aren't supported. But the parent group and custom core tab will appear regardless of whether custom contextual tabs are supported.

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>              
              ...
              <Group ...>
                ...
                <Control ... id="Contoso.MyButton1">
                  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
                  ...
                  <Action ...>
...
</OfficeApp>
```

For more examples, see [OverriddenByRibbonApi](/javascript/api/manifest/overriddenbyribbonapi).

When a parent group, or menu is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, then it isn't visible, and all of its child markup is ignored when custom contextual tabs aren't supported. So, it doesn't matter if any of those child elements have the `<OverriddenByRibbonApi>` element or what its value is. The implication of this is that if a menu item or control must be visible in all contexts, then not only should it not be marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, but *its ancestor menu and group must also not be marked this way*.

> [!IMPORTANT]
> Don't mark *all* of the child elements of a group or menu with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`. This is pointless if the parent element is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` for reasons given in the preceding paragraph. Moreover, if you leave out the `<OverriddenByRibbonApi>` on the parent (or set it to `false`), then the parent will appear regardless of whether custom contextual tabs are supported, but it will be empty when they are supported. So, if all the child elements shouldn't appear when custom contextual tabs are supported, mark the parent with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.

---

#### Use APIs that show or hide a task pane in specified contexts

As an alternative to using the manifest, your add-in can define a task pane with UI controls that duplicate the functionality of the controls on a custom contextual tab. Then use the [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-showastaskpane-member(1)) and [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-hide-member(1)) methods to show the task pane when the contextual tab would have been shown if it was supported. For details on how to use these methods, see [Show or hide the task pane of your Office Add-in](../develop/show-hide-add-in.md).

### Handle the HostRestartNeeded error

In some scenarios, Office is unable to update the ribbon and will return an error. For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened. Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`. Your code should handle this error. The following is an example of how. In this case, the `reportError` method displays the error to the user.

```javascript
function showDataTab() {
    try {
        Office.ribbon.requestUpdate({
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

## Resources

- [Code sample: Create custom contextual tabs on the ribbon](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-contextual-tabs)
- Community demo of contextual tabs sample

    > [!VIDEO https://www.youtube.com/embed/9tLfm4boQIo]
