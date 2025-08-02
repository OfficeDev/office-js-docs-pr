---
title: Enable and Disable Add-in Commands
description: Learn how to change the enabled or disabled status of custom ribbon buttons and menu items in your Office Web Add-in.
ms.date: 03/11/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Change the availability of add-in commands

When some functionality in your add-in should only be available in certain contexts, you can programmatically configure your custom add-in commands to only be available in these contexts. For example, a function that changes the header of a table should only be available when the cursor is in a table.

> [!NOTE]
>
> - This article assumes that you're familiar with the [basic concepts for add-in commands](add-in-commands.md). Please review it if you haven't worked with add-in commands (custom menu items and ribbon buttons) recently.

## Supported capabilities

You can programmatically change the availability of an add-in command for the following capabilities.

- Ribbon buttons, menus, and tabs.
- Context menu items.

## Office application and requirement set support

The following table outlines the Office applications that support configuring the availability of add-in commands. It also lists the requirement sets needed to use the API.

| Add-in command capability | Requirement set | Supported Office applications |
| ---- | ---- | ---- |
| Ribbon buttons, menus, and tabs | [RibbonApi 1.1](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets) | <ul><li>Excel</li><li>PowerPoint</li><li>Word</li></ul> |
| Context menu items | [ContextMenuApi 1.1](/javascript/api/requirement-sets/common/context-menu-api-requirement-sets) | <ul><li>Excel</li><li>PowerPoint</li><li>Word</li></ul> |

> [!TIP]
> To learn how to test for platform support with requirement sets, see [Office versions and requirement sets](../develop/office-versions-and-requirement-sets.md).

## Configure a shared runtime

To change the availability of a ribbon or context menu control or item, the manifest of your add-in must first be configured to use a [shared runtime](../testing/runtimes.md#shared-runtime). For guidance on how to set up a shared runtime, see [Configure your Office Add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## Programmatically change the availability of an add-in command

# [Ribbon](#tab/ribbon)

### Deactivate ribbon controls at launch

> [!NOTE]
> Only the controls on the ribbon can be deactivated when the Office application starts. You can't deactivate custom controls added to a context menu at launch.

By default, a custom button or menu item on the ribbon is available for use when the Office application launches. To deactivate it when Office starts, you must specify this in the manifest. The process depends on which type of manifest your add-in uses.

- [Unified manifest for Microsoft 365](#unified-manifest-for-microsoft-365)
- [Add-in only manifest](#add-in-only-manifest)

#### Unified manifest for Microsoft 365

[!include[Unified manifest host application support note](../includes/unified-manifest-support-note.md)]

Just add an [`"enabled"`](/microsoft-365/extensibility/schema/extension-common-custom-group-controls-item#enabled) property with the value `false` to the control or menu item object. The following shows the basic structure.

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
                                        "enabled": false
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

#### Add-in only manifest

Just add an [Enabled](/javascript/api/manifest/enabled) element immediately *below* (not inside) the [Action](/javascript/api/manifest/action) element of the control item. Then, set its value to `false`.

The following shows the basic structure of a manifest that configures the `<Enabled>` element.

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
                <Control ... id="Contoso.MyButton3">
                  ...
                  <Action ...>
                  <Enabled>false</Enabled>
...
</OfficeApp>
```

### Change the availability of a ribbon control

To update the availability of a button or menu item on the ribbon, perform the following steps.

1. Create a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that specifies the following:
    - The IDs of the command, including its parent group and tab. The IDs must match those declared in the manifest.
    - The availability status of the command.
1. Pass the **RibbonUpdaterData** object to the [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestupdate-member(1)) method.

The following is a simple example. Note that "MyButton", "OfficeAddinTab1", and "CustomGroup111" are copied from the manifest.

```javascript
function enableButton() {
    const ribbonUpdaterData = {
        tabs: [
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
            }
        ]
    };

    Office.ribbon.requestUpdate(ribbonUpdaterData);
}
```

There are several interfaces (types) to make it easier to construct the **RibbonUpdateData** object.

- [Office.Control](/javascript/api/office/office.control)
- [Office.Group](/javascript/api/office/office.group)
- [Office.Tab](/javascript/api/office/office.tab)

The following is the equivalent example in TypeScript and it makes use of these types.

```typescript
const enableButton = async () => {
    const button: Control = { id: "MyButton", enabled: true };
    const parentGroup: Group = { id: "CustomGroup111", controls: [button] };
    const parentTab: Tab = { id: "OfficeAddinTab1", groups: [parentGroup] };
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab] };
    Office.ribbon.requestUpdate(ribbonUpdater);
}
```

> [!TIP]
> You can `await` the call of **requestUpdate()** if the parent function is asynchronous, but note that the Office application controls when it updates the state of the ribbon. The **requestUpdate()** method queues a request to update. The method will resolve the promise object as soon as it has queued the request, not when the ribbon actually updates.

### Toggle tab visibility and the enabled status of a button at the same time

The **requestUpdate** method is also used to toggle the visibility of a custom contextual tab. For details about this and example code, see [Create custom contextual tabs in Office Add-ins](contextual-tabs.md#toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time).

# [Context menu](#tab/context-menu)

To programmatically change the availability of a custom item on a context menu, perform the following steps.

1. Create a [ContextMenuUpdaterData](/javascript/api/office/office.contextmenuupdaterdata) object. This object defines the context menu controls you want to update. It specifies the following:
    - The ID of each context menu item you want to update. The ID must match that specified in the manifest.
    - The availability status of the control.
1. Send your request to update the context menu by passing the **ContextMenuUpdaterData** object to [Office.contextMenu.requestUpdate](/javascript/api/office/office.contextmenu#office-office-contextmenu-requestupdate-member(1)).

    > [!IMPORTANT]
    > The Office application controls when it updates the state of the context menu. The **requestUpdate()** method queues an update request and resolves the promise object as soon as it queues the request, not when the ribbon actually updates.

The following is an example of how to change the availability of custom buttons on a context menu by configuring the `enabled` property of each button.

```javascript
await Office.contextMenu.requestUpdate({
    controls: [
        {
            id: Addin.CtxMenu.Button1,
            enabled: true
        },
        {
            id: Addin.CtxMenu.Button2,
            enabled: false
        },
    ]
});
```

---

## Change the state in response to an event

A common scenario in which the state of a ribbon or context menu control should change is when a user-initiated event changes the add-in context. Consider a scenario in which a button should be available when, and only when, a chart is activated. Although the following example uses ribbon controls, a similar implementation can be applied to custom items on a context menu.

1. First, set the `<Enabled>` element for the button in the manifest to `false`. For guidance on how to configure this, see [Deactivate ribbon controls at launch](#deactivate-ribbon-controls-at-launch).
1. Then, assign handlers. This is commonly done in the **Office.onReady** function as in the following example. In the example, handlers (created in a later step) are assigned to the **onActivated** and **onDeactivated** events of all the charts in an Excel worksheet.

    ```javascript
    Office.onReady(async () => {
        await Excel.run((context) => {
            const charts = context.workbook.worksheets
                .getActiveWorksheet()
                .charts;
            charts.onActivated.add(enableChartFormat);
            charts.onDeactivated.add(disableChartFormat);
            return context.sync();
        });
    });
    ```

1. Define the `enableChartFormat` handler. The following is a simple example. For a more robust way of changing a control's status, see [Best practice: Test for control status errors](#best-practice-test-for-control-status-errors).

    ```javascript
    function enableChartFormat() {
        const button =
            {
                id: "ChartFormatButton",
                enabled: true
            };
        const parentGroup =
            {
                id: "MyGroup",
                controls: [button]
            };
        const parentTab =
            {
                id: "CustomChartTab",
                groups: [parentGroup]
            };
        const ribbonUpdater = { tabs: [parentTab] };
        Office.ribbon.requestUpdate(ribbonUpdater);
    }
    ```

1. Define the `disableChartFormat` handler. It's identical to the `enableChartFormat` handler, except that the **enabled** property of the button object is set to `false`.

## Best practice: Test for control status errors

In some circumstances, the ribbon or context menu doesn't repaint after `requestUpdate` is called, so the control's clickable status doesn't change. For this reason it's a best practice for the add-in to keep track of the status of its controls. The add-in should conform to the following rules.

- Whenever `requestUpdate` is called, the code should record the intended state of the custom buttons and menu items.
- When a custom control is selected, the first code in the handler should check to see if the button should have been available. If it shouldn't have been available, the code should report or log an error and try again to set the buttons to the intended state.

The following example shows a function that deactivates a button on the ribbon and records the button's status. In this example, `chartFormatButtonEnabled` is a global boolean variable that's initialized to the same value as the [Enabled](/javascript/api/manifest/enabled) element for the button in the add-in's manifest. Although the example uses a ribbon button, a similar implementation can be applied to custom items on a context menu.

```javascript
function disableChartFormat() {
    const button =
    {
        id: "ChartFormatButton",
        enabled: false
    };
    const parentGroup =
    {
        id: "MyGroup",
        controls: [button]
    };
    const parentTab =
    {
        id: "CustomChartTab",
        groups: [parentGroup]
    };
    const ribbonUpdater = { tabs: [parentTab] };
    Office.ribbon.requestUpdate(ribbonUpdater);

    chartFormatButtonEnabled = false;
}
```

The following example shows how the button's handler tests for an incorrect state of the button. Note that `reportError` is a function that shows or logs an error.

```javascript
function chartFormatButtonHandler() {
    if (chartFormatButtonEnabled) {

        // Do work here.

    } else {
        // Report the error and try to make the button unavailable again.
        reportError("That action is not possible at this time.");
        disableChartFormat();
    }
}
```

## Error handling

In some scenarios, Office is unable to update the ribbon or context menu and will return an error. For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened. Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`. The following is an example of how to handle this error. In this case, the `reportError` method displays the error to the user. Although the example uses a ribbon button, a similar implementation can be applied to custom items on a context menu.

```javascript
function disableChartFormat() {
    try {
        const button =
        {
            id: "ChartFormatButton",
            enabled: false
        };
        const parentGroup =
        {
            id: "MyGroup",
            controls: [button]
        };
        const parentTab =
        {
            id: "CustomChartTab",
            groups: [parentGroup]
        };
        const ribbonUpdater = { tabs: [parentTab] };
        Office.ribbon.requestUpdate(ribbonUpdater);

        chartFormatButtonEnabled = false;
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, close the Office application, and restart it.");
        }
    }
}
```

## See also

- [Add-in commands](add-in-commands.md)
- [Create add-in commands with the add-in only manifest](../develop/create-addin-commands.md)
- [Create custom contextual tabs in Office Add-ins](contextual-tabs.md)
