---
title: Enable and Disable Add-in Commands
description: Learn how to change the enabled or disabled status of custom ribbon buttons and menu items in your Office Web Add-in.
ms.date: 10/08/2024
ms.topic: how-to
ms.localizationpriority: medium
---

# Enable and Disable Add-in Commands

When some functionality in your add-in should only be available in certain contexts, you can programmatically enable or disable your custom Add-in Commands. For example, a function that changes the header of a table should only be enabled when the cursor is in a table.

You can also specify whether the command is enabled or disabled when the Office client application opens.

> [!NOTE]
>
> - This article assumes that you're familiar with the [basic concepts for Add-in Commands](add-in-commands.md). Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.
>
> - Programmatically enabling or disabling [context menus](interface-elements.md#add-in-commands) isn't supported. Only ribbon buttons, menus, and tabs are supported.

## Office application and platform support

The APIs described in this article are available in **Excel**, **PowerPoint**, and **Word** as part of the [RibbonApi 1.1](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets) requirement set. To learn how to test for platform support with requirement sets, see [Office versions and requirement sets](../develop/office-versions-and-requirement-sets.md).

## Shared runtime required

The APIs and manifest markup described in this article require that the add-in's manifest specify that it should use a [shared runtime](../testing/runtimes.md#shared-runtime). To do this, take the following steps.

1. In the [Runtimes](/javascript/api/manifest/runtimes) element in the manifest, add the following child element: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`. (If there isn't already a **\<Runtimes\>** element in the manifest, create it as the first child under the **\<Host\>** element in the **\<VersionOverrides\>** section.)
1. In the [Resources.Urls](/javascript/api/manifest/resources) section of the manifest, add the following child element: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, where `{MyDomain}` is the domain of the add-in and `{path-to-start-page}` is the path for the start page of the add-in; for example: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.
1. Depending on whether your add-in contains a task pane, a function file, or an Excel custom function, you must do one or more of the following three steps.

    - If the add-in contains a task pane, set the `resid` attribute of the [Action](/javascript/api/manifest/action).[SourceLocation](/javascript/api/manifest/sourcelocation) element to exactly the same string as you used for the `resid` of the **\<Runtime\>** element in step 1; for example, `Contoso.SharedRuntime.Url`. The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.
    - If the add-in contains an Excel custom function, set the `resid` attribute of the [Page](/javascript/api/manifest/page).[SourceLocation](/javascript/api/manifest/sourcelocation) element exactly the same string as you used for the `resid` of the **\<Runtime\>** element in step 1; for example, `Contoso.SharedRuntime.Url`. The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.
    - If the add-in contains a function file, set the `resid` attribute of the [FunctionFile](/javascript/api/manifest/functionfile) element to exactly the same string as you used for the `resid` of the **\<Runtime\>** element in step 1; for example, `Contoso.SharedRuntime.Url`. The element should look like this: `<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.

## Set the default state to disabled

By default, any Add-in Command is enabled when the Office application launches. If you want a custom button or menu item to be disabled when the Office application launches, you specify this in the manifest. Just add an [Enabled](/javascript/api/manifest/enabled) element (with the value `false`) immediately *below* (not inside) the [Action](/javascript/api/manifest/action) element in the declaration of the control. The following shows the basic structure.

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

## Change the state programmatically

The essential steps to changing the enabled status of an Add-in Command are:

1. Create a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the command, and its parent group and tab, by their IDs as declared in the manifest; and (2) specifies the enabled or disabled state of the command.
1. Pass the **RibbonUpdaterData** object to the [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestupdate-member(1)) method.

The following is a simple example. Note that "MyButton", "OfficeAddinTab1", and "CustomGroup111" are copied from the manifest.

```javascript
function enableButton() {
    Office.ribbon.requestUpdate({
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
    });
}
```

We also provide several interfaces (types) to make it easier to construct the **RibbonUpdateData** object. The following is the equivalent example in TypeScript and it makes use of these types.

```typescript
const enableButton = async () => {
    const button: Control = {id: "MyButton", enabled: true};
    const parentGroup: Group = {id: "CustomGroup111", controls: [button]};
    const parentTab: Tab = {id: "OfficeAddinTab1", groups: [parentGroup]};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);
}
```

You can `await` the call of **requestUpdate()** if the parent function is asynchronous, but note that the Office application controls when it updates the state of the ribbon. The **requestUpdate()** method queues a request to update. The method will resolve the promise object as soon as it has queued the request, not when the ribbon actually updates.

## Change the state in response to an event

A common scenario in which the ribbon state should change is when a user-initiated event changes the add-in context.

Consider a scenario in which a button should be enabled when, and only when, a chart is activated. The first step is to set the [Enabled](/javascript/api/manifest/enabled) element for the button in the manifest to `false`. See above for an example.

Second, assign handlers. This is commonly done in the **Office.onReady** function as in the following example which assigns handlers (created in a later step) to the **onActivated** and **onDeactivated** events of all the charts in the worksheet.

```javascript
Office.onReady(async () => {
    await Excel.run(context => {
        const charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(enableChartFormat);
        charts.onDeactivated.add(disableChartFormat);
        return context.sync();
    });
});
```

Third, define the `enableChartFormat` handler. The following is a simple example, but see [Best practice: Test for control status errors](#best-practice-test-for-control-status-errors) below for a more robust way of changing a control's status.

```javascript
function enableChartFormat() {
    const button = {
                  id: "ChartFormatButton", 
                  enabled: true
                 };
    const parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    const parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    const ribbonUpdater = {tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);
}
```

Fourth, define the `disableChartFormat` handler. It would be identical to `enableChartFormat` except that the **enabled** property of the button object would be set to `false`.

### Toggle tab visibility and the enabled status of a button at the same time

The **requestUpdate** method is also used to toggle the visibility of a custom contextual tab. For details about this and example code, see [Create custom contextual tabs in Office Add-ins](contextual-tabs.md#toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time).

## Best practice: Test for control status errors

In some circumstances, the ribbon does not repaint after `requestUpdate` is called, so the control's clickable status does not change. For this reason it is a best practice for the add-in to keep track of the status of its controls. The add-in should conform to the following rules.

1. Whenever `requestUpdate` is called, the code should record the intended state of the custom buttons and menu items.
1. When a custom control is clicked, the first code in the handler, should check to see if the button should have been clickable. If shouldn't have been, the code should report or log an error and try again to set the buttons to the intended state.

The following example shows a function that disables a button and records the button's status. Note that `chartFormatButtonEnabled` is a global boolean variable that is initialized to the same value as the [Enabled](/javascript/api/manifest/enabled) element for the button in the manifest.

```javascript
function disableChartFormat() {
    const button = {
                  id: "ChartFormatButton", 
                  enabled: false
                 };
    const parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    const parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    const ribbonUpdater = {tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);

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
        const button = {
                      id: "ChartFormatButton", 
                      enabled: false
                     };
        const parentGroup = {
                           id: "MyGroup",
                           controls: [button]
                          };
        const parentTab = {
                         id: "CustomChartTab", 
                         groups: [parentGroup]
                        };
        const ribbonUpdater = {tabs: [parentTab]};
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
