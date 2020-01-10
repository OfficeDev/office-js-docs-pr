---
title: Enable and Disable Add-in Commands
description: ''
ms.date: 01/10/2020
localization_priority: Priority
---


# Enable and Disable Add-in Commands (preview)

> [!NOTE]
> This article assumes that you are familiar with the following documentation. Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.
>
> [Basic concepts for Add-in Commands](add-in-commands.md)

## Preview status

The APIs described in this article are in preview and are currently limited in the following ways:

- This feature is only available in Excel.
- These APIs and manifest markup only affect the single-line ribbon. They have no effect on the multiline ribbon.
- These APIs and manifest markup only work when the add-in's manifest specifies that it should use a shared runtime. To do this, add the following [Set](/office/dev/add-ins/reference/manifest/set) to the [Requirements](/office/dev/add-ins/reference/manifest/requirements).[Sets](/office/dev/add-ins/reference/manifest/sets) section of the manifest: `<Set Name="CustomFunctionsRuntime" MinVersion="1.2"/>`.

> [!NOTE]
> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

When some functionality in your add-in should only be available in certain contexts, you can programmatically enable or disable your custom Add-in Commands. For example, a function that changes the header of a table should only be enabled when the cursor is in a table.

You can also specify whether the command is enabled or disabled when your add-in launches.

## Set the default state to disabled

By default, any Add-in Command is enabled when the add-in launches. If you want a custom button or menu item to be disabled when the add-in launches, you specify this in the manifest. Just add an [Enabled](/office/dev/add-ins/reference/manifest/enabled) element (with the value `false`) immediately below the [Action](/office/dev/add-ins/reference/manifest/action) element in the declaration of the control. The following shows the basic structure:

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
                <Control ... id="MyButton"> 
                  ... 
                  <Action ...> 
                  <Enabled>false</Enabled> 
... 
</OfficeApp>
```

## Change the state programmatically

The essential steps to changing the enabled status of an Add-in Command are:

1. Create a [RibbonUpdaterData](/javascript/api/office/officeruntime.ribbonupdaterdata) object that (1) specifies the command, and its parent tab, by their IDs as specified in the manifest; and (2) specifies the enabled or disabled state of the command.
2. Get a reference to the [Ribbon](/javascript/api/office/officeruntime.ribbon) object with the [OfficeRuntime.ui.getRibbon](/javascript/api/office/officeruntime.ui.getribbon) method.
3. Pass the **RibbonUpdaterData** object to the **Ribbon.requestUpdate()** method.

The following is a simple example. Note that "MyButton" and "OfficeAddinTab1" are copied from the manifest.

```javascript
function enableButton() {
    OfficeRuntime.ui.getRibbon()
        .then(function (ribbon) { 
            ribbon.requestUpdate({
                tabs: [
                    {
                        id: "OfficeAppTab1", 
                        controls: [
                        {
                            id: "MyButton", 
                            enabled: true
                        }
                    ]}
                ]}); 
        }); 
}
```

We also provide several interfaces (types) to make it easier to construct the **RibbonUpdateData** object. The following is the equivalent example in TypeScript and it makes use of these types. All of the types are in the **OfficeRuntime** namespace.

```javascript
const enableButton = async () => { 
    const button: Control = {id: "MyButton", enabled: true}; 
    const parentTab: Tab = {id: "OfficeAddinTab1", controls: [button]}; 
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab]};  
    const ribbon: Ribbon = await OfficeRuntime.ui.getRibbon(); 
    await ribbon.requestUpdate(ribbonUpdater); 
}
```

Office controls when it updates the state of the ribbon. The **requestUpdate()** method queues a request to update. The method will resolve the Promise object as soon as it has queued the request, not when the ribbon actually updates.

## Change the state in response to an event

A common scenario in which the ribbon state should change is when a user-initiated event changes the add-in context.

Consider a scenario in which a button should be enabled when, and only when, a chart is activated. The first step is to set the [Enabled](/office/dev/add-ins/reference/manifest/enabled) element for the button in the manifest to `false`. See above for an example.

Second, assign handlers. This is commonly done in the **Office.onReady** method as in the following example which assigns handlers (created in a later step) to the **onActivated** and **onDeactivated** events of all the charts in the worksheet.

```javascript
Office.onReady(async () => {
    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(enableChartFormat);
        charts.onDeactivated.add(disableChartFormat);
        return context.sync();
    });
});
```

Third, define the `enableChartFormat` handler. The following is a simple example, but see **Best practice: Test for control status errors** below for a more robust way of changing a control's status.

```javascript
function enableChartFormat() {
    OfficeRuntime.ui.getRibbon()
        .then(function (ribbon) {
            var button = {id: "ChartFormatButton", enabled: true};
            var parentTab = {id: "CustomChartTab", controls: [button]};
            var ribbonUpdater = {tabs: [parentTab]};
            await ribbon.requestUpdate(ribbonUpdater);
        });
}
```

Fourth, define the `disableChartFormat` handler. It would be identical to `enableChartFormat` except that the **enabled** property of the button object would be set to `false`.

## Best practice: Test for control status errors

In some circumstances, the ribbon does not repaint after `requestUpdate` is called, so the control's clickable status does not change. For this reason it is a best practice for the add-in to keep track of the status of its controls. The add-in should conform to these rules:

1. Whenever `requestUpdate` is called, the code should record the intended state of the custom buttons and menu items.
2. When a custom control is clicked, the first code in the handler that runs, should check to see if the button should have been clickable. If shouldn't have been, the code should report or log an error and try again to set the buttons to the intended state.

The following example shows an example of a function that disables a button and records the button's status. About this code, note:

- `chartFormatButtonEnabled` is a global boolean variable that is initialized to the same value as the [Enabled](/office/dev/add-ins/reference/manifest/enabled) element for the button in the manifest.

```javascript
function disableChartFormat() {
    OfficeRuntime.ui.getRibbon()
        .then(function (ribbon) {
            var button = {id: "ChartFormatButton", enabled: false};
            var parentTab = {id: "CustomChartTab", controls: [button]};
            var ribbonUpdater = {tabs: [parentTab]};
            await ribbon.requestUpdate(ribbonUpdater);

            chartFormatButtonEnabled = false;
        });
}

```

The following example shows how the button's handler tests for an incorrect state of the button. About this code, note:

- `reportError` is a function that shows or logs an error.

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
