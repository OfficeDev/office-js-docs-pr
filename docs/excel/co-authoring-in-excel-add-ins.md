---
title: Coauthoring in Excel add-ins
description: Learn to coauthor an Excel workbook stored in OneDrive, OneDrive for Business, or SharePoint Online.
ms.date: 07/08/2021
ms.localizationpriority: medium
---


# Coauthoring in Excel add-ins  

With [coauthoring](https://support.microsoft.com/office/7152aa8b-b791-414c-a3bb-3024e46fb104), multiple people can work together and edit the same Excel workbook simultaneously. All coauthors of a workbook can see another coauthor's changes as soon as that coauthor saves the workbook. To coauthor an Excel workbook, the workbook must be stored in OneDrive, OneDrive for Business, or SharePoint Online.

> [!IMPORTANT]
> In Excel for Microsoft 365, you will notice AutoSave in the upper-left corner. When AutoSave is turned on, coauthors see your changes in real time. Consider the impact of this behavior on the design of your Excel add-in. Users can turn off AutoSave via the switch in the upper left of the Excel window.

## Coauthoring overview

When you change a workbook's content, Excel automatically synchronizes those changes across all coauthors. Coauthors can change the content of a workbook, but so can code running within an Excel add-in. For example, when the following JavaScript code runs in an Office Add-in, the value of a range is set to Contoso.

```js
range.values = [['Contoso']];
```

After 'Contoso' synchronizes across all coauthors, any user or add-in running in the same workbook will see the new value of the range.

Coauthoring only synchronizes the content within the shared workbook. Values copied from the workbook to JavaScript variables in an Excel add-in are not synchronized. For example, if your add-in stores the value of a cell (such as 'Contoso') in a JavaScript variable, and then a coauthor changes the value of the cell to 'Example', after synchronization all coauthors see 'Example' in the cell. However, the value of the JavaScript variable is still set to 'Contoso'. Furthermore, when multiple coauthors use the same add-in, each coauthor has their own copy of the variable, which is not synchronized. When you use variables that use workbook content, be sure you check for updated values in the workbook before you use the variable.

## Use events to manage the in-memory state of your add-in

Excel add-ins can read workbook content (from hidden worksheets and a setting object), and then store it in data structures such as variables. After the original values are copied into any of these data structures, coauthors can update the original workbook content. This means that the copied values in the data structures are now out of sync with the workbook content. When you build your add-ins, be sure to account for this separation of workbook content and values stored in data structures.

For example, you might build a content add-in that displays custom visualizations. The state of your custom visualizations might be saved in a hidden worksheet. When coauthors use the same workbook, the following scenario can occur.

- User A opens the document and the custom visualizations are shown in the workbook. The custom visualizations read data from a hidden worksheet (for example, the color of the visualizations is set to blue).
- User B opens the same document, and starts modifying the custom visualizations. User B sets the color of the custom visualizations to orange. Orange is saved to the hidden worksheet.
- User A's hidden worksheet is updated with the new value of orange.
- User A's custom visualizations are still blue.

If you want User A's custom visualizations to respond to changes made by coauthors on the hidden worksheet, use the [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) event. This ensures that changes to workbook content made by coauthors is reflected in the state of your add-in.

## Caveats to using events with coauthoring

As described earlier, in some scenarios, triggering events for all coauthors provides an improved user experience. However, be aware that in some scenarios this behavior can produce poor user experiences.

For example, in data validation scenarios, it is common to display UI in response to events. The [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) event described in the previous section runs when either a local user or coauthor (remote) changes the workbook content within the binding. If the event handler of the `BindingDataChanged` event displays UI, users will see UI that is unrelated to changes they were working on in the workbook, leading to a poor user experience. Avoid displaying UI when using events in your add-in.

## Avoid table row coauthoring conflicts

It is a known issue that calls to the [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1)) API can cause coauthoring conflicts. We do not recommend using that API if you anticipate your add-in will be run while other users are editing the add-in's workbook (specifically, if they are editing the table or any range under the table). The following guidance should help you avoid issues with the `TableRowCollection.add` method (and avoid triggering the yellow bar Excel shows that asks users to refresh).

1. Use [`Range.values`](/javascript/api/excel/excel.range#excel-excel-range-values-member) instead of [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1)). Setting the `Range` values directly below the table automatically expands the table. Otherwise, adding table rows through the `Table` APIs results in merge conflicts for coauth users.
1. There should be no [data validation rules](https://support.microsoft.com/office/29fecbcc-d1b9-42c1-9d76-eff3ce5f7249) applied to cells below the table, unless the data validation is applied to the entire column.
1. If there is data under the table, the add-in needs to handle that before setting the range value. Using [`Range.insert`](/javascript/api/excel/excel.range#excel-excel-range-insert-member(1)) to insert an empty row will move the data and make space for the expanding table. Otherwise, you risk overwriting cells below the table.
1. You cannot add an empty row to a table with `Range.values`. The table only automatically expands if data is present in the cells directly below the table. Use either temporary data or hidden columns as a workaround to add an empty table row.

## See also

- [About coauthoring in Excel (VBA)](/office/vba/excel/concepts/about-coauthoring-in-excel)
- [How AutoSave impacts add-ins and macros (VBA)](/office/vba/library-reference/concepts/how-autosave-impacts-addins-and-macros)
