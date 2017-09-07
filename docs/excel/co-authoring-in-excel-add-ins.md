# Co-authoring in Excel add-ins  

Learn about co-authoring in Excel, and about development guidelines to follow to support co-authoring in your Excel add-in.

With [co-authoring](https://support.office.com/en-US/article/Collaborate-on-Excel-workbooks-at-the-same-time-with-co-authoring-7152aa8b-b791-414c-a3bb-3024e46fb104 "co-authoring"), multiple people can work together and edit the same Excel workbook simultaneously. All co-authors of a workbook can see another co-author’s changes as soon as that co-author saves the workbook. To co-author an Excel workbook, it must be stored in OneDrive, OneDrive for Business, or SharePoint Online.

> Important: In Excel 2016 for Office 365, you will notice AutoSave in the upper-left corner. When AutoSave is turned on, co-authors see your changes in real-time. Ensure you consider the impact of this behavior on the design of your Excel add-in. Users can turn off AutoSave by simply pulling the switch on the upper left of the Excel window.

Co-authoring is available in Excel Online, Excel for Android, Excel for iOS, Excel Mobile for Windows 10, and Excel for Windows Desktop for Office 365 customers (Windows desktop build 16.0.8326.2076 or later, which is available to current channel customers from August 2017).

## Learn co-authoring concepts ##
 
When you change the workbook’s content, Excel automatically synchronizes those changes across all co-authors. Co-authors can change the content of a workbook, but so can code running within an Excel add-in. For example, when the following JavaScript code runs in an Office add-in, the value of a range is set to Contoso:


    range.values = [[‘Contoso’]];

After ‘Contoso’ synchronizes across all co-authors, any user or add-in running in the same workbook will then see the new value of the range. 

Co-authoring only synchronizes the content within the shared workbook. Values copied from the workbook to JavaScript variables in an Excel add-in are not synchronized. For example, if your add-in stores the value of a cell (for example, ‘Contoso’) in a JavaScript variable, and then a co-author changes the value of the cell to ‘Foo’, after synchronization all co-authors see ‘Foo’ in the cell. However, the value of the JavaScript variable is still set to ‘Contoso’. Furthermore, when multiple co-authors use the same add-in, each co-author has their own copy of the variable, which is not synchronized. When using variables that use workbook content, be sure you check for updated values in the workbook before using the variable. 

## Use events to manage the in-memory state of your add-in ##
 
Excel add-ins can read workbook content (from hidden worksheets and a setting object), and then store it in data structures such as variables. After the original values are copied into any of these data structures, co-authors can update the original workbook content. This means that the copied values in the data structures are now out-of-sync with the workbook content. When building add-ins, ensure that you cater for this separation of workbook content and values stored in data structures.

For example, you may build a content add-in that displays custom visualizations. The state of your custom visualizations may be saved in a hidden worksheet. When co-authors use the same workbook, the following scenario can occur:



- User A opens the document and the custom visualizations are shown in the workbook. The custom visualizations read data from a hidden worksheet (for example, the color of the visualizations is set to blue).
- User B opens the same document, and starts modifying the custom visualizations. User B sets the color of the custom visualizations to orange. Orange is saved to the hidden worksheet.
- User A’s hidden worksheet is updated with the new value of orange.
- User A’s custom visualizations are still blue. 

If you want User A’s custom visualizations to respond to changes made by co-authors on the hidden worksheet, use the [BindingDataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) event. This ensures that changes to workbook content made by co-authors is reflected in the state of your add-in.

## Caveats of using events with co-authoring ##

As described above, there are some scenarios where triggering events for all co-authors provides an improved user experience. However, be aware there are scenarios where this behavior can produce poor user experiences. 

For example, in data validation scenarios, it is common to display UI in response to events. The [BindingDataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) event described in the previous section runs when either a local user or co-author (remote) changes the workbook content within the binding. If the event handler of the BindingDataChanged event displays UI, users will see UI that is unrelated to changes they were working on in the workbook, leading to a poor user experience. Avoid displaying UI when using events in your add-in.

## See Also ##
[About co-authoring in Excel (VBA)](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/about-coauthoring-in-excel) 

[How AutoSave impacts add-ins and macros (VBA)](https://msdn.microsoft.com/en-us/vba/office-shared-vba/articles/how-autosave-impacts-addins-and-macros) 
