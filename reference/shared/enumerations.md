
# Enumerations

You can specify an enumerated value by using either its fully qualified enumeration name ( `Office.CoercionType.Text`) or its corresponding text value ( `"text"`). For example, the following method call uses enumeration names:


```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, {valueFormat:Office.ValueFormat.Unformatted, filterType:Office.FilterType.All},
   function (result) {
      if (result.status === Office.AsyncResultStatus.Success)
         var dataValue = result.value; // Get selected data.
         write('Selected data is ' + dataValue);
      else {
         var err = result.error;
         write(err.name + ": " + err.message);
      }
   });

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```


Here's the same call using the enumeration text values:




```js
Office.context.document.getSelectedDataAsync("text", {valueFormat:"unformatted", filterType:"all"},
   function (result) {
      if (result.status === "success")
         var dataValue = result.value; // Get selected data.
         write('Selected data is ' + dataValue);
      else {
         var err = result.error;
         write(err.name + ": " + err.message);
      }
   });
```


## Reference



|**Name**|**Definition**|
|:-----|:-----|
|[ActiveView](activeview-enumeration.md)|Specifies the state of the active view of the document, for example, whether the user can edit the document.|
|[AsyncResultStatus](asyncresultstatus-enumeration.md)|Specifies the result of an asynchronous call.|
|[AttachmentType](http://msdn.microsoft.com/library/83883a47-a937-4afb-a55e-e789057335c4%28Office.15%29.aspx)|Specifies the type of an attachment to an email message or meeting request. Outlook 2013 does not support this enumeration.|
|[BindingType](bindingtype-enumeration.md)|Specifies the type of the binding object that should be returned.|
|[BodyType](http://msdn.microsoft.com/library/31350fe6-4c42-4cbb-a5b2-4fb2d360fa11%28Office.15%29.aspx)|Specifies the text type for the body of an appointment or message.|
|[CoercionType](coerciontype-enumeration.md)|Specifies how to coerce data returned or set by the invoked method.|
|[CustomXMLNodeType](customxmlnodetype-enumeration.md)|Specifies the node type.|
|[DocumentMode](documentmode-enumeration.md)|Specifies whether the document in associated application is read-only or read-write. |
|[EntityType](http://msdn.microsoft.com/library/0035be38-8a65-4693-bcc4-0a8dd7b1495b%28Office.15%29.aspx)|Specifies an entity's type.|
|[EventType](eventtype-enumeration.md)|Specifies the kind of event that was raised.|
|[FileType](filetype-enumeration.md)|Specifies the format in which to return the document.|
|[GoToType](gototype-enumeration.md)|Specifies the type of place or object to navigate to.|
|[FilterType](filtertype-enumeration.md)|Specifies whether filtering from the host application is applied when the data is retrieved.|
|[InitializationReason](initializationreason-enumeration.md)|Specifies whether the add-in was just inserted or was already contained in the document.|
|[ItemType](http://msdn.microsoft.com/library/e0bb23fd-f360-4b0f-b72c-1cf08d4cab3f%28Office.15%29.aspx)|Specifies an item's type.|
|[notificationMessageType](http://msdn.microsoft.com/library/ff00c89d-0019-4545-a95b-7ed0db712ce9%28Office.15%29.aspx)|Specifies the notification message for an appointment or message.|
|[ProjectProjectFields](projectprojectfields-enumeration.md)|Specifies the project fields that are available as a parameter for the [getProjectFieldAsync](projectdocument.getprojectfieldasync.md) method.|
|[ProjectResourceFields](projectresourcefields-enumeration.md)|Specifies the resource fields that are available as a parameter for the [getResourceFieldAsync](projectdocument.gettaskfieldasync.md) method.|
|[ProjectTaskFields](projecttaskfields-enumeration.md)|Specifies the task fields that are available as a parameter for the [getTaskFieldAsync](projectdocument.gettaskfieldasync.md) method.|
|[ProjectViewTypes](projectviewtypes-enumeration.md)|Specifies the types of views that the [getSelectedViewAsync](projectdocument.getselectedviewasync.md) method can recognize.|
|[RecipientType](http://msdn.microsoft.com/library/6e7c4029-6e52-47f6-98d2-4cd3ce7bd8b4%28Office.15%29.aspx)|Specifies the type of recipient for an appointment.|
|[ResponseType](http://msdn.microsoft.com/library/b3e723ca-4be0-4846-ad97-0eecab4355eb%28Office.15%29.aspx)|Specifies the response to a meeting invitation.|
|[SelectionMode](selectionmode-enumeration.md)|Specifies whether to select (highlight) the location to navigate to (when using the [Document.goToByIdAsync](document.gotobyidasync.md) method).|
|[SourceProperty](http://msdn.microsoft.com/library/6a209a7f-57cd-4dc3-869e-07b0f5928b28%28Office.15%29.aspx)|Specifies the source of the data returned by the invoked method.|
|[Table](table-enumeration.md)|Specifies enumerated values for the  `cells:` property in the _cellFormat_ parameter of [table formatting methods](../../docs/excel/format-tables-in-add-ins-for-excel.md).|
|[ValueFormat](valueformat-enumeration.md)|Specifies whether values, such as numbers and dates, returned by the invoked method are returned with their formatting applied.|

## Support details


Support for each enumeration differs across Office host applications. See the "Support details" section of each enumerations's topic for host support information.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


|||
|:-----|:-----|
|**Add-in types**|Content, task pane, Outlook|
|**Library**|Office.js|
|**Namespace**|Office|
