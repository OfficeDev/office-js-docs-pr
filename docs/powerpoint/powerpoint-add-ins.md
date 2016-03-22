
# Create content and task pane add-ins for PowerPoint

The code examples in the article show you some basic tasks for developing PowerPoint content add-ins. To display information, these examples depend on the  `app.showNotification` function, which is included in the Visual StudioOffice Add-ins project templates. If you aren't using Visual Studio to develop your add-in, you'll need replace the `showNotification` function with your own code. Several of these examples also depend on this `globals` object that is declared outside of the scope of these functions: `var globals = {activeViewHandler:0, firstSlideId:0};`

These code examples require your project to [reference Office.js v1.1 library or later](../../docs/develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).


## Detect the presentation's active view and handle the ActiveViewChanged event

The  `getFileView` function calls the [Document.getActiveViewAsync](../../reference/shared/document.getactiveviewasync.md) method to return whether the presentation's current view is "edit" (any of the views in which you can edit slides, such as **Normal** or **Outline View**) or "read" ( **Slide Show** or **Reading View**) view.


```js
function getFileView() {
    //Gets whether the current view is edit or read.
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification(asyncResult.value);
        }
    });
}
```

The  `registerActiveViewChanged` function calls the [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) method to register a handler for the [Document.ActiveViewChanged](../../reference/shared/document.activeviewchanged.md) event. After executing this function, when you change the view of the presentation, the `app.showNotification` notification will display the active view mode ("read" or "edit").




```js
function registerActiveViewChanged() {
    Globals.activeViewHandler = function (args) {
        app.showNotification(JSON.stringify(args));
    }

    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, Globals.activeViewHandler, 
        function (asyncResult) {
            if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
            else {
            app.showNotification(asyncResult.status);
            }
        });
}
```


## Get the URL of the presentation

The  `getFileUrl` function calls the [Document.getFileProperties](../../reference/shared/document.getfilepropertiesasync.md) method to get the URL of the presentation file.


```js
function getFileUrl() {
    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            app.showNotification("The file hasn't been saved yet. Save the file and try again");
        }
        else {
            app.showNotification(fileUrl);
        }
    });
}
```


## Navigate to a particular slide in the presentation

The  `getSelectedRange` function calls the [Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) method to get a JSON object returned by `asyncResult.value`, which contains an array named "slides" that contains the ids, titles, and indexes of selected range of slides (or just the current slide). It also saves the id of the first slide in the selected range to a global variable.


```js
function getSelectedRange() {
    // Get the id, title, and index of the current slide (or selected slides) and store the first slide id */
    Globals.firstSlideId = 0;

    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            Globals.firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
}
```

The  `goToFirstSlide` function calls the [Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md) method to go to the id of the first slide stored by the `getSelectedRange` function above.




```js
function goToFirstSlide() {
    Office.context.document.goToByIdAsync(Globals.firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```


## Navigate between slides in the presentation

The  `goToSlideByIndex` function calls the **Document.goToByIdAsync** method to navigate to the next slide in the presentation.


```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```




## Additional resources

- [How to save add-in state and settings per document for content and task pane add-ins](../../docs/develop/persisting-add-in-state-and-settings.md#PersistSettingsContentTaskPaneApp)

- [Read and write data to the active selection in a document or spreadsheet](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
- [Get the whole document from an add-in for PowerPoint or Word](../../docs/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)
    
- [Use document themes in your PowerPoint add-ins](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md)
    
