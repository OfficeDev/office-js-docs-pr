---
title: PowerPoint add-ins
description: ''
ms.date: 10/16/2018
---

# PowerPoint add-ins

You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iOS, Office Online, and Mac. You can create two types of PowerPoint add-ins:

- Use **content add-ins** to add dynamic HTML5 content to your presentations. For example, see the [LucidChart Diagrams for PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) add-in, which you can use to inject an interactive diagram from LucidChart into your deck.

- Use **task pane add-ins** to bring in reference information or insert data into the slide via a service. For example, see the [Shutterstock Images](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false) add-in, which you can use to add professional photos to your presentation. 

## PowerPoint add-in scenarios

The code examples in the article show you some basic tasks for developing content add-ins for PowerPoint. 

To display information, these examples depend on the `app.showNotification` function, which is included in the Visual Studio Office Add-ins project templates. If you aren't using Visual Studio to develop your add-in, you'll need replace the `showNotification` function with your own code. Several of these examples also depend on this `globals` object that is declared outside of the scope of these functions: `var globals = {activeViewHandler:0, firstSlideId:0};`

These code examples require your project to [reference Office.js v1.1 library or later](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).


## Detect the presentation's active view and handle the ActiveViewChanged event

If you are building a content add-in, you will need to get the presentation's active view and handle the `ActiveViewChanged` event, as part of your `Office.Initialize` handler.


- The  `getActiveFileView` function calls the [Document.getActiveViewAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getactiveviewasync-options--callback-) method to return whether the presentation's current view is "edit" (any of the views in which you can edit slides, such as **Normal** or **Outline View**) or "read" ( **Slide Show** or **Reading View**) view.


- The  `registerActiveViewChanged` function calls the [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) method to register a handler for the [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) event. 

> [!NOTE]
> In PowerPoint Online, the [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) event will never fire as Slide Show mode is treated as a new session. In this case, the add-in must fetch the active view on load, as shown in the following code sample.

```js
//general Office.initialize function. Fires on load of the add-in.
Office.initialize = function(){

    //Gets whether the current view is edit or read.
    var currentView = getActiveFileView();

    //register for the active view changed handler
    registerActiveViewChanged();

    //render the content based off of the currentView
    //....
}

function getActiveFileView()
{
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification(asyncResult.value);
        }
    });

}

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
    

## Navigate to a particular slide in the presentation

The  `getSelectedRange` function calls the [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) method to get a JSON object returned by `asyncResult.value`, which contains an array named "slides" that contains the ids, titles, and indexes of selected range of slides (or just the current slide). It also saves the id of the first slide in the selected range to a global variable.


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

The  `goToFirstSlide` function calls the [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-) method to go to the id of the first slide stored by the `getSelectedRange` function above.




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

## Get the URL of the presentation

The  `getFileUrl` function calls the [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-) method to get the URL of the presentation file.


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



## See also
- [PowerPoint Code Samples](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,PowerPoint)
- [How to save add-in state and settings per document for content and task pane add-ins](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [Read and write data to the active selection in a document or spreadsheet](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [Get the whole document from an add-in for PowerPoint or Word](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [Use document themes in your PowerPoint add-ins](use-document-themes-in-your-powerpoint-add-ins.md)
    
