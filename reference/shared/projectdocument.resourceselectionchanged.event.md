

# ProjectDocument.ResourceSelectionChanged event
Occurs when the resource selection changes in the active project.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Added in**|1.0|

```js
Office.EventType.ResourceSelectionChanged
```


## Remarks

 **ResourceSelectionChanged** is an [EventType](../../reference/shared/eventtype-enumeration.md) enumeration constant that can be used in the [ProjectDocument.addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) and [ProjectDocument.removeHandlerAsync](../../reference/shared/projectdocument.removehandlerasync.md) methods to add or remove a handler for the event.


## Example

The following code example adds a handler for the  **ResourceSelectionChanged** event. When the resource selection changes in the document, it gets the GUID of the selected resource.

The example assumes your add-in has a reference to the jQuery library and that the following page control is defined in the content div in the page body.




```HTML
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            Office.context.document.addHandlerAsync(
                Office.EventType.ResourceSelectionChanged,
                getResourceGuid);
        });
    };

    // Get the GUID of the selected resource and display it in the add-in.
    function getResourceGuid() {
        Office.context.document.getSelectedResourceAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    $('#message').html(result.value);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```

For a complete code sample that shows how to use a  **ResourceSelectionChanged** event handler in a Project add-in, see [Create your first task pane add-in for Project 2013 by using a text editor](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## Support details


A capital Y in the following matrix indicates that this event is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this event.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office for Windows desktop**|**Office Online (in browser)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**Available in requirement sets**|Selection|
|**Minimum permission level**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history

|**Version**|**Changes**|
|:-----|:-----|
|1.0|Introduced|

## See also



#### Other resources


[Create your first task pane add-in for Project 2013 by using a text editor](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
[EventType enumeration](../../reference/shared/eventtype-enumeration.md)
[ProjectDocument.addHandlerAsync method](../../reference/shared/projectdocument.addhandlerasync.md)
[ProjectDocument.removeHandlerAsync method](../../reference/shared/projectdocument.removehandlerasync.md)
[ProjectDocument object](../../reference/shared/projectdocument.projectdocument.md)
