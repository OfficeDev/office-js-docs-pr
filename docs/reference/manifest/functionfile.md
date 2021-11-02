---
title: FunctionFile element in the manifest file
description: Specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI.
ms.date: 09/29/2021
ms.localizationpriority: medium
---

# FunctionFile element

Specifies the source code file for operations that an add-in exposes in one of the following ways.

* Add-in commands that execute a JavaScript function instead of displaying UI.
* Keyboard shortcuts that execute a JavaScript function.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

The `FunctionFile` element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md). The `resid` attribute of the `FunctionFile` element can be no more than 32 characters and is set to the value of the `id` attribute of a `Url` element in the `Resources` element that contains the URL to an HTML file that contains or loads all the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).

> [!NOTE]
> When the add-in is configured to use a [shared runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md), the functions in the code file run in the same JavaScript runtime (and share a common global namespace) as the JavaScript in the add-in's task pane (if any).
>
> The `FunctionFile` element and the associated code file also have a special role to play with [custom keyboard shortcuts](../../design/keyboard-shortcuts.md), which require a shared runtime.

The following is an example of the `FunctionFile` element.

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

The JavaScript in the HTML file indicated by the `FunctionFile` element must call `Office.initialize` and define named functions that take a single parameter: `event`. The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user. It should also call `event.completed` when it has finished execution. The name of the functions are used in the `FunctionName` element for UI-less buttons.

The following is an example of an HTML file defining a `trackMessage` function.

```js
Office.initialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

The following code shows how to implement the function used by `FunctionName`.

```js
// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

// Your function must be in the global namespace.
function writeText(event) {

    // Implement your custom code here. The following code is a simple example.

    Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed") {
                // Show error message.
            }
            else {
                // Show success message.
            }
        });
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}
```

> [!IMPORTANT]
> The call to `event.completed` signals that you have successfully handled the event. When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued. The first event runs automatically, while the other events remain on the queue. When your function calls `event.completed`, the next queued call to that function runs. You must call `event.completed`; otherwise your function will not run.
