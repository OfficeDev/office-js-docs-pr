---
title: Use the Dialog API in your Office Add-ins
description: ''
ms.date: 11/28/2018
---

# Use the Dialog API in your Office Add-ins

You can use the [Dialog API](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) to open dialog boxes in your Office Add-in. This article provides guidance for using the Dialog API in your Office Add-in.

> [!NOTE]
> For information about where the Dialog API is currently supported, see [Dialog API requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets?view=office-js). The Dialog API is currently supported for Word, Excel, PowerPoint, and Outlook.

> A primary scenario for the Dialog APIs is to enable authentication with a resource such as Google or Facebook.

Consider opening a dialog box from a task pane or content add-in or [add-in command](../design/add-in-commands.md) to do the following:

- Display sign in pages that cannot be opened directly in a task pane.
- Provide more screen space, or even a full screen, for some tasks in your add-in.
- Host a video that would be too small if confined to a task pane.

> [!NOTE]
> Because overlapping UI elements are discouraged, avoid opening a dialog from a task pane unless your scenario requires it. When you consider how to use the surface area of a task pane, note that task panes can be tabbed. For an example, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.

The following image shows an example of a dialog box.

![Add-in commands](../images/auth-o-dialog-open.png)

Note that the dialog box always opens in the center of the screen. The user can move and resize it. The window is *nonmodal*--a user can continue to interact with both the document in the host Office application and with the host page in the task pane, if there is one.

## Dialog API scenarios

The Office JavaScript APIs support the following scenarios with a [Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js) object and two functions in the [Office.context.ui namespace](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js).

### Open a dialog box

To open a dialog box, your code in the task pane calls the [displayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) method and passes to it the URL of the resource that you want to open. This is usually a page, but it can be a controller method in an MVC application, a route, a web service method, or any other resource. In this article, 'page' or 'website' refers to the resource in the dialog. The following code is a simple example:

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - The URL uses the HTTP**S** protocol. This is mandatory for all pages loaded in a dialog box, not just the first page loaded.
> - The dialog resource's domain is the same as the domain of the host page, which can be the page in a task pane or the [function file](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/functionfile?view=office-js) of an add-in command. This is required: the page, controller method, or other resource that is passed to the `displayDialogAsync` method must be in the same domain as the host page.

> [!IMPORTANT]
> The host page and the resources of the dialog must have the same full domain. If you attempt to pass `displayDialogAsync` a subdomain of the add-in's domain, it will not work. The full domain, including any subdomain, must match.

After the first page (or other resource) is loaded, a user can go to any website (or other resource) that uses HTTPS. You can also design the first page to immediately redirect to another site.

By default, the dialog box will occupy 80% of the height and width of the device screen, but you can set different percentages by passing a configuration object to the method, as shown in the following example:

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

For a sample add-in that does this, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)

> [!NOTE]
> You can open only one dialog box from a host window. An attempt to open another dialog box generates an error. For example, if a user opens a dialog box from a task pane, she cannot open a second dialog box, from a different page in the task pane. However, when a dialog box is opened from an [add-in command](../design/add-in-commands.md), the command opens a new (but unseen) HTML file each time it is selected. This creates a new (unseen) host window, so each such window can launch its own dialog box. For more information, see [Errors from displayDialogAsync](#errors-from-displaydialogasync).

### Take advantage of a performance option in Office Online

The `displayInIframe` property is an additional property in the configuration object that you can pass to `displayDialogAsync`. When this property is set to `true`, and the add-in is running in a document opened in Office Online, the dialog box will open as a floating iframe rather than an independent window, which makes it open faster. The following is an example:

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

The default value is `false`, which is the same as omitting the property entirely. If the add-in is not running in Office Online, the `displayInIframe` is ignored.

> [!NOTE]
> You should **not** use `displayInIframe: true` if the dialog will at any point redirect to a page that cannot be opened in an iframe. For example, the sign in pages of many popular web services, such as Google and Microsoft Account, cannot be opened in an iframe.

### Handling pop-up blockers with Office Online

Attempting to display a dialog while using Office Online may cause the browser's pop-up blocker to block the dialog. The browser's pop-up blocker can be circumvented if the user of your add-in first agrees to a prompt from the add-in. `displayDialogAsync`'s [DialogOptions](/javascript/api/office/office.dialogoptions) has the `promptBeforeOpen` property to trigger such a pop-up. `promptBeforeOpen` is a boolean value which provides the following behavior:
 
 - `true` - The framework displays a pop-up to trigger the navigation and avoid the browser's pop-up blocker. 
 - `false` - The dialog will not be shown and the developer will handle pop-ups (usually by providing a user artifact to trigger the navigation). 
 
The pop-up looks similiar to that in the following screenshot:

![The prompt an add-in's dialog can generate to avoid in-browser pop-up blockers.](../images/dialog-prompt-before-open.png)
 
### Send information from the dialog box to the host page

The dialog box cannot communicate with the host page in the task pane unless:

- The current page in the dialog box is in the same domain as the host page.
- The Office JavaScript library is loaded in the page. (Like any page that uses the Office JavaScript library, script for the page must assign a method to the `Office.initialize` property, although it can be an empty method. For details, see [Initializing your add-in](understanding-the-javascript-api-for-office.md#initializing-your-add-in).)

Code in the dialog page uses the `messageParent` function to send either a Boolean value or a string message to the host page. The string can be a word, sentence, XML blob, stringified JSON, or anything else that can be serialized to a string. The following is an example:

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!NOTE]
> - The `messageParent` function is one of *only* two Office APIs that can be called in the dialog box. The other is `Office.context.requirements.isSetSupported`. For information about it, see [Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md).
> - The `messageParent` function can only be called on a page with the same domain (including protocol and port) as the host page.

In the next example, `googleProfile` is a stringified version of the user's Google profile.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

The host page must be configured to receive the message. You do this by adding a callback parameter to the original call of `displayDialogAsync`. The callback assigns a handler to the `DialogMessageReceived` event. The following is an example:

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20},
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);
```

> [!NOTE]
> - Office passes an [AsyncResult]() object to the callback. It represents the result of the attempt to open the dialog box. It does not represent the outcome of any events in the dialog box. For more on this distinction, see the section [Handle errors and events](#handle-errors-and-events).
> - The `value` property of the `asyncResult` is set to a [Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js) object, which exists in the host page, not in the dialog box's execution context.
> - The `processMessage` is the function that handles the event. You can give it any name you want.
> - The `dialog` variable is declared at a wider scope than the callback because it is also referenced in `processMessage`.

The following is a simple example of a handler for the `DialogMessageReceived` event:

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
> - Office passes the `arg` object to the handler. Its `message` property is the Boolean or string sent by the call of `messageParent` in the dialog. In this example, it is a stringified representation of a user's profile from a service such as Microsoft Account or Google, so it is deserialized back to an object with `JSON.parse`.
> - The `showUserName` implementation is not shown. It might display a personalized welcome message on the task pane.

When the user interaction with the dialog box is completed, your message handler should close the dialog box, as shown in this example.

```js
function processMessage(arg) {
    dialog.close();
	// message processing code goes here;
}
```

> [!NOTE]
> - The `dialog` object must be the same one that is returned by the call of `displayDialogAsync`.
> - The call of `dialog.close` tells Office to immediately close the dialog box.

For a sample add-in that uses these techniques, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example:

```js
function processMessage(arg) {
    // message processing code goes here;
	window.location.replace("/newPage.html");
	// Alternatively ...
	// window.location.href = "/newPage.html";
}
```

For an example of an add-in that does this, see the [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) sample.

#### Conditional messaging
Because you can send multiple `messageParent` calls from the dialog box, but you have only one handler in the host page for the `DialogMessageReceived` event, the handler must use conditional logic to distinguish different messages. For example, if the dialog box prompts a user to sign in to an identity provider such as Microsoft Account or Google, it sends the user's profile as a message. If authentication fails, the dialog box sends error information to the host page, as in the following example:

```js
if (loginSuccess) {
    var userProfile = getProfile();
    var messageObject = {messageType: "signinSuccess", profile: userProfile};            
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
} else {
    var errorDetails = getError();
    var messageObject = {messageType: "signinFailure", error: errorDetails};            
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

> [!NOTE]
> - The `loginSuccess` variable would be initialized by reading the HTTP response from the identity provider.
> - The the implementation of the `getProfile` and `getError` functions are not not shown. They each get data from a query parameter or from the body of the HTTP response.
> - Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.

For samples that use conditional messaging, see:
- [Office Add-in that uses the Auth0 Service to Simplify Social Login](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Office Add-in that uses the OAuth.io Service to Simplify Access to Popular Online Services](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "signinSuccess") {
        dialog.close();
        showUserName(messageFromDialog.profile.name);
        window.location.replace("/newPage.html");
    } else {
        dialog.close();
        showNotification("Unable to authenticate user: " + messageFromDialog.error);
    }
}
```

### Closing the dialog box

You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example:

```js
function closeButtonClick() {
	var messageObject = {messageType: "dialogClosed"};            
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

The host page handler for `DialogMessageReceived` would call `dialog.close`, as in this example. (See previous examples that show how the dialog object is initialized.)


```js
function processMessage(arg) {
	var messageFromDialog = JSON.parse(arg.message);
	if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
	}
}
```

For a sample that uses this technique, see the [dialog navigation design pattern](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation) in the [UX design patterns for Office Add-ins](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) repo.

Even when you don't have your own close dialog UI, an end user can close the dialog box by choosing the **X** in the upper-right corner. This action triggers the `DialogEventReceived` event. If your host pane needs to know when this happens, it should declare a handler for this event. See the section [Errors and events in the dialog window](#errors-and-events-in-the-dialog-window) for details.

## Handle errors and events

Your code should handle two categories of events:

- Errors returned by the call of `displayDialogAsync` because the dialog box cannot be created.
- Errors, and other events, in the dialog window.

### Errors from displayDialogAsync

In addition to general platform and system errors, three errors are specific to calling `displayDialogAsync`.

|Code number|Meaning|
|:-----|:-----|
|12004|The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be the same domain as the host page (including protocol and port number).|
|12005|The URL passed to `displayDialogAsync` uses the HTTP protocol. HTTPS is required. (In some versions of Office, the error message returned with 12005 is the same one returned for 12004.)|
|<span id="12007">12007</span>|A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time.|
|12009|The user chose to ignore the dialog box. This error can occur in online versions of Office, where users may choose not to allow an add-in to present a dialog.|

When `displayDialogAsync` is called, it always passes an [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js) object to its callback function. When the call is successful - that is, the dialog window is opened - the `value` property of the `AsyncResult` object is a [Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js) object. An example of this is in the section [Send information from the dialog box to the host page](#send-information-from-the-dialog-box-to-the-host-page). When the call to `displayDialogAsync` fails, the window is not created, the `status` property of the `AsyncResult` object is set to `Office.AsyncResultStatus.Failed`, and the `error` property of the object is populated. You should always have a callback that tests the `status` and responds when it's an error. For an example that simply reports the error message regardless of its code number, see the following code:

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showNotification(asyncResult.error.code = ": " + asyncResult.error.message);
    } else {
	    dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
});
```

### Errors and events in the dialog window

Three errors and events, known by their code numbers, in the dialog box will trigger a `DialogEventReceived` event in the host page.

|Code number|Meaning|
|:-----|:-----|
|12002|One of the following:<br> - No page exists at the URL that was passed to `displayDialogAsync`.<br> - The page that was passed to `displayDialogAsync` loaded, but the dialog box was directed to a page that it cannot find or load, or it has been directed to a URL with invalid syntax.|
|12003|The dialog box was directed to a URL with the HTTP protocol. HTTPS is required.|
|12006|The dialog box was closed, usually because the user chooses the **X** button.|

Your code can assign a handler for the `DialogEventReceived` event in the call to `displayDialogAsync`. The following is a simple example:

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

For an example of a handler for the `DialogEventReceived` event that creates custom error messages for each error code, see the following example:

```js
function processDialogEvent(arg) {
    switch (arg.error) {
        case 12002:
            showNotification("The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.");
            break;
        case 12003:
            showNotification("The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.");            break;
        case 12006:
            showNotification("Dialog closed.");
            break;
		default:
			showNotification("Unknown error in dialog box.");
            break;
    }
}
```

For a sample add-in that handles errors in this way, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).


## Pass information to the dialog box

Sometimes the host page needs to pass information to the dialog box. You can do this in two primary ways:

- Add query parameters to the URL that is passed to `displayDialogAsync`.
- Store the information somewhere that is accessible to both the host window and dialog box. The two windows do not share a common session storage, but *if they have the same domain* (including port number, if any),  they share a common [local storage](https://www.w3schools.com/html/html5_webstorage.asp).

### Use local storage

To use local storage, your code calls the `setItem` method of the `window.localStorage` object in the host page before the `displayDialogAsync` call, as in the following example:

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

Code in the dialog window reads the item when it's needed, as in the following example:

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

For sample add-ins that uses local storage in this way, see:

- [Office Add-in that uses the Auth0 Service to Simplify Social Login](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Office Add-in that uses the OAuth.io Service to Simplify Access to Popular Online Services](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

### Use query parameters

The following example shows how to pass data with a query parameter:

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

For a sample that uses this technique, see [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

Code in your dialog window can parse the URL and read the parameter value.

> [!NOTE]
> Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`. (It is appended after your custom query parameters, if any. It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it. The same value is added to the dialog box's session storage. Again, *your code should neither read nor write to this value*.

## Use the Dialog APIs to show a video

To show a video in a dialog box:

1.  Create a page whose only content is an iframe. The `src` attribute of the iframe points to an online video. The protocol of the video's URL must be HTTP**S**. In this article we'll call this page "video.dialogbox.html". The following is an example of the markup:

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2.  The video.dialogbox.html page must be in the same domain as the host page.
3.  Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.
4.  If your add-in needs to know when the user closes the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 event. For details, see the section [Errors and events in the dialog window](#errors-and-events-in-the-dialog-window).

For a sample that shows a video in a dialog box, see the [video placemat design pattern](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat) in the [UX design patterns for Office Add-ins](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) repo.

![Screenshot of a video showing in an add-in dialog box](../images/video-placemats-dialog-open.png)

## Use the Dialog APIs in an authentication flow

A primary scenario for the Dialog APIs is to enable authentication with a resource or identity provider that does not allow its sign-in page to open in an Iframe, such as Microsoft Account, Office 365, Google, and Facebook.

> [!NOTE]
> When you are using the Dialog APIs for this scenario, do *not* use the `displayInIframe: true` option in the call to `displayDialogAsync`. See [Take advantage of a performance option in Office Online](#take-advantage-of-a-performance-option-in-office-online) previously in this article for details about this option.

The following is a simple and typical authentication flow:

1. The first page that opens in the dialog box is a local page (or other resource) that is hosted in the add-in's domain; that is, the host window's domain. This page can have a simple UI that says "Please wait, we are redirecting you to the page where you can sign in to *NAME-OF-PROVIDER*." Code in this page constructs the URL of the identity provider's sign-in page by using information that is passed to the dialog box as described in [Pass information to the dialog box](#pass-information-to-the-dialog-box).
2. The dialog window then redirects to the sign-in page. The URL includes a query parameter that tells the identity provider to redirect the dialog window, after the user signs in, to a specific page. In this article, we'll call this page "redirectPage.html". (*This must be a page in the same domain as the host window*, because the only way for the dialog window to pass the results of the sign-in attempt is with a call of `messageParent`, which can only be called on a page with the same domain as the host window.)
2. The identity provider's service processes the incoming GET request from the dialog window. If the user is already logged on, it immediately redirects the window to redirectPage.html and includes user data as a query parameter. If the user is not already signed in, the provider's sign-in page appears in the window, and the user signs in. For most providers, if the user cannot sign in successfully, the provider shows an error page in the dialog window and does not redirect to redirectPage.html. The user must close the window by selecting the **X** in the corner. If the user successfully signs in, the dialog window is redirected to redirectPage.html and user data is included as a query parameter.
3. When the redirectPage.html page opens, it calls `messageParent` to report the success or failure to the host page and optionally also report user data or error data.
4. The `DialogMessageReceived` event fires in the host page and its handler closes the dialog window and optionally does other processing of the message.

For sample add-ins that use this pattern, see:

- [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart): The resource that is initially opened in the dialog window is a controller method that has no view of its own. It redirects to the Office 365 sign in page.
- [Office Add-in Office 365 Client Authentication for AngularJS](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth): The resource that is initially opened in the dialog window is a page.

#### Support multiple identity providers

If your add-in gives the user a choice of providers, such as Microsoft Account, Google, or Facebook, you need a local first page (see preceding section) that provides a UI for the user to select a provider. Selection triggers the construction of the sign-in URL and redirection to it.

For a sample that uses this pattern, see [Office Add-in that uses the Auth0 Service to Simplify Social Login](https://github.com/OfficeDev/Office-Add-in-Auth0).

#### Authorization of the add-in to an external resource

In the modern web, web applications are security principals just as users are, and the application has its own identity and permissions to an online resource such as Office 365, Google Plus, Facebook, or LinkedIn. The application is registered with the resource provider before it is deployed. The registration includes:

- A list of the permissions that the application needs to a user's resources.
- A URL to which the resource service should return an access token when the application accesses the service.  

When a user invokes a function in the application that accesses the user's data in the resource service, they are prompted to sign in to the service and then prompted to grant the application the permissions it needs to the user's resources. The service then redirects the sign-in window to the previously registered URL and passes the access token. The application uses the access token to access the user's resources.

You can use the Dialog APIs to manage this process by using a flow that is similar to the one described for users to sign in. The only differences are:

- If the user hasn't previously granted the application the permissions it needs, she is prompted to do so in the dialog box after signing in.
- The dialog window sends the access token to the host window either by using `messageParent` to send the stringified access token or by storing the access token where the host window can retrieve it. The token has a time limit, but while it lasts, the host window can use it to directly access the user's resources without any further prompting.

The following samples use the Dialog APIs for this purpose:
- [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) - Stores the access token in a database.
- [Office Add-in that uses the OAuth.io Service to Simplify Access to Popular Online Services](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

For more information about authentication and authorization in add-ins, see:
- [Authorize external services in your Office Add-in](auth-external-add-ins.md)
- [Office JavaScript API Helpers library](https://github.com/OfficeDev/office-js-helpers)


## Use the Office Dialog API with single-page applications and client-side routing

If your add-in uses client-side routing, as single-page applications typically do, you have the option to pass the URL of a route to the [displayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) method, instead of the URL of a complete and separate HTML page.

> [!IMPORTANT]
>The dialog box is in a new window with its own execution context. If you pass a route, your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog window. So this technique launches a second instance of your application in the dialog window. Code that changes variables in the dialog window does not change the task pane version of the same variables. Similarly, the dialog window has its own session storage, which is not accessible from code in the task pane.
