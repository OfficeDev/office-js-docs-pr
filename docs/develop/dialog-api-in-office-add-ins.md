# Opening dialogs from Office Add-ins 

Get some guidance for using the JavaScript Dialog API in Office Add-ins.

>**Note:** The Dialog API is currently available only in
>  
>  - Office for Windows Desktop 2016 (build 16.0.6741.0000 or later)
>  - Office for IPad (build 1.22 or later)
>  - Office for Mac (build 15.20 or later) 
>
>Support is coming soon to online Office services. It can be used in Excel, Word, PowerPoint, and Outlook.

There are several scenarios in which you need to open a dialog box from a task pane, or content Office Add-in, or an [add-in command](https://dev.office.com/docs/add-ins/design/add-in-commands). Some examples: 

- To display sign-in pages that cannot be opened directly in a task pane.
- To provide more screen space, or even a full screen, for some tasks in your add-in.
- To host a video that would be too small if confined to a task pane.

>**Note:** Overlapping UI can quickly annoy users, so try to avoid opening a dialog from a task pane, unless your scenario really requires it. In this connection, note that task panes can be tabbed. For an example, see the sample [Excel Add-in JavaScriptSalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).

Here's an example of what such a dialog box looks like. 

![Add-in commands](../../images/Auth0DialogOpen.PNG)

Note that the dialog box always opens in the center of the screen. It can be moved and resized by the user. The window is *nonmodal*: a user can continue to interact with the both the document in the host Office application and with the host page in the task pane, if there is one.

## The basics

The Office JavaScript APIs support these scenarios with a [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog) object and two functions in the [Office.context.ui namespace](https://dev.office.com/reference/add-ins/shared/officeui), all of which are discussed in detail below. 

### Opening a dialog box

To open a dialog, your code in the task pane calls the [displayDialogAsync](http://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) method and passes to it the URL of page that should open. The following is the simplest possible example.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html'}); 
```

Note the following about this code:

- The URL uses the http**s** protocol. This is mandatory for all pages loaded in a dialog box, not just the first page loaded.
- The domain is the same as the domain of the add-in; that is, of the task pane, or content add-in, or the [function file](https://dev.office.com/reference/add-ins/manifest/functionfile) of an add-in command. This is not mandatory for the first page loaded in the dialog box, but if the first page is not in the same domain, then it's domain must be listed in the `<AppDomains>` section of the add-in manifest.

After the first page is loaded, a user can navigate to any website that uses https. You can also design the first page to immediately redirect to another site. 

By default the dialog box will occupy 80% of the height and width of the device screen, but you can set different percentages by passing a configuration object to the method as in the following example.

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20}); 
```

For a sample add-in that does this, see [Office Add-in Dialog API Simple](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)

>**Note:** Only one dialog box can be open from a host window. Attempting to open another generates an error. (See [Errors from displayDialogAsync](#Errors-from-displayDialogAsync) for more information.) So, for example, if a user has opened a dialog from a task pane, she cannot open a second dialog, not even from a different page in the task pane. However, when a dialog is opened from an [add-in command](https://dev.office.com/docs/add-ins/design/add-in-commands), the command opens a new (but unseen) HTML file each time it is selected. This creates a new (unseen) host window, so each such window can launch its own dialog box. 

### Sending information from the dialog to the host page

The dialog box cannot talk back to the host page in the task pane unless:

- The current page in the dialog box is in the host page's domain
- The Office JavaScript library is loaded in the page. (Like any page that uses the Office library, script for the page must assign a method to the `Office.initialize` property, although it can be an empty method. For details see [Initializing your add-in](http://dev.office.com/docs/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in).) 

Code in the dialog page uses the `messageParent` function to send either a boolean value or a string message to the host page. The string can be a word, sentence, XML blob, stringified JSON, or anything else that can be serialized to a string. The following is an example.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true); 
}
```

>**Note:** 
>
> - The `messageParent` function is the *only* Office API that can be called in the dialog window.
> - The `messageParent` function can only be called on a page with the same domain (including protocol and port) as the host page.

In the next example, `googleProfile` is a stringified version of the user's Google profile.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile); 
}
```

The host page must be configured to receive the message. You do this by adding a callback parameter to the original call of `displayDialogAsync`. The callback assigns a handler to the `DialogMessageReceived` event. The following is an example. Note the following about this code:

- Office passes an [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object to the callback. It represents the result of the attempt to open the dialog window. It does not represent the outcome of any events in the dialog window. For more on this distinction, see the section [Handling errors and events](#handling-errors-and-events). 
- The `value` property of the `asyncResult` is set to a [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog) object, which exists in the host page, not in the dialog window's execution context.
- The `processMessage` is the function that handles the event. There are examples below. You can give it any name you want.
- The `dialog` variable is declared at a wider scope than the callback because it is also referenced in `processMessage`.

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20},
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
); 
```

The following code is a very simple example of a handler for the `DialogMessageReceived` event. Note the following about this code.

- Office passes the `arg` object to the handler. It's `message` property is the boolean or string sent by the call of `messageParent` in the dialog. In this example, it is a stringified representation of a user's profile from a service such as Microsoft Account or Google, so it is deserialized back to an object with `JSON.parse`.
- The `showUserName` implementation is not shown. It might display a personalized welcome message on the task pane.

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

When there will be no more user interaction with the dialog box, your message handler should close the dialog as shown in this example. Note the following about this code:

- The `dialog` object must be the same one that is returned by the call of `displayDialogAsync`. 
- The call of `dialog.close` tells Office to immediately close the dialog. Typically, you call it as the first line of the handler.

```js
function processMessage(arg) {
    dialog.close();
	// message processing code goes here;
}
```

For a sample add-in that uses these techniques, see [Office Add-in Dialog API Simple](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

If the add-in needs to navigate the task pane to a different page immediately after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example.

```js
function processMessage(arg) {
    // message processing code goes here;
	window.location.replace("/newPage.html");
	// Alternatively ...
	// window.location.href = "/newPage.html";
}
```

For an example of an add-in which does this, see [PowerPoint-Add-in-Microsoft-Graph-ASP NET-InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).. 

#### Conditional messaging

There can be multiple calls of `messageParent` in the dialog, but there can be only one handler in the host page for the `DialogMessageReceived` event, so when different messages can be sent, the handler has to use conditional logic to distinguish them. For example, if the dialog box is used to sign in to an identity provider such as Microsoft Account or Google, the dialog sends the user's profile as a message, but if the sign-in fails, the dialog should send error information to the host page, as in the following example. Note the following about this code:

- The `loginSuccess` variable would be initialized by reading the HTTP response from the identity provider.
- The the implementation of the `getProfile` and `getError` functions are not not shown. They each get data from a query parameter or from the body of the HTTP response.
- Anonymous objects of different types are sent depending on whether the sign in was successful. Both have an `messageType` property, but one has a `profile` property and the other has a `error` property.

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

For a sample that uses conditional messaging, see [Office-Add-in-Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0).

The handler code in the host page uses the value of the `messageType` property to branch as in the following example. Note that the `showUserName` function is the same as in the example above and `showNotification` function displays the error in the host page's UI. 

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

### Closing the dialog

You can implement a button in the dialog that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example:

```js
function closeButtonClick() {
	var messageObject = {messageType: "dialogClosed"};            
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage); 
}
``` 

The host page handler for `DialogMessageReceived` would just call `dialog.close`, as in this example. (See examples above for how the dialog object is initialized.)

```js
function processMessage(arg) {
	var messageFromDialog = JSON.parse(arg.message);
	if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
	}
}
```

Even when you don't have your own close dialog UI, an end user can close the dialog by clicking the **X** in the upper right corner. This action triggers the `DialogEventReceived` event. If your host pane needs to know when this happens it should declare a handler for this event. See the section [Errors and events in the dialog window](#errors-and-events-in-the-dialog-window) for details.

## Handling errors and events 

There are two categories of errors that your code should handle:

- Errors returned by the call of `displayDialogAsync` because the dialog window cannot be created. 
- Errors, and other events, in the dialog window.

### Errors from displayDialogAsync

In addition to general platform and system errors, there are three errors specific to calling `displayDialogAsync`:

|Code number|Meaning|
|:-----|:-----|
|12004|The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be either the same domain as the host page (including protocol and port number) or it must be registered in the `<AppDomains>` section of the add-in manifest.|
|12005|The URL passed to `displayDialogAsync` uses the HTTP protocol. HTTPS is required. (In some versions of Office, the error message returned with 12005 is the same one returned for 12004.)|
|12007|There is already a dialog opened from this host window. A host window, such as a task pane, can only have one open at a time.|

When `displayDialogAsync` is called it always passes an [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object to its callback function. When the call is successful; that is, the dialog window is opened; the `value` property of the `AsyncResult` object is a [Dialog](https://dev.office.com/reference/add-ins/shared/officeui.dialog) object. There is an example of this in the section [Sending information from the dialog to the host page](#sending-information-from-the-dialog-to-the-host-page). When the call to `displayDialogAsync` fails, the window is not created, the `status` property of the `AsyncResult` object is set to "failed" and the `error` property of the object is populated. You should always have a callback that tests the `status` and responds when it's an error. The following is an example that simply reports the error message regardless of its code number. 

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', 
function (asyncResult) {
    if (asyncResult.status === "failed") { 
        showNotification(asynceResult.error.code = ": " + asyncResult.error.message); 
    } else {
	    dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
}); 
```

### Errors and events in the dialog window

There are three errors and events, known by their code numbers, in the dialog that will trigger a `DialogEventReceived` event in the host page. 

|Code number|Meaning|
|:-----|:-----|
|12002|*Either*, there is no page at the URL that was passed to `displayDialogAsync`, *or* the page that was passed to `displayDialogAsync` did load, but now the dialog has been directed to a page that it cannot find or load, or it has been directed to a URL with invalid syntax.|
|12003|The dialog has been directed to a URL with the HTTP protocol. HTTPS is required.|
|12006|The dialog has been closed, usually because the user click the **X** button.|

Your code can assign a handler for the `DialogEventReceived` event in the call to `displayDialogAsync`. The following is a simple example.

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', 
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
); 
```

The following is an example of a handler for the `DialogEventReceived` event which creates custom error messages for each error code. 

```js
function processDialogEvent(arg) {
    switch (arg.error) {
        case 12002:
            showNotification("The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.");
            break;
        case 12003:
            showNotification("The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.");
            break;
        case 12006:
            showNotification("Dialog closed");
            break;
		default:
			showNotification("Unknown error in dialog box.");
            break;
    }
}
```
  
## Passing information to the dialog

Sometimes the host page needs to pass information to the dialog window. There are two main ways to do this:

- Add query parameters to the URL that is passed to `displayDialogAsync`. 
- Store the information somewhere that is accessible to both the host window and dialog window. The two windows do not share a common session storage; but *if they have the same domain* (including port number, if any), then they share a common [local storage](http://www.w3schools.com/html/html5_webstorage.asp).

### Using local storage

To use local storage, your code calls the `setItem` method of the `window.localStorage` object in the host page before the call of `displayDialogAsync`, as in the following example:

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

Code in the dialog window, reads the item when its needed, as in the following example:

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

For a sample add-in that uses local storage in this way, see [Office-Add-in-Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0).

### Using query parameters

Here is an example of passing data with a query parameter.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248'}); 
```

For a sample that uses this technique, see [PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

Code in your dialog window can parse URL and read the parameter value.

>**Note:** Office automatically attaches a query parameter called `_host_info` onto the URL for the host page and also to the URL that is passed to `displayDialogAsync`. (It is appended after your custom query parameters, in any.) It provides information about the host Office application. The following is an example.

>```
>_host_Info=Word|Win32|16.01|en-US|telemetry|isDialog
>```
>
>The same value is also added to [session storage](http://www.w3schools.com/html/html5_webstorage.asp) with the key `hostInfoValue`. Code in your dialog window can parse and use this information if it is needed. This information is used by Office internally, **so do not change the `hostInfoValue` value in the window.sessionStorage**. The six parts are defined as follows:

> 1. **Host application**, such as Word, Excel, PowerPoint, etc.
> 2. **Platform**, which can be `Win32`, `Mac`, `iOS`, `Web` (Office Online), or `Winrt` (Windows Immersive).
> 3. **Version**
> 4. **Locale**
> 5. **Telemetry ID** is a GUID when the platform is `Web`. On all other platforms it is the unused placeholder "telemetry". It is not present on the host page URL.
> 6. **isDialog** is self-explanatory and is not present on the host page URL. 

## Using the Dialog APIs to show a video

To show a video in a dialog box take these steps:

1.  Create a page whose only content is an iFrame. The `src` attribute of the iFrame points to an online video. The protocol of the video's URL must be HTTP**S**. We'll call this page "video.dialogbox.html". The following is an example of the markup:

		<iframe class="ms-firstrun-video__player"  width="640" height="360" 
			src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1" 
			frameborder="0" allowfullscreen>
		</iframe>

2.  The video.dialogbox.html page must be either in the same domain as the host page or in a domain that is registered in the `<AppDomains>` section of the add-in manifest.
3.  Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.
4.  If your add-in needs to know when the user has closed the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 error. For details, see the section [Errors and events in the dialog window](#errors-and-events-in-the-dialog-window).

For a sample that shows a video with the Dialog APIs, see the [video placemat design pattern](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat) in the [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) repo.

![Add-in commands](../../images/VideoPlacematDialogOpen.PNG)

## Using the Dialog APIs in an authentication flow

A primary scenario for the Dialog APIs is to enable authentication with a resource or identity provider that does not allow its sign-in page to open in an iFrame, such as Microsoft Account, Office 365, Google, and Facebook. The following is a simple and typical flow of steps:

1. The user selects a UI element on the host page to sign-in. The handler for the element calls `displayDialogAsync` and passes the URL of an identity provider's sign-in page. *Since this is the first page opened in the dialog and it does not have the same domain as the host window, its domain must be listed in the `<AppDomains>` section of the add-in manifest.* The URL includes a query parameter that tells the identity provider where to redirect the dialog window after sign-in. We'll call this page "redirectPage.html". (*This must be a page in the same domain as the host window*, because the only way for the dialog window to pass the results of the sign-in attempt is with a call of `messageParent` and it can only be called on a page with the same domain as the host window.) 
2. The identity provider's service processes the incoming GET request from the dialog window. If the user is already logged in, it immediately redirects the window to redirectPage.html and includes user data as a query parameter. If the user is not already signed in, the provider's sign-in page appears in the window, and the user signs in. For most providers, if the user cannot sign-in successfully, the provider shows an error page in the dialog window and does not redirect to redirectPage.html. The user must close the window by selecting the **X** in the corner. If the user successfully signs in, the dialog window is redirected to redirectPage.html and user data is included as a query parameter.
3. When the redirectPage.html page opens, it calls `messageParent` to report the success or failure to the host page and optionally also report user data or error data. 
4. The `DialogMessageReceived` event fires in the host page and its handler closes the dialog window and optionally does other processing of the message. 

For a sample add-in that uses the above pattern, see [Excel Add-in ASP.NET QuickBooks](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)

### Alternate auth scenarios

#### Dealing with a slow network

If the network or the identity provider is especially slow, the dialog window may not open right away after the user selects the UI to open it. This could give the impression that nothing is happening. One way to ensure a better experience is to have the first page that opens in the dialog window be a local page hosted in the add-in's domain; that is, the host window's domain. This page could have a simple UI that says statement like "Please wait, we are redirecting you to sign-in page of *NAME-OF-PROVIDER*." 

Code in this page constructs the URL of the identity provider's sign-in page by using information that is passed to the dialog using one of the techniques in [Passing info to the dialog](#passing-info-to-the-dialog). It then redirects to the sign-in page. In this design, the provider's page is not the first page opened in the dialog window, so it is not necessary to list the provider's domain in the `<AppDomains>` section of the add-in manifest

For sample add-ins that use this pattern, see:

- [PowerPoint-Add-in-Microsoft-Graph-ASP NET-InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Office Add-in Office 365 Client Authentication for AngularJS](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth).

#### Supporting multiple identity providers

If your add-in gives the user has a choice of providers, such as Microsoft Account, Google, and Facebook, then you need a local first page (see preceding section) that has a UI where the user selects one. Selection triggers the construction of the sign-in URL and redirection to it. 

For a sample that uses this pattern, see [Office-Add-in-Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0).

#### Authorization of the add-in to an external resource

In the modern web, web applications are security principals just as users are, and the application has its own identity and permissions to an online resource such as Office 365, Google Plus, Facebook, LinkedIn, etc. The application is registered with the resource provider before it is deployed and the registration includes: 

- A listing of the permissions that the application needs to a user's resources.
- A URL to which the resource service should return an access token when the application accesses the service.  

When a user invokes a function in the application that accesses the user's data in the resource service, the user is prompted to sign-in to the service and is then prompted to grant the application the permissions it needs to the user's resources. The service then redirects the sign-in window to the previously registered URL and passes the access token. The application uses the access token to access the user's resources. 

The Office Dialog APIs can be used to manage this process by using a flow that is nearly identical to the one described above for user sign-in, or to the variation described in [Dealing with a slow network](#dealing-with-a-slow-network). The only differences are:

- If the user hasn't previously granted the application the permissions it needs, she is prompted to do so in the dialog window immediately after signing in. 
- The dialog window sends the access token to the host window either by using `messageParent` to send the stringified access token or by storing the access token where the host window can retrieve it. The token has a time limit, but while it lasts, the host window can use it to directly access the user's resources without any further prompting of the user.

We have two samples that use the Dialog APIs for this purpose:

- Stores the access token in a database: [PowerPoint-Add-in-Microsoft-Graph-ASP NET-InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).
- Passes the access token in `messageParent`: [Excel Add-in ASP.NET QuickBooks](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)

#### More information about authentication and authorization in add-ins

For more information about authentication and authorization in Office add-ins, see [Authorize external services in your Office Add-in](https://dev.office.com/docs/add-ins/develop/auth-external-add-ins) and see also the library [office-js-helpers](https://github.com/OfficeDev/office-js-helpers). 


## Using the Office Dialog API with single-page applications and client-side routing

If your add-in uses client-side routing, as single-page applications typically do, you have the option of passing  the URL of a route to the [displayDialogAsync](http://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) method, instead of the URL of a complete and separate HTML page. 

It is important to remember, if you pass a route, that the dialog box is in an entirely new window with it's own execution context. Your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog window. So this technique launches a second instance of your single page application in the dialog window. Code that changes variables in the dialog window does not change the task pane version of the same variables. Similarly, the dialog window has its own session storage, which is not accessible from code in the task pane.  

