---
title: Add and remove attachments in an Outlook add-in
description: Use various attachment APIs to manage the files or Outlook items attached to the item the user is composing.
ms.date: 03/12/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Manage an item's attachments in a compose form in Outlook

The Office JavaScript API provides several APIs to manage an item's attachments when the user is composing a message or appointment.

## Attach a file or Outlook item

Attach a file or Outlook item to a compose form by using the method that's appropriate for the type of attachment.

- [addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods): Attach a file.

    > [!NOTE]
    > The `addFileAttachmentAsync` method was introduced in [requirement set 1.1](/javascript/api/requirement-sets/outlook/requirement-set-1.1/outlook-requirement-set-1.1) for Outlook on Windows (classic) and on Mac. Support for `addFileAttachmentAsync` in Outlook on the web and new Outlook on Windows was introduced in [requirement set 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8).

- [addFileAttachmentFromBase64Async](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods): Attach a file using its Base64-encoded string.
- [addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods): Attach an Outlook item.

These are asynchronous methods, which means execution can go on without waiting for the action to complete. Depending on the original location and size of the attachment being added, the asynchronous call may take a while to complete.

If there are tasks that depend on the action to complete, you should carry out those tasks in a callback function. This callback function is optional and is invoked when the attachment upload has completed. The callback function takes an [AsyncResult](/javascript/api/office/office.asyncresult) object as an output parameter that provides any status, error, and returned value from adding the attachment. If the callback requires any extra parameters, you can specify them in the optional `options.asyncContext` parameter. `options.asyncContext` can be of any type that your callback function expects.

For example, you can define `options.asyncContext` as a JSON object that contains one or more key-value pairs. For more examples about passing optional parameters to asynchronous methods in the Office Add-ins platform, see [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-to-asynchronous-methods). The following example shows how to use the `asyncContext` parameter to pass two arguments to a callback function.

```javascript
const options = { asyncContext: { var1: 1, var2: 2 } };

Office.context.mailbox.item.addFileAttachmentAsync("https://contoso.com/rtm/icon.png", "icon.png", options, callback);
```

To check for the outcome of an asynchronous method call in the callback function, use the `status` and `error` properties of the `AsyncResult` object. If the attaching completes successfully, use the `AsyncResult.value` property to get the attachment ID. The attachment ID is an integer which you can subsequently use to remove the attachment.

> [!NOTE]
> The attachment ID is valid only within the same session and isn't guaranteed to map to the same attachment across sessions. Examples of when a session is over include when the user closes the add-in, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.

> [!TIP]
> There are limits to the files or Outlook items you can attach to a mail item, such as the number of attachments and their size. For further guidance, see [Limits for JavaScript API](limits-for-activation-and-javascript-api-for-outlook-add-ins.md#limits-for-javascript-api).

### Attach a file

You can attach a file to a message or appointment in a compose form by using the `addFileAttachmentAsync` method and specifying the URI of the file. You can also use the `addFileAttachmentFromBase64Async` method, specifying the Base64-encoded string as input. If the file is protected, you can include an appropriate identity or authentication token as a URI query string parameter. Exchange will make a call to the URI to get the attachment, and the web service which protects the file will need to use the token as a means of authentication.

> [!NOTE]
>
> - The URI of the file to be attached must support caching in production. The server hosting the image shouldn't return a `Cache-Control` header that specifies `no-cache`, `no-store`, or similar options in the HTTP response. However, when you're developing the add-in and making changes to files, caching can prevent you from seeing your changes. We recommend using `Cache-Control` headers during development.
>
> - The `addFileAttachmentAsync` method doesn't support bitmap (BMP) images if they're added as inline attachments.

The following JavaScript example is a compose add-in that attaches a file, **picture.png**, from a web server to the message or appointment being composed. The callback function takes `asyncResult` as a parameter, checks for the result status, and gets the attachment ID if the method succeeds.

```javascript
// Add the specified file attachment to the item
// being composed.
// When the attachment finishes uploading, the
// callback function is invoked and gets the attachment ID.
// You can optionally pass any object that you would
// access in the callback function as an argument to
// the asyncContext parameter.
Office.context.mailbox.item.addFileAttachmentAsync(
    "https://webserver/picture.png",
    "picture.png",
    { asyncContext: { var1: 1, var2: 2 } },
    (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error(asyncResult.error.message);
            return;
        }

        // Get the ID of the attached file.
        const attachmentID = asyncResult.value;
        console.log(`ID of added attachment: ${attachmentID}`);
    }
);
```

To add an inline Base64-encoded image to the body of a message or appointment being composed, use the [Body API](/javascript/api/outlook/office.body) methods, such as [prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1)), [setSignatureAsync](/javascript/api/outlook/office.body#outlook-office-body-setsignatureasync-member(1)), or [setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1)).

> [!TIP]
> Before inserting the image inline using `Office.context.mailbox.item.body.setAsync`, you must first call `Office.context.mailbox.item.body.getAsync` to get the current body of the mail item. Otherwise, the image won't render in the body once it's inserted. For guidance, see the [Add inline Base64-encoded image to message or appointment body (Compose)](https://raw.githubusercontent.com/OfficeDev/office-js-snippets/refs/heads/main/samples/outlook/20-item-body/add-inline-base64-image.yaml) sample in [Script Lab](../overview/explore-with-script-lab.md).

The following is an example of a Base64-encoded image prepended to the body of a mail item.

```javascript
const base64String =
  "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAnUExURQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAN0S+bUAAAAMdFJOUwAQIDBAUI+fr7/P7yEupu8AAAAJcEhZcwAADsMAAA7DAcdvqGQAAAF8SURBVGhD7dfLdoMwDEVR6Cspzf9/b20QYOthS5Zn0Z2kVdY6O2WULrFYLBaLxd5ur4mDZD14b8ogWS/dtxV+dmx9ysA2QUj9TQRWv5D7HyKwuIW9n0vc8tkpHP0W4BOg3wQ8wtlvA+PC1e8Ao8Ld7wFjQtHvAiNC2e8DdqHqKwCrUPc1gE1AfRVgEXBfB+gF0lcCWoH2tYBOYPpqQCNwfT3QF9i+AegJfN8CtAWhbwJagtS3AbIg9o2AJMh9M5C+SVGBvx6zAfmT0r+Bv8JMwP4kyFPir+cswF5KL3WLv14zAFBCLf56Tw9cparFX4upgaJUtPhrOS1QlY5W+vWTXrGgBFB/b72ev3/0igUdQPppP/nfowfKUUEFcP207y/yxKmgAYQ+PywoAFOfCH3A2MdCFzD3kdADBvq10AGG+pXQBgb7pdAEhvuF0AIc/VtoAK7+JciAs38KIuDugyAC/v4hiMCE/i7IwLRBsh68N2WQjMVisVgs9i5bln8LGScNcCrONQAAAABJRU5ErkJggg==";

// Add the Base64-encoded image to the beginning of the body.
Office.context.mailbox.item.addFileAttachmentFromBase64Async(base64String, "sample.png", { isInline: true }, (attachmentResult) => {
    if (attachmentResult.status === Office.AsyncResultStatus.Failed) {
      console.log(`Failed to attach file: ${attachmentResult.error.message}`);
      return;
    }

    Office.context.mailbox.item.body.prependAsync('<img src="cid:sample.png" />', { coercionType: Office.CoercionType.Html }, (prependResult) => {
      if (prependResult.status === Office.AsyncResultStatus.Failed) {
        console.log(`Failed to prepend image to body: ${attachmentResult.error.message}`);
        return;
      }

      console.log("Inline Base64-encoded image added to the beginning of the body.");
    })
});
```

### Attach an Outlook item

To attach an Outlook item (for example, email, calendar, or contact item) to a message or appointment in a compose form, specify the Exchange Web Services (EWS) ID of the item and use the `addItemAttachmentAsync` method. To get the EWS ID of the item, use the [item.itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) property.

The following JavaScript function, `addItemAttachment`, extends an earlier example, and adds an item as an attachment to the email or appointment being composed. The function takes the EWS ID of the item to be attached as an argument. If attaching succeeds, it gets the attachment ID for further processing.

```javascript
// Adds the specified item as an attachment to the composed item.
// ID is the EWS ID of the item to be attached.
function addItemAttachment(itemId) {
    // When the attachment finishes uploading, the
    // callback function is invoked. Here, the callback
    // function uses only asyncResult as a parameter,
    // and if the attaching succeeds, gets the attachment ID.
    // You can optionally pass any other object you wish to
    // access in the callback function as an argument to
    // the asyncContext parameter.
    Office.context.mailbox.item.addItemAttachmentAsync(
        itemId,
        "Welcome email",
        { asyncContext: { var1: 1, var2: 2 } },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error(asyncResult.error.message);
                return;
            }

            const attachmentID = asyncResult.value;
            console.log(`ID of added attachment: ${attachmentID}`);
        }
    );
}
```

> [!NOTE]
> You can use a compose add-in to attach an instance of a recurring appointment in Outlook on the web, on mobile devices, or in [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627). However, in a supporting Outlook client on Windows or on Mac, attempting to attach an instance would result in attaching the recurring series (the parent appointment).

## Get attachments

The following APIs to get attachments in compose mode are available from [requirement set 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8).

- [getAttachmentsAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [getAttachmentContentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)

Use the [getAttachmentsAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) method to get the attachments of the message or appointment being composed.

> [!NOTE]
> In Outlook on the web and the new Outlook on Windows, users can select the **Upload and share** option to upload an attachment to OneDrive and include a link to the file in the mail item. However, since only a link is included, `getAttachmentsAsync` doesn't return this type of attachment.

To get an attachment's content, use the [getAttachmentContentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) method. The supported formats are listed in the [AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) enum.

You should provide a callback function to check for the status and any error by using the `AsyncResult` output parameter object. You can also pass any additional parameters to the callback function by using the optional `asyncContext` parameter.

The following JavaScript example gets the attachments and allows you to set up distinct handling for each supported attachment format.

```javascript
const item = Office.context.mailbox.item;
const options = { asyncContext: { currentItem: item } };
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (let i = 0 ; i < result.value.length ; i++) {
      result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a Base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // Handle file attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      break;
    default:
      // Handle attachment formats that are not supported.
  }
}
```

> [!TIP]
> If the Outlook client in which your add-in is running doesn't support Mailbox requirement set 1.8, you can still get an attachment and its contents from an item being composed using Microsoft Graph or EWS. To learn more, see [Get an Outlook item's attachments from Exchange](get-attachments-of-an-outlook-item.md).

## Remove an attachment

To remove a file or item attachment from a message or appointment item in a compose form, specify the corresponding attachment ID in the [removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) method.

> [!IMPORTANT]
> If you're using requirement set 1.7 or earlier, you should only remove attachments that the same add-in has added in the same session.

Similar to the `addFileAttachmentAsync`, `addItemAttachmentAsync`, and `getAttachmentsAsync` methods, `removeAttachmentAsync` is an asynchronous method. You should provide a callback function to check for the status and any error by using the `AsyncResult` output parameter object. You can also pass any additional parameters to the callback function by using the optional `asyncContext` parameter.

The following JavaScript function, `removeAttachment`, continues to extend the examples above, and removes the specified attachment from the email or appointment being composed. The function takes as an argument the ID of the attachment to be removed. You can obtain the ID of an attachment after a successful `addFileAttachmentAsync`, `addFileAttachmentFromBase64Async`, or `addItemAttachmentAsync` method call, and use it in a subsequent `removeAttachmentAsync` method call. You can also call `getAttachmentsAsync` (introduced in requirement set 1.8) to get the attachments and their IDs for that add-in session.

```js
// Removes the specified attachment from the composed item.
function removeAttachment(attachmentId) {
    // When the attachment is removed, the callback function is invoked.
    // Here, the callback function uses an asyncResult parameter and
    // gets the ID of the removed attachment if the removal succeeds.
    // You can optionally pass any object you wish to access in the
    // callback function as an argument to the asyncContext parameter.
    Office.context.mailbox.item.removeAttachmentAsync(
        attachmentId,
        { asyncContext: null },
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error(asyncResult.error.message);
                return;
            }

            console.log(`Removed attachment with the ID: ${asyncResult.value}`);
        }
    );
}
```

> [!TIP]
> The `removeAttachmentAsync` method doesn't remove inline attachments from a mail item. To remove an inline attachment, first get the item's body, then remove any references of the attachment from its contents. Use the [Office.Body](/javascript/api/outlook/office.body) APIs to get and set the body of an item.

## See also

- [Get an Outlook item's attachments from Exchange](get-attachments-of-an-outlook-item.md)
- [Create Outlook add-ins for compose forms](compose-scenario.md)
- [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md)
- [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
