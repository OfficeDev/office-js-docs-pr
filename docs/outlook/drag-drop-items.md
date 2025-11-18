---
title: Drag and drop messages and attachments into the task pane
description: Learn how to enable drag and drop of messages and file attachments into the task pane of your Outlook add-in.
ms.date: 10/28/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Drag and drop messages and attachments into the task pane of an Outlook add-in

Drag and drop functionality allows users to seamlessly transfer messages and file attachments from their mailbox directly into your add-in's task pane. With the drag-and-drop feature, users can perform the following without leaving the Outlook client.

- Import files into a document management interface for processing or archiving.
- Upload customer records and communication logs to a customer relationship management (CRM) system for tracking.
- Convert a file into another format.

## Supported Outlook clients and surfaces

The following table outlines the Outlook clients that support the drag-and-drop feature and the APIs used to implement it.

| Outlook client | Support for drag and drop | Implementation method | Supported Outlook surfaces |
| ----- | ----- | ----- | ----- |
| Outlook on the web | Supported | Office.js API ([Office.EventType.DragAndDropEvent](/javascript/api/office/office.eventtype)) | <ul><li>Appointment Compose</li><li>Appointment Read</li><li>Message Compose</li><li>Message Read</li></ul> |
| New Outlook on Windows | Supported | Office.js API ([Office.EventType.DragAndDropEvent](/javascript/api/office/office.eventtype)) | <ul><li>Appointment Compose</li><li>Appointment Read</li><li>Message Compose</li><li>Message Read</li></ul> |
| Classic Outlook on Windows | Supported | [HTML Drag and Drop API](https://developer.mozilla.org/docs/Web/API/HTML_Drag_and_Drop_API) | <ul><li>Appointment Compose</li><li>Appointment Read</li><li>Message Compose</li><li>Message Read</li></ul> |
| Outlook on Mac | Supported | [HTML Drag and Drop API](https://developer.mozilla.org/docs/Web/API/HTML_Drag_and_Drop_API) | <ul><li>Appointment Compose</li><li>Appointment Read</li><li>Message Compose</li><li>Message Read</li></ul> |
| Outlook on iOS | Not supported | Not applicable | Not applicable |
| Outlook on Android | Not supported | Not applicable | Not applicable |

For information on which file types and scenarios are supported, see [Feature behavior and limitations](#feature-behavior-and-limitations).

## Implement the drag-and-drop feature

The drag or drop event occurs when the mouse pointer enters an add-in's task pane. Handling the drag and drop events differs depending on the Outlook client. Select the tab for your applicable client.

> [!NOTE]
> This section assumes that a task pane has already been implemented for your Outlook add-in. For information about task panes, see [Add-in commands](../design/add-in-commands.md). To create an add-in sample that already implements a task pane, follow the [Outlook quickstart](../quickstarts/outlook-quickstart-yo.md).

# [Web and Windows (new)](#tab/web)

In Outlook on the web and the new Outlook on Windows, create a handler in your JavaScript file for the [Office.EventType.DragAndDropEvent](/javascript/api/office/office.eventtype) event using the [Office.context.mailbox.addHandlerAsync](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-addhandlerasync-member(1)) method. When the `DragAndDropEvent` event occurs, the handler receives a [DragAndDropEventArgs](/javascript/api/outlook/office.draganddropeventargs) object so that you can identify when a user drags an item over the task pane, when they drop the item into the task pane, and what data is associated with the item. Depending on whether a drag or drop event occurred, the [dragAndDropEventData](/javascript/api/outlook/office.draganddropeventargs#outlook-office-draganddropeventargs-draganddropeventdata-member) property of the `DragAndDropEventArgs` object returns a [DragoverEventData](/javascript/api/outlook/office.dragovereventdata) or [DropEventData](/javascript/api/outlook/office.dropeventdata) object. Both these objects provide information on the position of the mouse pointer. Note that only the `DropEventData` object contains the data being transferred to the task pane.

When messages are dragged to the task pane, they're dropped as .eml files. Attachments that are dropped retain their current format. For a list of supported types, see [Supported item types](#supported-item-types).

The following example shows how to implement the drag-and-drop feature.

```javascript
// Handle the DragAndDropEvent event.
Office.context.mailbox.addHandlerAsync(
    Office.EventType.DragAndDropEvent,
    (event) => {
        console.log(`Event occurred: ${event.type}`);
        const eventData = event.dragAndDropEventData;

        // Get the file name and the contents of the items dropped into the task pane.
        if (eventData.type == "drop") {
            const files = eventData.dataTransfer.files;
            files.forEach((file) => {
                const content = file.fileContent;
                const name = file.name;

                // Add operations to process the item here, such as uploading the file to a CRM system.
            });
        }
    },
    (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("Failed to add event handler:", asyncResult.error.message); 
            return;
        }

        console.log("Event handler added successfully.");
    }
);
```

### Try it out

Try the interactive sample to see the drag-and-drop feature in action. Install the [Script Lab for Outlook add-in](https://appsource.microsoft.com/product/office/wa200001603) then run the "Drag and drop an item into the task pane" sample.

To learn more about Script Lab, see [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md).

# [Windows (classic) and Mac](#tab/desktop)

In Outlook on Windows (classic) and on Mac, use the [HTML Drag and Drop API](https://developer.mozilla.org/docs/Web/API/HTML_Drag_and_Drop_API) to handle the [DragEvent](https://developer.mozilla.org/docs/Web/API/DragEvent) DOM event. With the `DragEvent` object, you can identify when a user drags an item over the task pane, when they drop the item into the task pane, and what data is associated with the item. When a message is dragged and dropped into a task pane, its format varies depending on the Outlook client.
- **Windows (classic)**: Dropped as a .msg file.
- **Mac**: Dropped as a .eml file.

Attachments that are dropped into a task pane retain their current format. For a list of supported types, see [Supported item types](#supported-item-types).

> [!TIP]
> In classic Outlook on Windows, if you need the Base64-encoded .eml format to process a message, call [Office.context.mailbox.item.getAsFileAsync](/javascript/api/outlook/office.messageread#outlook-office-messageread-getasfileasync-member(1)).

---

## Feature behavior and limitations

### Supported item types

The following file types are supported by the drag-and-drop feature.

- **Messages**: Messages in the .eml or .msg format. Additionally, the following types of encrypted messages are also supported.
    - Messages encrypted using the S/MIME (Secure/Multipurpose Internet Mail Extensions) protocol.
    - Messages protected by Information Rights Management (IRM) with a sensitivity label that has the **Allow programmatic access** custom policy option set to `true`.
- **Attachments**: File types supported by Outlook. For guidance, see [Blocked attachments in Outlook](https://support.microsoft.com/office/434752e1-02d3-4e90-9124-8b81e49a8519).

> [!TIP]
> For information on the types of data that the HTML Drag and Drop API supports, see [Recommended drag types](https://developer.mozilla.org/docs/Web/API/HTML_Drag_and_Drop_API/Recommended_drag_types).

### Supported scenarios

The following table identifies which scenarios support the drag-and-drop feature in Outlook.

| Scenario | Supports drag and drop |
| ----- | ----- |
| Drag and drop a message or attachment from the [Reading Pane](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0) to an add-in's task pane in the same window | Supported |
| Drag and drop a message or an attachment from the Reading Pane to a task pane open in a different window | Supported |
| While a message is open in a different window, drag and drop an attachment contained in the message to a task pane open in main window of the Outlook client | Supported |
| While a message is open in a different window, drag and drop an attachment contained in the message to a task pane in the same window | Supported |
| Drag and drop multiple attachments at the same time | Supported |
| Drag and drop multiple messages at the same time | Supported |
| Drag and drop a mix of multiple messages and attachments at the same time | Not supported |
| Drag and drop a file from a task pane to the mailbox | Not supported |
| Drag and drop a file from the desktop to an add-in's task pane in Outlook | Depends on the drag-and-drop API used.<ul><li>**Office.js API**: Not supported</li><li>**HTML Drag and Drop API**: Supported</li></ul>|
| Drag and drop an item from another mailbox | Not supported |
| Drag and drop an item across two instances of the main window of the Outlook client | Not supported |
| Drag and drop an item across different Outlook clients | Not supported |

### Limitations

Be aware of the following limitations when implementing drag and drop in your add-in.

- If a user navigates to another mail item while an item that's been dragged to an add-in's task pane is being processed, the behavior varies depending on whether the task pane is pinned. If the task pane is pinned, processing isn't interrupted. Otherwise, processing fails. We recommend including progress indicators and displaying error messages for user awareness.
- Inline image attachments and links in messages can't be dropped into the task pane. For guidance on supported items, see [Supported item types](#supported-item-types).

## See also

- [Add-in commands](../design/add-in-commands.md)
- [Implement a pinnable task pane in Outlook](pinnable-taskpane.md)
- [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
