---
title: Use the Office dialog box to play a video
description: Learn how to open and play a video in the Office dialog box.
ms.date: 07/08/2021
ms.topic: how-to
ms.localizationpriority: medium
---

# Use the Office dialog box to show a video

This article explains how to play a video in an Office Add-in dialog box.

> [!NOTE]
> This article presumes you're familiar with the basics of using the Office dialog box as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).

To play a video in a dialog box with the Office dialog API, follow these steps.

1. Create a page containing an iframe and no other content. The page must be in the same domain as the host page. For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page). In the `src` attribute of the iframe, point to the URL of an online video. The protocol of the video's URL must be HTTPS. In this article, we'll call this page "video.dialogbox.html". The following is an example of the markup.

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.
3. If your add-in needs to know when the user closes the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 event. For details, see [Errors and events in the Office dialog box](dialog-handle-errors-events.md).

For a sample of a video playing in a dialog box, see the [video placemat design pattern](../design/first-run-experience-patterns.md#video-placemat).

![A video playing in an add-in dialog box in front of Excel.](../images/video-placemats-dialog-open.png)
