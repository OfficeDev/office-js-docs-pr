---
title: Use the Office Dialog to show a video
description: 'Shows how to open a video in the Office Dialog'
ms.date: 01/16/2020
localization_priority: Normal
---

# Use the Office Dialog to show a video

> [!NOTE]
> This article presupposes that you are familiar with the basics of using the Office Dialog as described in [Use the Dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).

To show a video in a dialog box with the Office Dialog API take these steps:

1. Create a page whose only content is an iframe. The `src` attribute of the iframe points to an online video. The protocol of the video's URL must be HTTP**S**. In this article we'll call this page "video.dialogbox.html". The following is an example of the markup:

    ```HTML
    <iframe class="ms-firstrun-video__player"  width="640" height="360"
        src="https://www.youtube.com/embed/XVfOe5mFbAE?rel=0&autoplay=1"
        frameborder="0" allowfullscreen>
    </iframe>
    ```

2. The video.dialogbox.html page must be in the same domain as the host page.
3. Use a call of `displayDialogAsync` in the host page to open video.dialogbox.html.
4. If your add-in needs to know when the user closes the dialog box, register a handler for the `DialogEventReceived` event and handle the 12006 event. For details, see the section [Errors and events in the dialog](errors-and-events-in-the-dialog-window.md).

For a sample that shows a video in a dialog box, see the [video placemat design pattern](/office/dev/add-ins/design/first-run-experience-patterns#video-placemat).

![Screenshot of a video showing in an add-in dialog box](../images/video-placemats-dialog-open.png)