---
title: Insert slides in a PowerPoint presentation
description: Learn how to insert slides from one presentation into another.
ms.date: 03/07/2021
ms.localizationpriority: medium
---

# Insert slides in a PowerPoint presentation

A PowerPoint add-in can insert slides from one presentation into the current presentation by using PowerPoint's application-specific JavaScript library. You can control whether the inserted slides keep the formatting of the source presentation or the formatting of the target presentation.

The slide insertion APIs are primarily used in presentation template scenarios: There are a small number of known presentations which serve as pools of slides that can be inserted by the add-in. In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as slide titles or images) with slide IDs. The APIs can also be used in scenarios where the user can insert slides from any arbitrary presentation, but in that scenario the user is effectively limited to inserting *all* the slides from the source presentation. See [Selecting which slides to insert](#selecting-which-slides-to-insert) for more information about this.

There are two steps to inserting slides from one presentation into another.

1. Convert the source presentation file (.pptx) into a base64-formatted string.
1. Use the `insertSlidesFromBase64` method to insert one or more slides from the base64 file into the current presentation.

## Convert the source presentation to base64

There are many ways to convert a file to base64. Which programming language and library you use, and whether to convert on the server-side of your add-in or the client-side is determined by your scenario. Most commonly, you'll do the conversion in JavaScript on the client-side by using a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) object. The following example shows this practice.

1. Begin by getting a reference to the source PowerPoint file. In this example, we will use an `<input>` control of type `file` to prompt the user to choose a file. Add the following markup to the add-in page.

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    This markup adds the UI in the following screenshot to the page.

    ![Screenshot showing an HTML file type input control preceded by an instructional sentence reading "Select a PowerPoint presentation from which to insert slides". The control consists of a button labelled "Choose file" followed by the sentence "No file chosen".](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > There are many other ways to get a PowerPoint file. For example, if the file is stored on OneDrive or SharePoint, you can use Microsoft Graph to download it. For more information, see [Working with files in Microsoft Graph](/graph/api/resources/onedrive) and [Access Files with Microsoft Graph](/learn/modules/msgraph-access-file-data/).

2. Add the following code to the add-in's JavaScript to assign a function to the input control's `change` event. (You create the `storeFileAsBase64` function in the next step.)

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. Add the following code. Note the following about this code.

    - The `reader.readAsDataURL` method converts the file to base64 and stores it in the `reader.result` property. When the method completes, it triggers the `onload` event handler.
    - The `onload` event handler trims metadata off of the encoded file and stores the encoded string in a global variable.
    - The base64-encoded string is stored globally because it will be read by another function that you create in a later step.

    ```javascript
    let chosenFileBase64;

    async function storeFileAsBase64() {
        const reader = new FileReader();

        reader.onload = async (event) => {
            const startIndex = reader.result.toString().indexOf("base64,");
            const copyBase64 = reader.result.toString().substr(startIndex + 7);

            chosenFileBase64 = copyBase64;
        };

        const myFile = document.getElementById("file") as HTMLInputElement;
        reader.readAsDataURL(myFile.files[0]);
    }
    ```

## Insert slides with insertSlidesFromBase64

Your add-in inserts slides from another PowerPoint presentation into the current presentation with the [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-insertslidesfrombase64-member(1)) method. The following is a simple example in which all of the slides from the source presentation are inserted at the beginning of the current presentation and the inserted slides keep the formatting of the source file. Note that `chosenFileBase64` is a global variable that holds a base64-encoded version of a PowerPoint presentation file.

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

You can control some aspects of the insertion result, including where the slides are inserted and whether they get the source or target formatting , by passing an [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) object as a second parameter to `insertSlidesFromBase64`. The following is an example. About this code, note:

- There are two possible values for the `formatting` property: "UseDestinationTheme" and "KeepSourceFormatting". Optionally, you can use the `InsertSlideFormatting` enum, (e.g., `PowerPoint.InsertSlideFormatting.useDestinationTheme`).
- The function will insert the slides from the source presentation immediately after the slide specified by the `targetSlideId` property. The value of this property is a string of one of three possible forms: ***nnn*#**, **#*mmmmmmmmm***, or ***nnn*#*mmmmmmmmm***, where *nnn* is the slide's ID (typically 3 digits) and *mmmmmmmmm* is the slide's creation ID (typically 9 digits). Some examples are `267#763315295`, `267#`, and `#763315295`.

```javascript
async function insertSlidesDestinationFormatting() {
  await PowerPoint.run(async function(context) {
    context.presentation
    .insertSlidesFromBase64(chosenFileBase64,
                            {
                                formatting: "UseDestinationTheme",
                                targetSlideId: "267#"
                            }
                          );
    await context.sync();
  });
}
```

Of course, you typically won't know at coding time the ID or creation ID of the target slide. More commonly, an add-in will ask users to select the target slide. The following steps show how to get the ***nnn*#** ID of the currently selected slide and use it as the target slide.

1. Create a function that gets the ID of the currently selected slide by using the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) method of the Common JavaScript APIs. The following is an example. Note that the call to `getSelectedDataAsync` is embedded in a Promise-returning function. For more information about why and how to do this, see [Wrap Common-APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).

 
    ```javascript
    function getSelectedSlideID() {
      return new OfficeExtension.Promise<string>(function (resolve, reject) {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
          try {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              reject(console.error(asyncResult.error.message));
            } else {
              resolve(asyncResult.value.slides[0].id);
            }
          }
          catch (error) {
            reject(console.log(error));
          }
        });
      })
    }
    ```

1. Call your new function inside the [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function and pass the ID that it returns (concatenated with the "#" symbol) as the value of the `targetSlideId` property of the `InsertSlideOptions` parameter. The following is an example.

    ```javascript
    async function insertAfterSelectedSlide() {
        await PowerPoint.run(async function(context) {

            const selectedSlideID = await getSelectedSlideID();

            context.presentation.insertSlidesFromBase64(chosenFileBase64, {
                formatting: "UseDestinationTheme",
                targetSlideId: selectedSlideID + "#"
            });

            await context.sync();
        });
    }
    ```

### Selecting which slides to insert

You can also use the [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) parameter to control which slides from the source presentation are inserted. You do this by assigning an array of the source presentation's slide IDs to the `sourceSlideIds` property. The following is an example that inserts four slides. Note that each string in the array must follow one or another of the patterns used for the `targetSlideId` property.

```javascript
async function insertAfterSelectedSlide() {
    await PowerPoint.run(async function(context) {
        const selectedSlideID = await getSelectedSlideID();
        context.presentation.insertSlidesFromBase64(chosenFileBase64, {
            formatting: "UseDestinationTheme",
            targetSlideId: selectedSlideID + "#",
            sourceSlideIds: ["267#763315295", "256#", "#926310875", "1270#"]
        });

        await context.sync();
    });
}
```

> [!NOTE]
> The slides will be inserted in the same relative order in which they appear in the source presentation, regardless of the order in which they appear in the array.

There is no practical way that users can discover the ID or creation ID of a slide in the source presentation. For this reason, you can really only use the `sourceSlideIds` property when either you know the source IDs at coding time or your add-in can retrieve them at Configure your Office Add-in to use a shared runtime from some data source. Because users cannot be expected to memorize slide IDs, you also need a way to enable the user to select slides, perhaps by title or by an image, and then correlate each title or image with the slide's ID.

Accordingly, the `sourceSlideIds` property is primarily used in presentation template scenarios: The add-in is designed to work with a specific set of presentations that serve as pools of slides that can be inserted. In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as titles or images) with slide IDs or slide creation IDs that has been constructed from the set of possible source presentations.
