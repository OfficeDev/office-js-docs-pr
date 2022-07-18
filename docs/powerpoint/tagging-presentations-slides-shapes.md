---
title: Use custom tags on presentations, slides, and shapes in PowerPoint
description: Learn how to use tags for custom metadata about presentations, slides, and shapes.
ms.date: 12/14/2021
ms.localizationpriority: medium
---

# Use custom tags for presentations, slides, and shapes in PowerPoint

An add-in can attach custom metadata, in the form of key-value pairs, called "tags", to presentations, specific slides, and specific shapes on a slide.

There are two main scenarios for using tags:

- When applied to a slide or a shape, a tag enables the object to be categorized for batch processing. For example, suppose a presentation has some slides that should be included in presentations to the East region but not the West region. Similarly, there are alternative slides that should be shown only to the West. Your add-in can create a tag with the key `REGION` and the value `East` and apply it to the slides that should only be used in the East. The tag's value is set to `West` for the slides that should only be shown to the West region. Just before a presentation to the East, a button in the add-in runs code that loops through all the slides checking the value of the `REGION` tag. Slides where the region is `West` are deleted. The user then closes the add-in and starts the slide show.
- When applied to a presentation, a tag is effectively a custom property in the presentation document (similar to a [CustomProperty](/javascript/api/word/word.customproperty) in Word).

## Tag slides and shapes

A tag is a key-value pair, where the value is always of type `string` and is represented by a [Tag](/javascript/api/powerpoint/powerpoint.tag) object. Each type of parent object, such as a [Presentation](/javascript/api/powerpoint/powerpoint.presentation), [Slide](/javascript/api/powerpoint/powerpoint.slide), or [Shape](/javascript/api/powerpoint/powerpoint.shape) object, has a `tags` property of type [TagsCollection](/javascript/api/powerpoint/powerpoint.tagcollection).

### Add, update, and delete tags

To add a tag to an object, call the [TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-add-member(1)) method of the parent object's `tags` property. The following code adds two tags to the first slide of a presentation. About this code, note:

- The first parameter of the `add` method is the key in the key-value pair.
- The second parameter is the value.
- The key is in uppercase letters. This isn't strictly mandatory for the `add` method; however, the key is always stored by PowerPoint as uppercase, and *some tag-related methods do require that the key be expressed in uppercase*, so we recommend as a best practice that you always use uppercase in your code for a tag key.

```javascript
async function addMultipleSlideTags() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("OCEAN", "Arctic");
    slide.tags.add("PLANET", "Jupiter");

    await context.sync();
  });
}
```

The `add` method is also used to update a tag. The following code changes the value of the `PLANET` tag.

```javascript
async function updateTag() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("PLANET", "Mars");

    await context.sync();
  });
}
```

To delete a tag, call the `delete` method on it's parent `TagsCollection` object and pass the key of the tag as the parameter. For an example, see [Set custom metadata on the presentation](#set-custom-metadata-on-the-presentation).

### Use tags to selectively process slides and shapes

Consider the following scenario: Contoso Consulting has a presentation they show to all new customers. But some slides should only be shown to customers that have paid for "premium" status. Before showing the presentation to non-premium customers, they make a copy of it and delete the slides that only premium customers should see. An add-in enables Contoso to tag which slides are for premium customers and to delete these slides when needed. The following list outlines the major coding steps to create this functionality.

1. Create a function that tags the currently selected slide as intended for `Premium` customers. About this code, note:

    - The `getSelectedSlideIndex` function is defined in the next step. It returns the 1-based index of the currently selected slide.
    - The value returned by the `getSelectedSlideIndex` function has to be decremented because the [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitemat-member(1)) method is 0-based.

    ```javascript
    async function addTagToSelectedSlide() {
      await PowerPoint.run(async function(context) {
        let selectedSlideIndex = await getSelectedSlideIndex();
        selectedSlideIndex = selectedSlideIndex - 1;
        const slide = context.presentation.slides.getItemAt(selectedSlideIndex);
        slide.tags.add("CUSTOMER_TYPE", "Premium");
    
        await context.sync();
      });
    }
    ```

2. The following code creates a method to get the index of the selected slide. About this code, note:

    - It uses the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) method of the Common JavaScript APIs.
    - The call to `getSelectedDataAsync` is embedded in a promise-returning function. For more information about why and how to do this, see [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).
    - `getSelectedDataAsync` returns an array because multiple slides can be selected. In this scenario, the user has selected just one, so the code gets the first (0th) slide, which is the only one selected.
    - The `index` value of the slide is the 1-based value the user sees beside the slide in the PowerPoint UI thumbnails pane.

    ```javascript
    function getSelectedSlideIndex() {
        return new OfficeExtension.Promise<number>(function(resolve, reject) {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function(asyncResult) {
                try {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        reject(console.error(asyncResult.error.message));
                    } else {
                        resolve(asyncResult.value.slides[0].index);
                    }
                } 
                catch (error) {
                    reject(console.log(error));
                }
            });
        });
    }
    ```

3. The following code creates a function to delete slides that are tagged for premium customers. About this code, note:

    - Because the `key` and `value` properties of the tags are going to be read after the `context.sync`, they must be loaded first.

    ```javascript
    async function deleteSlidesByAudience() {
      await PowerPoint.run(async function(context) {
        const slides = context.presentation.slides;
        slides.load("tags/key, tags/value");
    
        await context.sync();
    
        for (let i = 0; i < slides.items.length; i++) {
          let currentSlide = slides.items[i];
          for (let j = 0; j < currentSlide.tags.items.length; j++) {
            let currentTag = currentSlide.tags.items[j];
            if (currentTag.key === "CUSTOMER_TYPE" && currentTag.value === "Premium") {
              currentSlide.delete();
            }
          }
        }
    
        await context.sync();
      });
    }
    ```

## Set custom metadata on the presentation

Add-ins can also apply tags to the presentation as a whole. This enables you to use tags for document-level metadata similar to how the [CustomProperty](/javascript/api/word/word.customproperty)class is used in Word. But unlike the Word `CustomProperty` class, the value of a PowerPoint tag can only be of type `string`.

The following code is an example of adding a tag to a presentation. 

```javascript
async function addPresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.add("SECURITY", "Internal-Audience-Only");

    await context.sync();
  });
}
```

The following code is an example of deleting a tag from a presentation. Note that the key of the tag is passed to the `delete` method of the parent `TagsCollection` object.

```javascript
async function deletePresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.delete("SECURITY");

    await context.sync();
  });
}
```
