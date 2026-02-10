---
title: First-run experience tutorial
description: Learn how to implement a first-run experience for your Office Add-in.
ms.date: 02/10/2026
ms.topic: how-to
ms.localizationpriority: medium
---

# Build an Office Add-in with a basic first-run experience

Implementing a first-run experience helps onboard users and highlights your add-in's value. This tutorial guides you through adding a simple first-run experience to your Office Add-in using local storage to track whether the user has previously launched the add-in.

> [!NOTE]
> The [first-run experience](../design/first-run-experience-patterns.md) is a recommended pattern for Office Add-ins. It helps users understand your add-in's features and increases engagement.

## What you'll learn

- How to add a first-run UI to your add-in.
- How to use [browser local storage](../develop/persisting-add-in-state-and-settings.md#browser-storage) to persist user state.
- How to update your add-in's HTML, TypeScript or JavaScript, and CSS files to support the first-run experience.

## Overview

When a user opens your add-in for the first time, you'll display a welcome message and a list of key features. On subsequent launches, the add-in will skip the welcome and show the main UI. This is accomplished by checking for a flag in local storage and updating the UI accordingly.

This tutorial provides instructions and screenshots for Excel but you can use a similar pattern to implement a first-run experience in other Office applications where Office Web Add-ins are supported.

## Steps

Follow these steps to implement the first-run experience:

1. **Update the HTML**: Add a container for the first-run experience.
2. **Update the TypeScript or JavaScript**: Check local storage and display the first-run UI if needed.
3. **Update the CSS**: Ensure the new UI is styled consistently.
4. **Test your add-in**: Verify the first-run experience works as expected.

Let's get started!

> [!TIP]
> If you want a completed version of this tutorial, visit the [Office Add-ins samples repo on GitHub](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/tutorials/first-run-experience-tutorial).

## Prerequisites

1. Select the Yo Office quick start for the Office application you'd like to use.

    - [Excel](../quickstarts/excel-quickstart-jquery.md)
    - [OneNote](../quickstarts/onenote-quickstart.md)
    - [Outlook](../quickstarts/outlook-quickstart-yo.md)
    - [PowerPoint](../quickstarts/powerpoint-quickstart-yo.md)
    - [Project](../quickstarts/project-quickstart.md)
    - [Word](../quickstarts/word-quickstart-yo.md)

1. Follow the instructions in your selected quick start. After you complete its "Try it out" section, return here to continue.

## Implement the first-run experience

### Update the HTML file

Be clear about the area of the UI that will be part of the first-run experience. In this tutorial, you'll create a `<div>` element with the `id` named "first-run-experience" that represents what users see only the first time they run your add-in.

1. Open the **taskpane.html**. Replace the `<main>` element with the following markup, then save the file. Some notes about this markup:

    - The "first-run-experience" `<div>` is inserted in the `<main>` element. It surrounds the list of Office Add-ins features. By default, this `<div>` isn't displayed.
    - The first `<p>` element provides the user with instructions for using the add-in.

    ```html
    <main id="app-body" class="ms-welcome__main" style="display: none;">
        <div id="first-run-experience" style="display: none;">
            <h2 class="ms-font-xl"> Discover what Office Add-ins can do for you today! </h2>
            <ul class="ms-List ms-welcome__features">
                <li class="ms-ListItem">
                    <i class="ms-Icon ms-Icon--Ribbon ms-font-xl"></i>
                    <span class="ms-font-m">Achieve more with Office integration</span>
                </li>
                <li class="ms-ListItem">
                    <i class="ms-Icon ms-Icon--Unlock ms-font-xl"></i>
                    <span class="ms-font-m">Unlock features and functionality</span>
                </li>
                <li class="ms-ListItem">
                    <i class="ms-Icon ms-Icon--Design ms-font-xl"></i>
                    <span class="ms-font-m">Create and visualize like a pro</span>
                </li>
            </ul>
        </div>
        <p class="ms-font-l">Select any range of cells in the worksheet, then click <b>Run</b>.</p>
        <div role="button" id="run" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
            <span class="ms-Button-label">Run</span>
        </div>
        <p><label id="item-subject"></label></p>    
    </main>
    ```

1. If you selected an Office application besides Excel, update the first `<p>` element with more appropriate instructions.

### Update the TypeScript or JavaScript file

Update the TypeScript or JavaScript file to display the first-run experience if this is the first time the user is running the add-in.

1. Open the **taskpane.ts** or **taskpane.js** file. Replace the `Office.onReady` statement with the following code, then save the file. Some notes about this code:

    - It checks local storage for a key called "showedFRE". If the key doesn't exist, then show the first-run experience.
    - It adds a new function called `showFirstRunExperience` that displays the "first-run-experience" `<div>` added to the HTML. This function also adds the "showedFRE" item to local storage.

    ```javascript
    Office.onReady((info) => {
      if (info.host === Office.HostType.Excel) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";

        // showedFRE is created and set to "true" when you call showFirstRunExperience().
        if (!localStorage.getItem("showedFRE")) {
          showFirstRunExperience();
        }
    
        document.getElementById("run").onclick = run;
      }
    });
    
    async function showFirstRunExperience() {
      document.getElementById("first-run-experience").style.display = "flex";
      localStorage.setItem("showedFRE", "true");
    }  
    ```

1. If you selected an Office application besides Excel, update the condition of the first `if` statement to check for your chosen [Office.HostType](/javascript/api/office/office.hosttype).

### Update the CSS file

Update the CSS file to ensure that the add-in UI is styled appropriately given the addition of the "first-run-experience" `<div>`.

- Open the **taskpane.css** file. Replace the line `.ms-welcome__main {` with the following code, then save the file.

    ```css
    .ms-welcome__main, .ms-welcome__main > div {
    ```

## Try it out

1. Ensure that the web server is running and the add-in has been sideloaded, then open the task pane. For details, see the instructions in the quick start you used.

1. Verify that the task pane includes the list of features.

    :::image type="content" source="../images/fre-tutorial-addin-first-run.png" alt-text="The add-in task pane UI on first run.":::

1. Close the task pane then reopen it. Verify that the task pane no longer displays the list of features.

    :::image type="content" source="../images/fre-tutorial-addin-next-run.png" alt-text="The add-in task pane UI on second run.":::

## Next steps

Congratulations, you've successfully created an Office task pane add-in with a first-run experience!

### Make it production ready

Using this tutorial, you implemented a basic [first-run experience](../design/first-run-experience-patterns.md). For the first-run experience to be ready for users, you should consider the following:

- Update the features listed in the value placemat to match what your add-in actually does.
- Implement a different pattern (for example, video placemat or carousel) that better showcases the benefits of your add-in.
- Use a more secure and robust option for tracking first-run state. For example, use storage partitioning if available, or implement a Single Sign-on (SSO) authentication solution.
  - For more about available settings options, see [Persist add-in state and settings](../develop/persisting-add-in-state-and-settings.md).
  - For more about available authentication options, see [Overview of authentication and authorization](../develop/overview-authn-authz.md).

If you're planning to make your add-in available in Microsoft Marketplace, you must have a robust and useful first-run experience. For more information, see [Best practices for developing Office Add-ins](../concepts/add-in-development-best-practices.md).

## Code samples

- [Completed first-run experience tutorial](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/tutorials/first-run-experience-tutorial): The result of completing this tutorial with Excel.

## See also

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Best practices for developing Office Add-ins](../concepts/add-in-development-best-practices.md)
- [Develop Office Add-ins](../develop/develop-overview.md)
- [Office Add-in code samples](../overview/office-add-in-code-samples.md)
- [Using Visual Studio Code to publish](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)
