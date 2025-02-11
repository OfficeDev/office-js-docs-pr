---
title: Build your first PowerPoint content add-in
description: Learn how to build a simple PowerPoint content add-in by using the Office JS API.
ms.date: 07/08/2024
ms.service: powerpoint
ms.localizationpriority: medium
---

# Build your first PowerPoint content add-in

In this article, you'll walk through the process of building a PowerPoint [content add-in](../design/content-add-ins.md) using Visual Studio.

## Prerequisites

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### Create the add-in project

1. In Visual Studio, choose **Create a new project**.

1. Using the search box, enter **add-in**. Choose **PowerPoint Web Add-in**, then select **Next**.

1. Name your project and select **Create**.

1. In the **Create Office Add-in** dialog window, choose **Insert content into PowerPoint slides**, and then choose **Finish** to create the project.

1. Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.

## Explore the Visual Studio solution

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

## Update the code

1. **Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, find the `<p>` element that contains the text "This example will read the current document selection." and the `<button>` element where the `id` is "get-data-from-selection". Replace these entire elements with the following markup then save the file.

    ```html
    <p class="ms-font-m-plus">This example will get some details about the current slide.</p>

    <button class="Button Button--primary" id="get-data-from-selection">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Get slide details</span>
        <span class="Button-description">Gets and displays the current slide's details.</span>
    </button>
    ```

1. Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Find the `getDataFromSelection` function and replace the entire function with the following code then save the file.

    ```js
    // Gets some details about the current slide and displays them in a notification.
    function getDataFromSelection() {
        if (Office.context.document.getSelectedDataAsync) {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        showNotification('Some slide details are:', '"' + JSON.stringify(result.value) + '"');
                    } else {
                        showNotification('Error:', result.error.message);
                    }
                }
            );
        } else {
            app.showNotification('Error:', 'Reading selection data is not supported by this host application.');
        }
    }
    ```

## Update the manifest

1. Open the add-in only manifest file in the add-in project. This file defines the add-in's settings and capabilities.

1. The `ProviderName` element has a placeholder value. Replace it with your name.

1. The `DefaultValue` attribute of the `DisplayName` element has a placeholder. Replace it with **My Office Add-in**.

1. The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A content add-in for PowerPoint.**.

1. Save the file. The updated lines should look like the following code sample.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A content add-in for PowerPoint."/>
    ...
    ```

## Try it out

1. Using Visual Studio, test the newly created PowerPoint add-in by pressing <kbd>F5</kbd> or choosing the **Start** button to launch PowerPoint with the content add-in displayed over the slide.

1. In PowerPoint, choose the **Get slide details** button in the content add-in to get details about the current slide.

    :::image type="content" source="../images/powerpoint-quickstart-content-ui.png" alt-text="The add-in content open in PowerPoint.":::

[!include[Console tool note](../includes/console-tool-note.md)]

## Next steps

Congratulations, you've successfully created a PowerPoint content add-in! Next, learn more about [developing Office Add-ins with Visual Studio](../develop/develop-add-ins-visual-studio.md).

[!include[The common troubleshooting section for all Visual Studio quick starts](../includes/quickstart-troubleshooting-vs.md)]

## See also

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Develop Office Add-ins](../develop/develop-overview.md)
- [Using Visual Studio Code to publish](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)
