---
title: Build your first Excel content add-in
description: Learn how to build a simple Excel content add-in by using the Office JS API.
ms.date: 06/25/2024
ms.service: excel
ms.localizationpriority: medium
---

# Build an Excel content add-in

In this article, you'll walk through the process of building an Excel [content add-in](../design/content-add-ins.md) using Visual Studio.

## Prerequisites

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

## Create the add-in project

1. In Visual Studio, choose **Create a new project**.

1. Using the search box, enter **add-in**. Choose **Excel Web Add-in**, then select **Next**.

1. Name your project **ExcelWebAddIn1** and select **Create**.

1. In the **Create Office Add-in** dialog window, choose the **Insert content into Excel spreadsheets** add-in type, then choose **Next**.

1. Choose the **Basic Add-in** or **Document Visualization Add-in** add-in template, and then choose **Finish** to create the project.

1. Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.

## Explore the Visual Studio solution

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

## Update the manifest

1. In **Solution Explorer**, go to the **ExcelWebAddIn1** add-in project and open the **ExcelWebAddIn1Manifest** directory. This directory contains your manifest file, **ExcelWebAddIn1.xml**. The manifest file defines the add-in's settings and capabilities. See the preceding section [Explore the Visual Studio solution](#explore-the-visual-studio-solution) for more information about the two projects created by your Visual Studio solution.

1. The `ProviderName` element has a placeholder value. Replace it with your name.

1. The `DefaultValue` attribute of the `DisplayName` element has a placeholder. Replace it with **My Office Add-in**.

1. The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A content add-in for Excel.**.

1. Save the file. The updated lines should look like the following code sample.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A content add-in for Excel."/>
    ...
    ```

## Try it out

1. Using Visual Studio, test the newly created Excel add-in by pressing <kbd>F5</kbd> or choosing the **Start** button to launch Excel with the content add-in displayed in the spreadsheet.

1. Ensure that there's text in the worksheet, then select any range of cells containing text in the worksheet.

1. Select the tab for the template you chose, then follow the instructions.

    # [Basic Add-in](#tab/basic)

    - In the content add-in, choose the **Get data from selection** button to get the text from the selected range.

      :::image type="content" source="../images/excel-quickstart-content-basic-ui.png" alt-text="The add-in content open in Excel.":::

    # [Document Visualization Add-in](#tab/advanced)

    - In the content add-in, choose the **Insert sample data** button to add sample data to the worksheet and display a visualization.

      :::image type="content" source="../images/excel-quickstart-content-advanced-visualization.png" alt-text="The add-in content visualization open in Excel.":::

    ---

[!include[Console tool note](../includes/console-tool-note.md)]

## Next steps

Congratulations, you've successfully created an Excel content add-in! Next, learn more about [developing Office Add-ins with Visual Studio](../develop/develop-add-ins-visual-studio.md).

[!include[The common troubleshooting section for all Visual Studio quick starts](../includes/quickstart-troubleshooting-vs.md)]

## Code samples

- [Excel content add-in: Humongous Insurance](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-content-add-in)

## See also

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Develop Office Add-ins](../develop/develop-overview.md)
- [Excel JavaScript object model in Office Add-ins](../excel/excel-add-ins-core-concepts.md)
- [Excel add-in code samples](https://developer.microsoft.com/microsoft-365/gallery/?filterBy=Excel,Samples)
- [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md)
- [Using Visual Studio Code to publish](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)
