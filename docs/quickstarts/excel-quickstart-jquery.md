---
title: Build your first Excel task pane add-in
description: Learn how to build a simple Excel task pane add-in by using the Office JS API and the Yo Office tool.
ms.date: 12/11/2025
ms.service: excel
ms.localizationpriority: high
---

# Build an Excel task pane add-in

In this article, you'll walk through the process of building an Excel task pane add-in.

## Prerequisites

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## Create the add-in project

Decide the type of manifest that you'd like to use, either the **unified manifest for Microsoft 365** or the **add-in only manifest**. To learn more about them, see [Office Add-ins manifest](../develop/add-in-manifests.md).

# [Unified manifest for Microsoft 365 (preview)](#tab/jsonmanifest)

> [!NOTE]
> Using the unified manifest for Microsoft 365 with Excel add-ins is in public developer preview. The unified manifest for Microsoft 365 shouldn't be used in production Excel add-ins. We invite you to try it out in test or development environments. For more information, see the [Microsoft 365 app manifest schema reference](/microsoft-365/extensibility/schema).

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Choose a project type:** `Excel, PowerPoint, and/or Word Task Pane with unified manifest for Microsoft 365 (preview)`
- **What do you want to name your add-in?** `My Office Add-in`
- **Which Office client application would you like to support?** `Excel`

:::image type="content" source="../images/yo-office-excel-json-manifest-preview.png" alt-text="The Yeoman Generator for Office Add-ins command line interface when the unified manifest is selected.":::

# [Add-in only manifest](#tab/xmlmanifest)

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Choose a project type:** `Office Add-in Task Pane project`
- **Choose a script type:** `TypeScript`
- **What do you want to name your add-in?** `My Office Add-in`
- **Which Office client application would you like to support?** `Excel`

:::image type="content" source="../images/yo-office-excel-xml-manifest-ts.png" alt-text="The Yeoman Generator for Office Add-ins command line interface when the add-in only manifest is selected.":::

---

After you complete the wizard, the generator creates the project and installs supporting Node components.

## Explore the project

[!include[Yeoman generator add-in project components](../includes/yo-task-pane-project-components-ts.md)]

## Try it out

1. Navigate to the root folder of the project.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)]

1. In Excel, choose the **Home** tab, and then choose the **Show Task Pane** button on the ribbon to open the add-in task pane.

    :::image type="content" source="../images/excel-quickstart-add-in-3b.png" alt-text="The Excel Home menu, with the Show Task Pane button highlighted.":::

1. Select any range of cells in the worksheet.

1. At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.

    :::image type="content" source="../images/excel-quickstart-add-in-3c.png" alt-text="The add-in task pane open in Excel, with the Run button highlighted in the add-in task pane.":::

1. [!include[Instructions to stop web server and uninstall dev add-in](../includes/stop-uninstall-dev-add-in.md)]

[!include[The common troubleshooting section for all Yo Office quick starts](../includes/quickstart-troubleshooting-yo.md)]

## Next steps

Congratulations, you've successfully created an Excel task pane add-in! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the [Excel add-in tutorial](../tutorials/excel-tutorial.md).

## Code samples

- [Excel "Hello world" add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/excel-hello-world): Learn how to build a simple Office Add-in with only a manifest, HTML web page, and a logo.

## See also

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Develop Office Add-ins](../develop/develop-overview.md)
- [Excel JavaScript object model in Office Add-ins](../excel/excel-add-ins-core-concepts.md)
- [Excel add-in code samples](https://developer.microsoft.com/microsoft-365/gallery/?filterBy=Excel,Samples)
- [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md)
- [Using Visual Studio Code to publish](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)
