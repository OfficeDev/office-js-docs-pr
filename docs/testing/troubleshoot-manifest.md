---
title: Validate an Office Add-in's manifest
description: Learn how to validate the manifest of an Office Add-in.
ms.date: 05/19/2025
ms.localizationpriority: medium
---

# Validate an Office Add-in's manifest

You should validate your add-in's manifest file to ensure that it's correct and complete. Validation can also identify issues that are causing the error "Your add-in manifest is not valid" when you attempt to sideload your add-in. This article describes multiple ways to validate the manifest file. Except as specified otherwise, they work for both the unified manifest for Microsoft 365 and the add-in only manifest.

> [!NOTE]
> For details about using runtime logging to troubleshoot issues with your add-in's manifest, see [Debug your add-in with runtime logging](runtime-logging.md).

## Validate your manifest with the validate command

If you used [Microsoft 365 Agents Toolkit](../develop/agents-toolkit-overview.md) or [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md) to create your add-in, you can validate your project's manifest file with the following command in the root directory of your project.

```command&nbsp;line
npm run validate
```

[!INCLUDE [validate also runs Microsoft 365 and Copilot store validation](../includes/office-store-validate.md)]

If you're having trouble with that command, try the following (replacing `MANIFEST_FILE` with the name of the manifest file).

```command&nbsp;line
npx office-addin-manifest validate -p MANIFEST_FILE
```

## Validate your manifest with office-addin-manifest

If you didn't use [Microsoft 365 Agents Toolkit](../develop/agents-toolkit-overview.md) or [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md) to create your add-in, you can validate the manifest by using [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).

1. Install [Node.js](https://nodejs.org/download/).

1. Open a command prompt and install the validator with the following command.

    ```command&nbsp;line
    npm install -g office-addin-manifest
    ```

1. Run the following command *in the folder of your project that contains the manifest file* (replacing `MANIFEST_FILE` with the name of the manifest file).

    ```command&nbsp;line
    office-addin-manifest validate MANIFEST_FILE
    ```

    > [!NOTE]
    > If this command isn't working, run the following command instead to force the use of the latest version of the office-addin-manifest tool (replacing `MANIFEST_FILE` with the name of the manifest file).
    >
    > ```command&nbsp;line
    > npx office-addin-manifest validate MANIFEST_FILE
    > ```

## Validate the manifest in the UI of Agents Toolkit

If you're working in Agents Toolkit and using the unified manifest, you can use the toolkit's validation options. For instructions, see [Validate application](/microsoftteams/platform/toolkit/teamsfx-preview-and-customize-app-manifest#validate-application).

## See also

- [Office Add-ins manifest](../develop/add-in-manifests.md)
- [Clear the Office cache](clear-cache.md)
- [Debug your add-in with runtime logging](runtime-logging.md)
- [Sideload Office Add-ins for testing](sideload-office-add-ins-for-testing.md)
- [Debug add-ins using developer tools for Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
- [Debug add-ins using developer tools for Edge Legacy](debug-add-ins-using-devtools-edge-legacy.md)
- [Debug add-ins using developer tools in Microsoft Edge (Chromium-based)](debug-add-ins-using-devtools-edge-chromium.md)
