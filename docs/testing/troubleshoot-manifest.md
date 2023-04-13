---
title: Validate an Office Add-in's manifest
description: Learn how to validate the manifest of an Office Add-in using the XML schema and other tools.
ms.date: 04/14/2023
ms.localizationpriority: medium
---

# Validate an Office Add-in's manifest

You may want to validate your add-in's manifest file to ensure that it's correct and complete. Validation can also identify issues that are causing the error "Your add-in manifest is not valid" when you attempt to sideload your add-in. This article describes multiple ways to validate the manifest file.

> [!NOTE]
> For details about using runtime logging to troubleshoot issues with your add-in's manifest, see [Debug your add-in with runtime logging](runtime-logging.md).

## Validate your manifest with the Yeoman generator for Office Add-ins

If you used the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md) to create your add-in, you can also use it to validate your project's manifest file. Run the following command in the root directory of your project.

```command&nbsp;line
npm run validate
```

![Animated GIF that shows the Yo Office validator being run at the command line and generating results that show Validation Passed.](../images/yo-office-validator.gif)

> [!NOTE]
> To access this functionality, your add-in project must be created using the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md) version 1.1.17 or later.

## Validate your manifest with office-addin-manifest

If you didn't use the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md) to create your add-in, you can validate the manifest by using [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).

1. Install [Node.js](https://nodejs.org/download/).

1. Open a command prompt and install the validator with the following command.

    ```command&nbsp;line
    npm install -g office-addin-manifest
    ```

1. Run the following command *in the root directory of your project*.

    ```command&nbsp;line
    npm run validate
    ```

    > [!NOTE]
    > If this command is not available or not working, run the following command instead to force the use of the latest version of the office-addin-manifest tool (replacing `MANIFEST_FILE` with the name of the manifest file).
    >
    > ```command&nbsp;line
    > npx office-addin-manifest validate MANIFEST_FILE
    > ```

[!INCLUDE [validate also runs Office Store validation](../includes/office-store-validate.md)]

If you're having trouble with that command, try the following (replacing `MANIFEST_FILE` with the name of the manifest file).

```command&nbsp;line
npx office-addin-manifest validate -p MANIFEST_FILE
```

## Validate your manifest against the XML schema

You can validate the manifest file against the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) files. This will ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using. If you copied elements from other sample manifests double check that you also **include the appropriate namespaces**. You can use an XML schema validation tool to perform this validation.

### To use a command-line XML schema validation tool to validate your manifest

1. Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.

1. Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.

    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## See also

- [Office Add-ins XML manifest](../develop/add-in-manifests.md)
- [Clear the Office cache](clear-cache.md)
- [Debug your add-in with runtime logging](runtime-logging.md)
- [Sideload Office Add-ins for testing](sideload-office-add-ins-for-testing.md)
- [Debug add-ins using developer tools for Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
- [Debug add-ins using developer tools for Edge Legacy](debug-add-ins-using-devtools-edge-legacy.md)
- [Debug add-ins using developer tools in Microsoft Edge (Chromium-based)](debug-add-ins-using-devtools-edge-chromium.md)
