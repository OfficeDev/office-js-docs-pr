---
title: Validate and troubleshoot issues with your manifest
description: Learn how to validate an Office Add-in's manifest.
ms.date: 12/31/2019
localization_priority: Priority
---

# Validate and troubleshoot issues with your manifest

You may want to validate your add-in's manifest file to ensure that it's correct and complete. Validation can also identify issues that are causing the error "Your add-in manifest is not valid" when you attempt to sideload your add-in. This article describes multiple ways to validate the manifest file.

> [!NOTE]
> For details about using runtime logging to troubleshoot issues with your add-in's manifest, see [Debug your add-in with runtime logging](runtime-logging.md).

## Validate your manifest with the Yeoman generator for Office Add-ins

If you used the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can also use it to validate your project's manifest file. Run the following command in the root directory of your project:

```command&nbsp;line
npm run validate
```

![Animated gif that shows the Yo Office validator being run at the command line and generating results that show Validation Passed](../images/yo-office-validator.gif)

> [!NOTE]
> To have access to this functionality, your add-in project must have been created by using [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) version 1.1.17 or later.

## Validate your manifest with office-addin-manifest

If you didn't use the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can validate the manifest by using [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).

1. Install [Node.js](https://nodejs.org/download/).

2. Run the following command in the root directory of your project. Replace `MANIFEST_FILE` with the name of the manifest file.

	```command&nbsp;line
	npx office-addin-manifest validate MANIFEST_FILE
	```

	> [!NOTE]
	> If running this command results in the error message "The command syntax is not valid." (because the `validate` command is not recognized), run the following command to validate the manifest (replacing `MANIFEST_FILE` with the name of the manifest file): 
	> 
	> `npx --ignore-existing office-addin-manifest validate MANIFEST_FILE`

## Validate your manifest against the XML schema

You can validate the manifest file against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) files. This will ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using. If you copied elements from other sample manifests double check that you also **include the appropriate namespaces**. You can use an XML schema validation tool to perform this validation.

### To use a command-line XML schema validation tool to validate your manifest

1. Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.

2. Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.
	
	```command&nbsp;line
	xmllint --noout --schema XSD_FILE XML_FILE
	```

## See also

- [Office Add-ins XML manifest](../develop/add-in-manifests.md)
- [Clear the Office cache](clear-cache.md)
- [Debug your add-in with runtime logging](runtime-logging.md)
- [Sideload Office Add-ins for testing](sideload-office-add-ins-for-testing.md)
- [Debug Office Add-ins](debug-add-ins-using-f12-developer-tools-on-windows-10.md)