---
title: Validate and troubleshoot issues with your manifest
description: Use these methods to validate the Office Add-ins manifest.
ms.date: 05/21/2019
localization_priority: Priority
---

# Validate and troubleshoot issues with your manifest

Use these methods to validate and troubleshoot issues in your Office Add-ins manifest: 

- [Validate your manifest with the Office Add-in Validator](#validate-your-manifest-with-the-office-add-in-validator)	
- [Validate your manifest against the XML schema](#validate-your-manifest-against-the-xml-schema)
- [Validate your manifest with the Yeoman generator for Office Add-ins](#validate-your-manifest-with-the-yeoman-generator-for-office-add-ins)
- [Use runtime logging to debug your add-in](#use-runtime-logging-to-debug-your-add-in)


## Validate your manifest with the Office Add-in Validator

To help ensure that the manifest file that describes your Office Add-in is correct and complete, validate it against the [Office Add-in Validator](https://github.com/OfficeDev/office-addin-validator).

### To use the Office Add-in Validator to validate your manifest

1. Install [Node.js](https://nodejs.org/download/). 

2. Open a command prompt / terminal as an administrator, and install the Office Add-in Validator and its dependencies globally by using the following command:

	```command&nbsp;line
	npm install -g office-addin-validator
	```
	
	> [!NOTE]
	> If you already have Yo Office installed, upgrade to the latest version, and the validator will be installed as a dependency.

3. Run the following command to validate your manifest. Replace MANIFEST.XML with the path to the manifest XML file.

	```command&nbsp;line
	validate-office-addin MANIFEST.XML
	```

## Validate your manifest against the XML schema

To help ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using. If you copied elements from other sample manifests double check you also **include the appropriate namespaces**. You can validate a manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) files. You can use an XML schema validation tool to perform this validation. 



### To use a command-line XML schema validation tool to validate your manifest

1.	Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.

2.	Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.
	
	```command&nbsp;line
	xmllint --noout --schema XSD_FILE XML_FILE
	```

## Validate your manifest with the Yeoman generator for Office Add-ins

If you've created your Office Add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office), you can ensure that the manifest file follows the correct schema by running the following command within the root directory of your project:

```command&nbsp;line
npm run validate
```

![Animated gif that shows the Yo Office validator being run at the command line and generating results that show Validation Passed](../images/yo-office-validator.gif)

> [!NOTE]
> To have access to this functionality, your add-in project must have been created by using [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) version 1.1.17 or later.

## Use runtime logging to debug your add-in 

You can use runtime logging to debug your add-in's manifest as well as several installation errors. This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs. Runtime logging is particularly  useful for debugging add-ins that implement add-in commands and Excel custom functions.   

> [!NOTE]
> The runtime logging feature is currently available for Office 2016 desktop.

### To turn on runtime logging

> [!IMPORTANT]
> Runtime Logging affects performance. Turn it on only when you need to debug issues with your add-in manifest.

To turn on runtime logging:

1. Make sure that you are running Office 2016 desktop build **16.0.7019** or later. 

2. Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`. 

    > [!NOTE]
    > If the `Developer` key (folder) does not already exist under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, complete the following steps to create it: 
	> 1. Right-click the **WEF** key (folder) and select **New** > **Key**.
	> 2. Name the new key **Developer**.

3. Set the default value of the key to the full path of the file where you want the log to be written. For an example, see [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip). 

    > [!NOTE]
    > The directory in which the log file will be written must already exist, and you must have write permissions to it. 
 
The following image shows what the registry should look like. To turn the feature off, remove the `RuntimeLogging` key from the registry. 

![Screenshot of the registry editor with a RuntimeLogging registry key](http://i.imgur.com/Sa9TyI6.png)


### To troubleshoot issues with your manifest

To use runtime logging to troubleshoot issues loading an add-in:
 
1. [Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing. 

	> [!NOTE]
	> We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.

2. If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.

3. Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`. 

In the following example, the log file identifies a control that points to a resource file that doesn't exist. For this example, the fix would be to correct the typo in the manifest or to add the missing resource.

![Screenshot of a log file with an entry that specifies a Resource ID that is not found](http://i.imgur.com/f8bouLA.png) 

### Known issues with runtime logging

You might see messages in the log file that are confusing or that are classified incorrectly. For example:

- The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.

- If you see the message `Unexpected Add-in is missing required manifest fields	DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging. 

- Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail. 

## Clear the Office cache

If changes you've made in the manifest, such as file names of ribbon button icons, or text of add-in commands, do not seem to take effect, try clearing the Office cache on your computer. 

#### For Windows:
Delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.

#### For Mac:
Delete the contents of the folder `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`. 

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### For iOS:
Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.

## See also

- [Office Add-ins XML manifest](../develop/add-in-manifests.md)
- [Sideload Office Add-ins for testing](sideload-office-add-ins-for-testing.md)
- [Debug Office Add-ins](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
