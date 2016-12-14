# Validate and troubleshoot issues with your manifest

Use these methods to validate and troubleshoot issues in your manifest. 

- [Validate the Office Add-ins manifest against the XML schema](validate-the-office-add-ins-manifest-against-the-xml-schema)
- [Use runtime logging to debug the manifest for your Office Add-in](use-runtime-logging-to-debug-the-manifest-for-your-office-add-in)

## Validate the Office Add-ins manifest against the XML schema

To help to make sure that the manifest file that describes your Office Add-in is correct and complete, validate it against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs/tree/master/docs/overview/schemas) files. 
You can use an XML schema validation tool or [Visual Studio](../get-started/create-and-debug-office-add-ins-in-visual-studio.md) to validate the manifest. 

To use Visual Studio, go to Build > Publish, and choose **Perform Validation check**.

To use a command-line XML schema validation tool to validate your manifest:

1.	Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already. 
2.	Run the following command. Replace XSD_FILE with the path to the manifest XSD file and XML_FILE with the path to the manifest XML file.

	xmllint --noout --schema XSD_FILE XML_FILE

## Use runtime logging to debug the manifest for your Office Add-in

You can use runtime logging to debug your add-in's manifest. This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs. Runtime logging is particularly  useful for debugging add-ins that implement add-in commands.  

>**Note:** The runtime logging feature is currently available for Office 2016 desktop.

### Turn on runtime logging

>**Important**: Runtime Logging affects performance. Turn it on only when you need to debug issues with your add-in manifest.

1. Make sure that you are running Office 2016 desktop build **16.0.7019** or later. 
2. Add the `RuntimeLogging` registry key under 'HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\'. 
3. Set the default value of the key to the full path of the file where you want the log to be written. For an example, see [EnableRuntimeLogging.zip](RuntimeLogging/EnableRuntimeLogging.zip). 

 > **Note:** The directory in which the log file will be written must already exist and you must have write permissions to it. 
 
The following image shows what the registry should look like.
![Screenshot of the registry editor with a RuntimeLogging registry key](http://i.imgur.com/Sa9TyI6.png)

To turn the feature off, remove the `RuntimeLogging` key from the registry. 

### Troubleshoot issues with your manifest

To use runtime logging to troubleshoot issues loading an add-in:
 
1. [Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing. 

	>Note: We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.
2. If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.
3. Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`. 

In the following example, the log file identifies a control that points to a resource file that doesn't exist. For this example, the fix would be to correct the typo in the manifest or to add the missing resource.

![Screenshot of a log file with an entry that specifies a Resource ID that is not found](http://i.imgur.com/f8bouLA.png) 

### Known issues with runtime logging

You might see messages in the log file that are confusing or that are classified incorrectly. For example:

- The message `Medium	Current host not in add-in's host list` followed by `Unexpected	Parsed manifest targeting different host` is incorrectly classified as an error.
- If you see the message `Unexpected	Add-in is missing required manifest fields	DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging. 
- Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail. 

## Additional resources

- [Office Add-ins XML manifest](../overview/add-in-manifests.md)
- [Sideload Office Add-ins for testing](sideload-office-add-ins-for-testing.md)
- [Debug Office Add-ins](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

