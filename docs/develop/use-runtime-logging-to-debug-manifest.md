# Use runtime logging to debug the manifest for your Office Add-in

You can use runtime logging to debug your add-in's manifest. This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs. Runtime logging is particularly  useful for debugging add-ins that implement add-in commands.  

>**Note:** The runtime logging feature is currently available for Office 2016 desktop.

## Turn on runtime logging

>**Important**: Runtime Logging affects performance. Turn it on only when you need to debug issues with your add-in manifest.

1. Make sure that you are running Office 2016 desktop build **16.0.7019** or later. 
2. Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\`. 
3. Set the default value of the key to the full path of the file where you want the log to be written. For an example, see [EnableRuntimeLogging.zip](RuntimeLogging/EnableRuntimeLogging.zip). 

The following image shows what the registry should look like.
![Screenshot of the registry editor with a RuntimeLogging registry key](http://i.imgur.com/Sa9TyI6.png)

To turn the feature off, remove the `RuntimeLogging` key from the registry. 

## Troubleshoot issues with your manifest

To use runtime logging to troubleshoot issues with add-in commands:
 
1. [Sideload your add-in](testing/sideload-office-add-ins-for-testing.md) for testing. 

	>Note: We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.
2. If you don't see your buttons on the ribbon and nothing appears on the add-ins dialog box, open the log file.
3. Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`. 

In the following example, the log file identifies a control that points to a resource file that doesn't exist. For this example, the fix would be to correct the typo in the manifest or to add the missing resource.

![Screenshot of a log file with an entry that specifies a Resource ID that is not found](http://i.imgur.com/f8bouLA.png) 

##Known issues with runtime logging
You might see messages in the log file that are confusing or that are classified incorrectly. For example:

- The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.
- The message `Unexpected Add-in is missing required manifest fields	DisplayName` doesn't contain the SolutionId of the add-in. This error is most likely not related to the add-in you are debugging. 
- Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail. 

##Additional resources

- [Sideload Office Add-ins for testing](testing/sideload-office-add-ins-for-testing.md)
- [Debug Office Add-ins](testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md)
