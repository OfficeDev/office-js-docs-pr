# Use Runtime Logging to debug Add-in Commands

Office 16 Desktop clients have a new feature available to log useful information. Among other things, this tool can help you diagnose errors in your add-in manifest which comes particularly handy if you are creating manifests with add-in commands. 

Full documentation for the feature is on the way but in the meantime here is how you can use it to debug issues when parsing manifests with add-in commands.

##Turn On Runtime Logging

**Important**: Runtime Logging has a **performance hit**. Only turn it on when you need to debug issues with your add-ins

1. Ensure that you have a build that supports Runtime Logging. You need **Office 16 Desktop** clients with build equal or greater than **16.0.7019**
2. Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\` 
3. Set the key's default value to the full path of the file where you want the log to be written. See [sample registry key](RuntimeLogging/EnableRuntimeLogging.zip) (unzip)

Your registry should look like this:
![](http://i.imgur.com/Sa9TyI6.png)

If you need to turn the feature off, simply remove the key from the registry. 

##Diagnose issues with commands
Runtime Logging is useful to detect **issues with your manifest** that are hard to catch, for example, mismatch between resource Ids, invalid lengths, that are not caught by XSD schema validation. 

Here are the steps to try things out:
 
1. Follow the instructions on the [Readme](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/README.md) to sideload your add-in. 
2. If you don't see your Ribbon buttons project and nothing appears on the add-ins dialog, check the logs
3. Search for the id of your add-in, which your define in your manifest, to find messages belonging to that add-in. Logs report this id as `SolutionId` It is recommended that you only side-load one add-in at the time to avoid seeing too many messages that don't belong to your add-in. 

In the example below, RuntimeLogging helped identify a control that is pointing to a non-existent resource file. The fix is to correct the typo (if one exists) or to actually add the missing resource.

![](http://i.imgur.com/f8bouLA.png) 

##Known issues with logging
Runtime Logging still has known bugs. You may see several messages that are confusing or inappropriately classified. For example:

- The messages `Medium	Current host not in add-in's host list` followed by `Unexpected	Parsed manifest targeting different host` are incorrectly classified. They are not errors, you can safely ignore them.
- The message `Unexpected	Add-in is missing required manifest fields	DisplayName` doesn't contain the SolutionId of the offending add-in. However, most likely this is NOT related to the add-in you are debugging. 
- Any `Monitorable` messages are expected errors from a system point of view. Sometimes they may indicate an issue with your manifest (e.g. a misspelled element that was skipped but didn't cause the manifest to fail). 

