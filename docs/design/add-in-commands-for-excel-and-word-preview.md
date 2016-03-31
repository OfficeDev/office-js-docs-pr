
# Add-in commands for Excel, Word and PowerPoint

Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.

**Add-in with commands running in Excel Desktop**
![Add-in commands](../../images/addincommands1.png)

**Add-in with commands running in Excel Online**
![Add-in commands](../../images/addincommands2.png)

##Commands capabilities
You can use add-in commands to make your add-ins easier to find and use. Current supported capabilities include the following.

####Extension points
- Ribbon tabs - Extend built-in tabs or create a new custom tab.
- Context menus - Extend selected context menus. 

####Control types
- Simple buttons that trigger actions.
- Menus that contain multiple buttons that trigger actions.

####Actions
- ShowTaskpane - Display one or multiple panes that load custom HTML pages inside them.
- ExecuteFunction - Load an invisible HTML page and then execute a JavaScript function within it. To show UI within your function, you can use displayDialog.  


For an overview of the feature, see the video [Add-in Commands in Office Ribbon](https://channel9.msdn.com/Events/Visual-Studio/Connect-event-2015/316).


##Supported platforms
Add-in commands are currently supported on the following platforms:

- Office Desktop 2016 (build 16.0.6769.0000 or later)
- Office Online

More platforms are coming soon.

## Get started creating add-in commands

To get started using add-in commands, see the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.

For an example of a manifest that uses commands, see the VersionOverrides section in [Office Add-ins XML manifest](../overview/add-in-manifests.md). 





