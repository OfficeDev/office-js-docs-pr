
# Add-in commands for Excel, Word and PowerPoint

Add-in commands are UI elements that extend the Office User Interface and start actions in your add-in. You can add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.

**Add-in with commands running in Excel Desktop**
![Add-in commands](../../images/addincommands1.png)

**Add-in with commands running in Excel Online**
![Add-in commands](../../images/addincommands2.png)

##Commands capabilities
You can use add-in commands to make your add-ins easier to find and use. Current capabilities supported include:

####Extension Points
- Ribbon tabs. You can extend built-in tabs or create new custom tab
- Context menu. You can extent selected context menus. 

####Control Types
- Simple button. You can add simple buttons that trigger actions.
- Menu button. You can add menus with multiple buttons in them that trigger actions.

####Actions
- ShowTaskpane. You can display one or multiple panes that load custom HTML pages inside them.
- ExecuteFunction. You can load an invisible HTML page and then execute a JavaScript function within it. If you need to present UI within your function you can use the displayDialog API to do so.  


Watch this video for an overview of the feature: [Add-in Commands in Office Ribbon](https://channel9.msdn.com/Events/Visual-Studio/Connect-event-2015/316).


##Supported platforms
- Office Desktop 2016 (build 16.0.6769.0000 or higher)
- Office Online
- More platforms soon

## Get started creating add-in commands

 - To get started using add-in commands follow these [samples and documentation](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/)
 - For a detailed reference of the manifest that uses commands see the VersionOverrides section of [Office Add-ins XML manifest](../overview/add-in-manifests.md) 





