
# Add-in commands for Excel, Word, and PowerPoint

Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.

For an overview of the feature, see the video [Add-in Commands in the Office Ribbon](https://channel9.msdn.com/events/Build/2016/P551).

>**Note:** SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [centralized deployment](https://support.office.com/en-ie/article/Deploy-Office-Add-ins-in-the-Office-365-new-Admin-Center-737e8c86-be63-44d7-bf02-492fa7cd9c3f?ui=en-US&rs=en-IE&ad=IE) or the [Office Store](https://msdn.microsoft.com/en-us/library/jj220033.aspx), or use [sideloading](https://dev.office.com/docs/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) to deploy your add-in command for testing. 

**Add-in with commands running in Excel Desktop**
![Add-in commands](../../images/addincommands1.png)

**Add-in with commands running in Excel Online**
![Add-in commands](../../images/addincommands2.png)

## Command capabilities
The following command capabilities are currently supported.

> **Note:** Content add-ins do not currently support add-in commands.

**Extension points**

- Ribbon tabs - Extend built-in tabs or create a new custom tab.
- Context menus - Extend selected context menus. 

**Control types**

- Simple buttons - trigger specific actions.
- Menus - simple menu dropdown with buttons that trigger actions.

**Actions**

- ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.
- ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (e.g. errors, progress, additional input) you can use the [displayDialog](http://dev.office.com/reference/add-ins/shared/officeui) API.  

## Supported platforms
Add-in commands are currently supported on the following platforms:

- Office for Windows Desktop 2016 (build 16.0.6769+)
- Office for Mac (build 15.33+)
- Office Online 

More platforms are coming soon.

## Get started with add-in commands

The best way to get started using add-in commands is via **samples**. See the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.

For detailed manifest rerence information, see [Define add-in commands in your manifest](http://dev.office.com/docs/add-ins/outlook/manifests/define-add-in-commands).





