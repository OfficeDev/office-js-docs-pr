
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
- ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](http://dev.office.com/reference/add-ins/shared/officeui) API.  

## Supported platforms
Add-in commands are currently supported on the following platforms:

- Office for Windows Desktop 2016 (build 16.0.6769+)
- Office for Mac (build 15.33+)
- Office Online 

More platforms are coming soon.

## Best practices

Apply the following best practices when you develop add-in commands:

- Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.
- Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.
- For the placement of your commands in the Office ribbon:
	- Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../overview/add-in-manifests.md). 
	- Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office Desktop and Office Online) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office Online).  
	- Place commands on a custom tab if you have more than six top-level commands. 
    - Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.
    - Do not add superfluous buttons to increase the real estate of your add-in.

     >**Note:**  Add-ins that take up too much space might not pass [Office Store validation](https://dev.office.com/officestore/docs/validation-policies).

- For all icons, follow the [icon design guidelines](../design/design-icons.md).
- Provide a version of your add-in that also works on hosts that do not support commands. A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a taskpane) hosts.

    ![A screenshot that shows a task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016](../../images/4f90a3cc-8cc4-4879-9a03-0bb2b6079026.png)


## Get started with add-in commands

The best way to get started using add-in commands is via **samples**. See the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.

For detailed manifest reference information, see [Define add-in commands in your manifest](http://dev.office.com/docs/add-ins/outlook/manifests/define-add-in-commands).





