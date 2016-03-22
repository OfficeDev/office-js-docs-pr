
# Add-in commands for Excel, Word and PowerPoint (Preview)

Add-in commands are UI elements that extend the default Office UI, and start actions in your add-in. You can add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.

You can use add-in commands to:


- Add buttons or a drop-down list of buttons to the ribbon.

- Add individual menu items, with an optional submenu, to the context menu.

- Display your add-in, which displays a UI for the user to interact with.

- Run JavaScript code, which normally runs without displaying a UI.

- Make your add-ins easier to find by providing multiple locations for users to open your add-in. You can start your add-in from buttons, context menus, or by using  **Insert** > **My Add-ins**.

- Open different pages in your add-in. Previously, when an add-in opened, it would always start on the same page. To transition to other pages in your add-in, your code would have to manage the UI transitions. With add-in commands, you can open any page in your add-in.

- In Excel 2016 or Word 2016, display more than one task pane add-in at a time.

For an overview of add-in commands in Excel and Word, see [Add-in Commands in Office Ribbon (Public preview)](https://channel9.msdn.com/Events/Visual-Studio/Connect-event-2015/316).


## Get started creating add-in commands

 **To get started using add-in commands**


1. Download the [Office Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Command-Sample).

2. Create your add-in command by updating your manifest as described in [Create add-in commands in your manifest for Excel, Word, and PowerPoint (Preview)](../../docs/design/create-add-in-commands-in-your-manifest-preview.md). For best practice information, see [Best practices for developing Office Add-ins](../../docs/design/add-in-development-best-practices.md).


## Install and test your add-in

After you update your manifest file, install your add-in and test your add-in commands. You can test your add-in commands on:


- Excel Online

- Office on a Windows desktop


### Install your add-in on Excel Online



1. Open [Microsoft Office Online](https://office.live.com/).

2. In  **Get started with the online apps now**, choose  **Excel** > **New blank workbook** > **Insert**.

3. In  **Add-ins**, choose  **Office Add-ins**.

4. On  **Office Add-ins**, choose  **Upload My Add-in** > **More** ( **...**).

5. Choose the manifest file on your computer, and then choose  **Upload Add-in**.

6. Verify that your add-in commands appear on either the ribbon or the context menu.


### Known issues with add-in commands in Excel Online



- When you close Excel Online, the add-in and its commands are removed. When next you open a document in Excel Online, you must perform the steps in [Install your add-in on Excel Online](#install-your-add-in-on-excel-online) again.

- To use add-in commands in your add-in, upload your add-in by using  **Insert** > **Office Add-ins** > **Upload My Add-in**. Add-ins submitted to the Office Store or to an internal organization add-in catalog in SharePoint won't display add-in commands.

- Only Excel Online supports add-in commands. Support for Word Online and PowerPoint Online will be added in the future.

- If you choose  **INSERT** > **Office Add-ins** > **MY ORGANIZATION** and are redirected to the Office Store, try signing in to Excel Online, and then choose **MY ORGANIZATION** again.

- When the task pane doesn't update after choosing different add-in commands, try refreshing the page in your browser.


### Install your add-in on Excel 2016 or Word 2016 on a Windows desktop


Before you install your add-in, verify that you are running Excel 2016 or Word 2016, version 16.0.6326, at a minimum.


1. Open  **Excel 2016** > **Blank workbook** > **File** > **Account**.

2. In  **Office Updates**, make sure the version number is  **16.0.6326**, at a minimum.

To get the latest version of Office that includes developer features in Preview, do one of the following based on your subscription:


- If you're an Office 365 Home, Office 365 Personal, or Office 365 University subscriber, install the [Office Insider build](https://products.office.com/en-us/office-insider).

    or

- If you're a commercial Office 365 subscriber, opt-in for [First Release](https://support.office.com/en-us/article/Office-365-release-options-3B3ADFA4-1777-4FF0-B606-FB8732101F47?ui=en-US&amp;rs=en-001&amp;ad=US), and then perform the following steps.


1. Download and run the [Office 2016 Deployment Tool](http://www.microsoft.com/en-us/download/details.aspx?id=49117).

2. Replace  **configuration.xml** with the [First Release Configuration File](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).

3. Open a command prompt with elevated privileges (Run as administrator), and run the following command: **setup.exe /configuration configuration.xml**

4. Verify that the version number of Office is  **16.0.6326.0000** or higher for Excel or Word, and **16.0.6568.2025** or higher for PowerPoint.

Before performing the following steps, make sure that your add-in runs without generating errors. For example, when using Visual Studio, press F5 to verify that your add-in runs without generating errors. If your add-in generates errors after performing the following steps, you'll be able to identify the source of the error easier.

 **To install your add-in on Office on a Windows desktop**


1. Download and run the registry key from [Office Add-in Commands Samples](https://github.com/OfficeDev/Office-Add-in-Command-Sample) to enable add-in commands in Office.

2. Publish your add-in's manifest file on a network file share. For more information, see [Create a network shared folder catalog for task pane and content add-ins](../publish/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).

3. Close and reopen the Office host application (Excel 2016 or Word 2016).

4. Choose  **Insert** > **My Add-ins** > **SHARED FOLDER**.

5. Choose your add-in, and then choose  **OK**.

6. Verify that your add-in commands display in Office. If you don't see your add-in commands in the Office UI, choose  **Insert** > **My Add-ins** > **Refresh**.


### Known issues with add-in commands in Office on Windows desktop



- If you're using a Click-to-Run version of Office, such as Office 365 ProPlus, make sure that the updates finish installing before you install your add-in.

- Add-in commands are supported in Excel 2016, Word 2016, and PowerPoint 2016 only.

- Add-ins submitted to the Office Store or to an internal organization add-in catalog in SharePoint won't display add-in commands. You can only load your manifest file from the file share.

- Updates to add-in commands don't render in Office unless you refresh. If any of the following occur, choose  **Insert** > **My Add-ins** > **Refresh**:

  - You want to remove non-working or duplicate buttons from the ribbon. You might also try clearing your cache by deleting all the files in  **%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\AppCommands**, and then choose **Insert** > **My Add-ins** > **Refresh**.

  - Your task pane add-in isn't loading.

  - You need to update your add-in commands. You might need to refresh twice for your add-in commands to display on the ribbon.

- If you want to change your icon files after installing your add-in, you can perform one of the following steps:

  - Go to  **%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\icons** and delete all the files in that folder.

    or

  - Change the icon's file name in the  **Resources** element in the manifest, and republish your add-in.

- Add-in commands are always available in the Office UI, even when no document is displayed.


## Additional resources



- [Create add-in commands in your manifest for Excel, Word, and PowerPoint (Preview)](../../docs/design/create-add-in-commands-in-your-manifest-preview.md)

- [Best practices for developing Office Add-ins](../../docs/design/add-in-development-best-practices.md)

