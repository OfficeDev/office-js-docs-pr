---
title: Create and debug Office Add-ins in Visual Studio
description: 'Use Visual Studio to create and debug Office Add-ins in the Office desktop client on Windows'
ms.date: 10/11/2019
localization_priority: Priority
---

# Create and debug Office Add-ins in Visual Studio

This article describes how to use Visual Studio 2019 to create an Office Add-in for Excel, Word, PowerPoint, or Outlook and debug the add-in in the Office desktop client on Windows. If you're using another version of Visual Studio, the procedures might vary slightly.

> [!NOTE]
> Visual Studio does not support creating Office Add-ins for OneNote or Project, but you can use the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) to create these types of add-ins.
> - To get started with an add-in for OneNote, see [Build your first OneNote add-in](../quickstarts/onenote-quickstart.md).
>
> - To get started with an add-in for Project, see [Build your first Project add-in](../quickstarts/project-quickstart.md).

## Prerequisites

- [Visual Studio 2019](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed

    > [!TIP]
    > If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed. If this workload is not yet installed, use the Visual Studio Installer to [install it](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-workloads).

- Office 2013 or later

    > [!TIP]
    > If you don't already have Office, you can join the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program) to get an Office 365 subscription, or you can [register for a free 1-month trial](https://products.office.com/en-US/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).

## Create the add-in project in Visual Studio

Start by completing these three steps, and then complete the steps in the following section that corresponds to the type of add-in you're creating. 

1. Open Visual Studio and from the Visual Studio menu bar, choose  **Create a new project**.

2. Using the search box, enter **Add-ins**. Then choose the type of add-in project you want to create.

3. Name the project, and then choose **OK**.

### Word Web Add-in or Outlook Web Add-in

If you've chosen to create a **Word Web Add-in** or an **Outlook Web Add-in**, Visual Studio creates a solution and its two projects appear in **Solution Explorer**. Next, you can [explore the Visual Studio solution](#explore-the-visual-studio-solution).

### PowerPoint Web Add-in

If you've chosen to create a **PowerPoint Web Add-in**, the **Create Office Add-in** dialog appears. 

- To create a task pane add-in, select **Add new functionalities to PowerPoint** and then choose the **Finish** button to create the Visual Studio solution.

- To create a content add-in, select **Insert content into PowerPoint slides** and then choose the **Finish** button to create the Visual Studio solution.

Next, you can [explore the Visual Studio solution](#explore-the-visual-studio-solution).

### Excel Web Add-in

If you've chosen to create an **Excel Web Add-in**, the **Create Office Add-in** dialog appears. 

- To create a task pane add-in, select **Add new functionalities to Excel** and then choose the **Finish** button to create the Visual Studio solution.

- To create a content add-in, select **Insert content into Excel spreadsheets**, choose the **Next** button, select one of the following options, and then choose the **Finish** button to create the Visual Studio solution:

    - **Basic Add-in** - to create a content add-in project with minimal starter code

    - **Document Visualization Add-in** - to create a content add-in project with starter code to visualize and bind to data  

### Explore the Visual Studio solution

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

## Modify your add-in settings

To modify the settings of your add-in, edit the XML manifest file in the add-in project. In  **Solution Explorer**, expand the add-in project node, expand the folder that contains the XML manifest, and choose the XML manifest. You can point to any element in the file to view a tooltip that describes the purpose of the element. For more information about the manifest file, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).

## Develop the contents of your add-in

While the add-in project lets you modify the settings that describe your add-in, the web application provides the content that appears in the add-in. 

The web application project contains a default HTML file, JavaScript file, and CSS file that you can use to get started. Some of these files contain references to other JavaScript libraries including the JavaScript API for Office. You can develop your add-in by updating these files and/or adding more HTML and JavaScript files. The following table describes the default files that the web application project contains when the Visual Studio solution is created.

|**File name**|**Description**|
|:-----|:-----|
|**Home.html**<br/>(Excel, PowerPoint, Word)<br/><br/>**MessageRead.html**<br/>(Outlook)|The default HTML page of the add-in. This page appears as the first page inside of the add-in when it is activated in a document, email message, or appointment item. This file contains all of the file references that you need to get started. You can start developing your add-in by adding your HTML code to this file.|
|**Home.js**<br/>(Excel, PowerPoint, Word)<br/><br/>**MessageRead.js**<br/>(Outlook)|The JavaScript file associated with the **Home.html** page (Excel, PowerPoint, Word) or the **MessageRead.html** page (Outlook). This file should contain any code that is specific to the behavior of the **Home.html** page (Excel, PowerPoint, Word) or the **MessageRead.html** page (Outlook). This file contains some example code to get you started.|
|**Home.css**<br/>(Excel, PowerPoint, Word)<br/><br/>**MessageRead.css**<br/>(Outlook)|Defines the default styles to apply to your add-in. We recommend using the Office UI Fabric for design and styles. For more information see [Office UI Fabric in Office Add-ins](../design/office-ui-fabric.md).|

> [!NOTE]
> You don't have to use these files. Feel free to add other files to the project and use those instead. If you want another HTML file to appear as the initial page of the add-in, open the manifest editor, and then set the  **SourceLocation** property to the name of the file.

## Debug your add-in

You can use Visual Studio to debug your add-in in the Office desktop client on Windows, as described in the following sections:

- [Review the build and debug properties](#review-the-build-and-debug-properties)
- [Use an existing document to debug the add-in](#use-an-existing-document-to-debug-the-add-in)
- [Start the project](#start-the-project)
- [Debug the code for an Excel, PowerPoint, or Word add-in](#debug-the-code-for-an-excel-powerpoint-or-word-add-in)
- [Debug the code for an Outlook add-in](#debug-the-code-for-an-outlook-add-in)

> [!NOTE]
> You cannot use Visual Studio to debug add-ins in Office on the web or Mac. For information about debugging on these platforms, see [Debug Office Add-ins in Office on the web](../testing/debug-add-ins-in-office-online.md) or [Debug Office Add-ins on iPad and Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)

### Review the build and debug properties

Before you start debugging, review the properties of each project to confirm that Visual Studio will open the desired host application and that other build and debug properties are set appropriately.

#### Add-in project properties

Open the **Properties** window for the add-in project to review project properties:

1. In  **Solution Explorer**, choose the add-in project (*not* the web application project).

2. From the menu bar, choose  **View** >  **Properties Window**.

The following table describes the properties of the add-in project.

|**Property**|**Description**|
|:-----|:-----|
|**Start Action**|Specifies the debug mode for your add-in. Currently only **Office Desktop Client** mode is supported for Office Add-in projects.|
|**Start Document**<br/>(Excel, PowerPoint, and Word add-ins only)|Specifies what document to open when you start the project.|
|**Web Project**|Specifies the name of the web project associated with the add-in.|
|**Email Address**<br/>(Outlook add-ins only)|Specifies the email address of the user account in Exchange Server or Exchange Online that you want to use to test your Outlook add-in.|
|**EWS Url**<br/>(Outlook add-ins only)|Exchange Web service URL (For example: `https://www.contoso.com/ews/exchange.aspx`). |
|**OWA Url**<br/>(Outlook add-ins only)|Outlook on the web URL (For example: `https://www.contoso.com/owa`).|
|**Use multi-factor auth**<br/>(Outlook add-ins only)|Boolean value that indicates whether multi-factor authentication should be used.|
|**User Name**<br/>(Outlook add-ins only)|Specifies the name of the user account in Exchange Server or Exchange Online that you want to use to test your Outlook add-in.|
|**Project File**|Specifies the name of the file containing build, configuration, and other information about the project.|
|**Project Folder**|The location of the project file.|

> [!NOTE]
> For an Outlook add-in, you may choose to specify values for one or more of the *Outlook add-in only* properties in the **Properties** window, but doing so is not required.

#### Web application project properties

Open the **Properties** window for the web application project to review project properties:

1. In  **Solution Explorer**, choose the web application project.

2. From the menu bar, choose  **View** >  **Properties Window**.

The following table describes the properties of the web application project that are most relevant to Office Add-in projects.

|**Property**|**Description**|
|:-----|:-----|
|**SSL Enabled**|Specifies whether SSL is enabled on the site. This property should be set to **True** for Office Add-in projects.|
|**SSL URL**|Specifies the secure HTTPS URL for the site. Read-only.|
|**URL**|Specifies the HTTP URL for the site. Read-only.|
|**Project File**|Specifies the name of the file containing build, configuration, and other information about the project.|
|**Project Folder**|Specifies the location of the project file. Read-only. The manifest file that Visual Studio generates at runtime is written to the `bin\Debug\OfficeAppManifests` folder in this location.|

### Use an existing document to debug the add-in

If you have a document that contains test data you want to use while debugging your Excel, PowerPoint, or Word add-in, Visual Studio can be configured to open that document when you start the project. To specify an existing document to use while debugging the add-in, complete the following steps.

1. In **Solution Explorer**, choose the add-in project (*not* the web application project).

2. From the menu bar, choose **Project** > **Add Existing Item**.

3. In the **Add Existing Item** dialog box, locate and select the document that you want to add.

4. Choose the **Add** button to add the document to your project.

5. In **Solution Explorer**, choose the add-in project (*not* the web application project).

6. From the menu bar, choose **View** > **Properties Window**.

7. In the **Properties** window, choose the **Start Document** list, and then select the document that you added to the project. The project is now configured to start the add-in in that document.

### Start the project

Start the project by choosing **Debug** > **Start Debugging** from the menu bar. Visual Studio will automatically build the solution and start Office to host your add-in.

> [!NOTE]
> When you start an Outlook add-in project, you'll be prompted for login credentials. If you're asked to log in repeatedly or if you receive an error that you are unauthorized, then Basic Auth may be disabled for accounts on your Office 365 tenant. In this case, try using a Microsoft account instead. You may also need to set the property "Use multi-factor auth" to True in the Outlook Web Add-in project properties dialog.

When Visual Studio builds the project it performs the following tasks:

1. Creates a copy of the XML manifest file and adds it to  `_ProjectName_\bin\Debug\OfficeAppManifests` directory. The host application consumes this copy when you start Visual Studio and debug the add-in.

2. Creates a set of registry entries on your computer that enable the add-in to appear in the host application.

3. Builds the web application project, and then deploys it to the local IIS web server (https://localhost).

4. If this is the first add-in project that you have deployed to local IIS web server, you may be prompted to install a Self-Signed Certificate to the current user's Trusted Root Certificate store. This is required for IIS Express to display the content of your add-in correctly.


> [!NOTE]
> The latest version of Office may use a newer web control to display the add-in contents when running on Windows 10. If this is the case, Visual Studio may prompt you to add a local network loopback exemption. This is required for the web control, in the Office host application, to be able to access the website deployed to the local IIS web server. You can also change this setting anytime in Visual Studio under **Tools** > **Options** > **Office Tools (Web)** > **Web Add-In Debugging**.


Next, Visual Studio does the following:

1. Modifies the [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) element of the XML manifest file by replacing the `~remoteAppUrl` token with the fully qualified address of the start page (for example, `https://localhost:44302/Home.html`).

2. Starts the web application project in IIS Express.

3. Opens the host application.

Visual Studio doesn't show validation errors in the  **OUTPUT** window when you build the project. Visual Studio reports errors and warnings in the **ERRORLIST** window as they occur. Visual Studio also reports validation errors by showing wavy underlines (known as squiggles) of different colors in the code and text editor. These marks notify you of problems that Visual Studio detected in your code. For more information, see [Code and Text Editor](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx). For more information about how to enable or disable validation, see [Options, Text Editor, JavaScript, IntelliSense](/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2019).

To review the validation rules of the XML manifest file in your project, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).

### Debug the code for an Excel, PowerPoint, or Word add-in

If your add-in isn't visible within the document that's displayed in the host application (Excel, PowerPoint, or Word) after you've [started the project](#start-the-project), manually launch the add-in in the host application. For example, launch your task pane add-in by choosing the **Show Taskpane** button in the ribbon of the **Home** tab. After your add-in is displayed in Excel, PowerPoint, or Word, you can debug your code by doing the following:

1. In Excel, PowerPoint, or Word, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.

    ![Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)

2. In the list of available add-ins, find the **Developer Add-ins** section and select the your add-in to register it.

3. In Visual Studio, set breakpoints in your code.

4. In Excel, PowerPoint, or Word, interact with your add-in.

5. As breakpoints are hit in Visual Studio, step through the code as needed.

You can change your code and review the effects of those changes in your add-in without having to close the host application and restart the project. After you save changes to your code, simply reload the add-in in the host application. For example, reload a task pane add-in by choosing the top-right corner of the task pane to activate the [personality menu](../design/task-pane-add-ins.md#personality-menu) and then choose **Reload**.

### Debug the code for an Outlook add-in

After you've [started the project](#start-the-project) and Visual Studio launches Outlook to host your add-in, open an email message or appointment item. 

Outlook activates the add-in for the item as long as the activation criteria are met. The add-in bar appears at the top of the Inspector window or Reading Pane, and your Outlook add-in appears as a button in the add-in bar. If your add-in has an add-in command, a button will appear in the ribbon, either in the default tab or a specified custom tab, and the add-in will not appear in the add-in bar.

To view your Outlook add-in, choose the button for your Outlook add-in. After your add-in is displayed in Outlook, you can debug your code by doing the following:

1. In Visual Studio, set breakpoints in your code.

2. In Outlook, interact with your add-in.

3. As breakpoints are hit in Visual Studio, step through the code as needed.

You can change your code and review the effects of those changes in your add-in without having to close Outlook and restart the project. After you save changes to your code, simply open the shortcut menu for the add-in (in Outlook), and then choose **Reload**.

## Next steps

After your add-in is working as desired, see [Deploy and publish your Office Add-in](../publish/publish.md) to learn about the ways you can distribute the add-in to users.
