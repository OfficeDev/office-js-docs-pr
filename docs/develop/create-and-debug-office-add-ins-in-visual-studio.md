---
title: Create and debug Office Add-ins in Visual Studio
description: ''
ms.date: 10/30/2018
---

# Create and debug Office Add-ins in Visual Studio

This article describes how to use Visual Studio 2017 to create and debug an Office Add-in for Excel, Word, PowerPoint, or Outlook. If you're using another version of Visual Studio, the procedures might vary slightly.

> [!NOTE]
> Visual Studio does not support creating Office Add-ins for OneNote or Project, but you can use the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) to create these types of add-ins.
> - To get started with an add-in for OneNote, see [Build your first OneNote add-in](../onenote/onenote-add-ins-getting-started.md).
> 
> - To get started with an add-in for Project, see [Build your first Project add-in](../project/project-add-ins-get-started.md).

## Prerequisites

- [Visual Studio 2017](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed

    > [!TIP]
    > If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](https://docs.microsoft.com/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed. If this workload is not yet installed, use the Visual Studio Installer to [install it](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio?view=vs-2017#modify-workloads).

- Office 2013 or later

    > [!TIP]
    > If you don't already have Office, you can join the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program) to get an Office 365 subscription, or you can [register for a free 1-month trial](https://products.office.com/en-US/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).

## Create the add-in project in Visual Studio

1. Open Visual Studio and from the Visual Studio menu bar, choose  **File** > **New** > **Project**.

2. In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose the type of add-in project you want to create. 

3. Name the project, and then choose **OK**.

4. If you've chosen to create a **Word Web Add-in** or an **Outlook Web Add-in**, Visual Studio creates a solution and its two projects appear in **Solution Explorer**. Proceed to the next section of this article to [explore the Visual Studio solution](#explore-the-visual-studio-solution). 

5. If you've chosen to create a **PowerPoint Web Add-in**, the **Create Office Add-in** dialog appears. Select one of the following options and choose the **Finish** button to create the Visual Studio solution. Then proceed to the next section of this article to [explore the Visual Studio solution](#explore-the-visual-studio-solution). 

    - **Add new functionalities to PowerPoint** - to create a task pane add-in

    - **Insert content into PowerPoint slides** - to create a content add-in

6. If you've chosen to create an **Excel Web Add-in**, the **Create Office Add-in** dialog appears. 

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

> [!NOTE]
> The files in the table below may be in the root folder of the web project, or in the **Home** folder, depending on the type of project template you used to create the Visual Studio project.

|**File name**|**Description**|
|:-----|:-----|
|**Home.html**<br/>(Excel, PowerPoint, Word)<br/><br/>**MessageRead.html**<br/>(Outlook)|The default HTML page of the add-in. This page appears as the first page inside of the add-in when it is activated in a document, email message, or appointment item. This file contains all of the file references that you need to get started. You can start developing your add-in by adding your HTML code to this file.|
|**Home.js**<br/>(Excel, PowerPoint, Word)<br/><br/>**MessageRead.js**<br/>(Outlook)|The JavaScript file associated with the **Home.html** page (Excel, PowerPoint, Word) or the **MessageRead.html** page (Outlook). This file should contain any code that is specific to the behavior of the **Home.html** page (Excel, PowerPoint, Word) or the **MessageRead.html** page (Outlook). This file contains some example code to get you started.|
|**Home.css**<br/>(Excel, PowerPoint, Word)<br/><br/>**MessageRead.css**<br/>(Outlook)|Defines the default styles to apply to your add-in. We recommend using the Office UI Fabric for design and styles. For more information see [Office UI Fabric in Office Add-ins](../design/office-ui-fabric.md).|

> [!NOTE]
> You don't have to use these files. Feel free to add other files to the project and use those instead. If you want another HTML file to appear as the initial page of the add-in, open the manifest editor, and then set the  **SourceLocation** property to the name of the file.

## Debug your add-in

You can use Visual Studio to debug your add-in, as described in the following sections:

- [Review the build and debug properties](#review-the-build-and-debug-properties)
- [Use an existing document to debug the add-in](#use-an-existing-document-to-debug-the-add-in)
- [Start the solution](#start-the-solution)
- [Debug the code for an Excel add-in or Word add-in](#debug-the-code-for-an-excel-add-in-or-word-add-in)
- [Debug the code for an Outlook add-in](#debug-the-code-for-an-outlook-add-in)

### Review the build and debug properties

Before you start debugging, review the properties of the add-in project to confirm that Visual Studio will open the desired host application and that other build and debug properties are set appropriately.

To view project properties, open the **Properties** window for the add-in project:

1. In  **Solution Explorer**, choose the add-in project (*not* the web application project).

2. From the menu bar, choose  **View** >  **Properties Window**.

The following table describes the properties of the project.

|**Property**|**Description**|
|:-----|:-----|
|**Start Action**|Specifies whether to debug your add-in in an Office desktop client or in an Office Online client in the specified browser.|
|**Start Document** (Excel, PowerPoint, and Word add-ins only)|Specifies what document to open when you start the project.|
|**Web Project**|Specifies the name of the web project associated with the add-in.|
|**Email Address** (Outlook add-ins only)|Specifies the email address of the user account in Exchange Server or Exchange Online that you want to test your Outlook add-in with.|
|**EWS Url** (Outlook add-ins only)|Exchange Web service URL (For example: https://www.contoso.com/ews/exchange.aspx). |
|**OWA Url** (Outlook add-ins only)|Outlook Web App URL (For example: https://www.contoso.com/owa).|
|**User name** (Outlook add-ins only)|Specifies the name of your user account in Exchange Server or Exchange Online.|
|**Project File**|Specifies the name of the file containing build, configuration, and other information about the project.|
|**Project Folder**|The location of the project file.|

### Use an existing document to debug the add-in

If you have a document that contains test data you want to use while debugging your Excel, PowerPoint, or Word add-in, Visual Studio can be configured to open that document when you start the project. To specify an existing document to use while debugging the add-in, complete the following steps.

1. In **Solution Explorer**, choose the add-in project (*not* the web application project).
    
2. From the menu bar, choose **Project** > **Add Existing Item**.
    
3. In the **Add Existing Item** dialog box, locate and select the document that you want to add.
    
4. Choose the **Add** button to add the document to your project.
    
5. In **Solution Explorer**, choose the add-in project (*not* the web application project).

6. From the menu bar, choose **View** > **Properties Window**.

7. In the **Properties** window, choose the **Start Document** list, and then select the document that you added to the project. The project is now configured to start the add-in in that document.

### Start the solution

Start the solution from the menu bar by choosing **Debug** > **Start Debugging**. Visual Studio will automatically build the solution and start Office to host your add-in.

When Visual Studio builds the project it performs the following tasks:

1. Creates a copy of the XML manifest file and adds it to  _ProjectName_\Output directory. The host application consumes this copy when you start Visual Studio and debug the add-in.
    
2. Creates a set of registry entries on your computer that enable the add-in to appear in the host application.
    
3. Builds the web application project, and then deploys it to the local IIS web server (http://localhost). 
    
Next, Visual Studio does the following:

1. Modifies the [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation?view=office-js) element of the XML manifest file by replacing the ~remoteAppUrl token with the fully qualified address of the start page (for example, http://localhost/MyAgave.html).
    
2. Starts the web application project in IIS Express.
    
3. Opens the host application. 
    
Visual Studio doesn't show validation errors in the  **OUTPUT** window when you build the project. Visual Studio reports errors and warnings in the **ERRORLIST** window as they occur. Visual Studio also reports validation errors by showing wavy underlines (known as squiggles) of different colors in the code and text editor. These marks notify you of problems that Visual Studio detected in your code. For more information, see [Code and Text Editor](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx). For more information about how to enable or disable validation, see [Options, Text Editor, JavaScript, IntelliSense](https://docs.microsoft.com/en-us/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2017).
    
To review the validation rules of the XML manifest file in your project, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).

### Debug the code for an Excel add-in or Word add-in

If you set the **Start Document** property of the add-in project to **Excel** or **Word**, Visual Studio creates a new document and the add-in appears. If you set the **Start Document** property of the add-in project to use an existing document, Visual Studio opens the document, but you have to insert the add-in manually by choosing the **Show Taskpane** button in the ribbon of the **Home** tab.

After your add-in is displayed in Excel or Word, you can debug your code by doing the following:

1. In Excel or Word, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.

    ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)

2. In the list of available add-ins, find the **Developer Add-ins** section and select the your add-in to register it.

3. In Visual Studio, set breakpoints in your code.

4. In Excel or Word, interact with your add-in.

5. As breakpoints are hit in Visual Studio, step through the code as needed.

You can change your code and review the effects of those changes in your add-in without having to close the host application and restart the project. After you save changes to your code, simply open the shortcut menu for the add-in (in Excel or Word), and then choose **Reload**.

### Debug the code for an Outlook add-in

To view the add-in in Outlook, open an email message or appointment item. 

Outlook activates the add-in for the item as long as the activation criteria are met. The add-in bar appears at the top of the Inspector window or Reading Pane, and your Outlook add-in appears as a button in the add-in bar. If your add-in has an add-in command, a button will appear in the ribbon, either in the default tab or a specified custom tab, and the add-in will not appear in the add-in bar.

To view your Outlook add-in, choose the button for your Outlook add-in. After your add-in is displayed in Outlook, you can debug your code by doing the following:

1. In Visual Studio, set breakpoints in your code.

2. In Outlook, interact with your add-in.

3. As breakpoints are hit in Visual Studio, step through the code as needed.

You can change your code and review the effects of those changes in your add-in without having to close Outlook and restart the project. After you save changes to your code, simply open the shortcut menu for the add-in (in Outlook), and then choose **Reload**.

## Next steps

After your add-in is working as desired, see [Deploy and publish your Office Add-in](../publish/publish.md) to learn about the ways you can distribute the add-in to users.
    
