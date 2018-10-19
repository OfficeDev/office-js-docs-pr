---
title: Create and debug Office Add-ins in Visual Studio
description: ''
ms.date: 10/01/2018
---


# Create and debug Office Add-ins in Visual Studio

This article describes how to use Visual Studio to create your first Office Add-in. The steps in this article based on Visual Studio 2017. If you're using another version of Visual Studio, the procedures might vary slightly.

> [!NOTE]
> To get started with an add-in for OneNote, see [Build your first OneNote add-in](../onenote/onenote-add-ins-getting-started.md).

## Create an Office Add-in project in Visual Studio


To get started, make sure you have the [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) installed, and a version of Microsoft Office. You can join the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program), or follow these instructions to get the [latest version](../develop/install-latest-office-version.md).

1. On the Visual Studio menu bar, choose  **File** > **New** > **Project**.
2. In the list of project types under  **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose  **Add-ins**, and then select one of the add-in projects.
3. Name the project, and then choose **OK** to create the project.

In Visual Studio 2017, the following add-in project templates have additional choices after you choose **OK**:

**PowerPoint**
- You can choose to **Add new functionalities to PowerPoint** which creates a task pane add-in.
- Or you can choose to **Insert content into PowerPoint slides** which creates a content add-in.

**Excel** 
- You can choose to **Add new functionalities to Excel** which creates a task pane add-in.
- Or you can choose to **Insert content into Excel spreadsheet** which creates a content add-in.
    - If you create a content add-in, you have an additional choice of **Basic Add-in** which creates a content add-in project with minimal starter code.
    - Or you can choose a **Document Visualization Add-in** which includes starter code to visualize and bind to data.

After you complete the wizard Visual Studio creates a solution for you that contains two projects. You'll see the default Home.html page open.

|**Project**|**Description**|
|:-----|:-----|
|Add-in project|Contains only an XML manifest file, which contains all the settings that describe your add-in. These settings help the Office host determine when your add-in should be activated and where the add-in should appear. Visual Studio generates the contents of this file for you so that you can run the project and use your add-in immediately. You change these settings any time by using the Manifest editor.|
|Web application project|Contains the content pages of your add-in, including all the files and file references that you need to develop Office-aware HTML and JavaScript pages. While you develop your add-in, Visual Studio hosts the web application on your local IIS server. When you're ready to publish, you'll have to find a server to host this project. To learn more about ASP.NET web application projects, see [ASP.NET Web Projects](http://msdn.microsoft.com/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).|

## Modify your add-in settings


To modify the settings of your add-in, edit the XML manifest file of the project. In  **Solution Explorer**, expand the add-in project node, expand the folder that contains the XML manifest, and choose the XML manifest. You can point to any element in the file to view a tooltip that describes the purpose of the element. For more information about the manifest file, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).


## Develop the contents of your add-in

While the add-in project lets you modify the settings that describe your add-in, the web application provides the content that appears in the add-in. 

The web application project contains a default HTML page and JavaScript file that you can use to get started. These files contain references to other JavaScript libraries including the JavaScript API for Office. You can develop your add-in by updating these files, and adding more HTML and JavaScript files. The following table describes the default HTML and JavaScript files.

> [!NOTE]
> The files in the table below may be in the root folder of the web project, or the **Home** folder depending on the type of project template you used.

|**File**|**Description**|
|:-----|:-----|
|**Home.html**|The default HTML page of the add-in. This page appears as the first page inside of the add-in when it is activated in a document, email message, or appointment item. This file contains all of the file references that you need to get started. You can start developing your add-in by adding your HTML code to this file.|
|**Home.js**|The JavaScript file associated with the Home.html page. You can place any code that is specific to the behavior of the Home.html page in the Home.js file. The Home.js file contains some example code to get you started.|
|**Home.css**|Defines the default styles to apply to your add-in. We recommend using the Office UI Fabric for design and styles. For more information see [Office UI Fabric in Office Add-ins](../design/office-ui-fabric.md).|

> [!NOTE]
> You don't have to use these files. Feel free to add other files to the project and use those instead. If you want another HTML file to appear as the initial page of the add-in, open the manifest editor, and then set the  **SourceLocation** property to the name of the file.

## Debug your add-in

Visual Studio provides build and debug properties to assist with debugging your add-in in the Office desktop client on Windows. Visual Studio does not debug add-ins in other Office configurations, such as Office Online, or Office for Mac. For more information on debugging in other configurations, see [Debug Office Add-ins in Office Online](../testing/debug-add-ins-in-office-online.md) or [Debug Office Add-ins on iPad and Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md).

### Review the build and debug properties

Before you start the solution, verify that Visual Studio will open the host application that you want. That information appears in the property pages of the project along with several other properties that relate to building and debugging the add-in.

### To open the property pages of a project

1. In  **Solution Explorer**, choose the basic add-in project (not the Web project).    
2. On the menu bar, choose  **View** >  **Properties Window**.
    
The following table describes the properties of the project.



|**Property**|**Description**|
|:-----|:-----|
|**Start Action**|Specifies whether to debug your add-in in an Office desktop client or in an Office Online client in the specified browser.|
|**Start Document** (Content and task pane add-ins only)|Specifies what document to open when you start the project.|
|**Web Project**|Specifies the name of the web project associated with the add-in.|
|**Email Address** (Outlook add-ins only)|Specifies the email address of the user account in Exchange Server or Exchange Online that you want to test your Outlook add-in with.|
|**EWS Url** (Outlook add-ins only)|Exchange Web service URL (For example: https://www.contoso.com/ews/exchange.aspx). |
|**OWA Url** (Outlook add-ins only)|Outlook Web App URL (For example: https://www.contoso.com/owa).|
|**User name** (Outlook add-ins only)|Specifies the name of your user account in Exchange Server or Exchange Online.|
|**Project File**|Specifies the name of the file containing build, configuration, and other information about the project.|
|**Project Folder**|The location of the project file.|

### Use an existing document to debug the add-in (content and task pane add-ins only)

You can add documents to the add-in project. If you have a document that contains test data that you want to use with your add-in, Visual Studio opens that document for you when you start the project.

### To use an existing document to debug the add-in

1. In  **Solution Explorer**, choose the add-in project folder.
    
    > [!NOTE]
    > Choose the add-in project and not the web application project.

2. On the  **Project** menu, choose **Add Existing Item**.
    
3. In the  **Add Existing Item** dialog box, locate and select the document that you want to add.
    
4. Choose the  **Add** button to add the document to your project.
    
5. In  **Solution Explorer**, choose the add-in project folder.
6. On the menu bar, choose  **View** > **Properties Window**.
7. In the properties window, choose the **Start Document** list, and then choose the document that you added to the project. Now the project is configured to start your add-in in your existing document.

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
    
Visual Studio doesn't show validation errors in the  **OUTPUT** window when you build the project. Visual Studio reports errors and warnings in the **ERRORLIST** window as they occur. Visual Studio also reports validation errors by showing wavy underlines (known as squiggles) of different colors in the code and text editor. These marks notify you of problems that Visual Studio detected in your code. For more information, see [Code and Text Editor](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx). For more information about how to enable or disable validation, see: 

- [Options, Text Editor, JavaScript, IntelliSense](https://docs.microsoft.com/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2015)
    
- [How to: Set Validation Options for HTML Editing in Visual Web Developer](https://msdn.microsoft.com/library/0byxkfet(v=vs.100).aspx)
    
- [CSS, see Validation, CSS, Text Editor, Options Dialog Box](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)
    
To review the validation rules of the XML manifest file in your project, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).

### Show an add-in in Excel, or Word and step through your code

If you set the  **Start Document** property of the add-in project to Excel or Word, Visual Studio creates a new document and the add-in appears. If you set the **Start Document** property of the add-in project to use an existing document, Visual Studio opens the document, but you have to insert the add-in manually.

1. In Excel or Word, on the  **Insert** tab, choose the **My Add-ins** drop down list. Choose the list from the drop-down arrow, not the button itself which opens the **Office Add-ins** dialog.
2. Under **Developer Add-ins**, choose your add-in.

In Visual Studio, you can then set break-points and interact with your add-in and step through the code in your HTML or JavaScript files.

### Show the Outlook add-in in Outlook and step through your code

To view the add-in in Outlook, open an email message or appointment item.

Outlook activates the add-in for the item as long as the activation criteria are met. The add-in bar appears at the top of the Inspector window or Reading Pane, and your Outlook add-in appears as a button in the add-in bar. If your add-in has an add-in command, a button will appear in the ribbon, either in the default tab or a specified custom tab, and the add-in will not appear in the add-in bar.

To view your Outlook add-in, choose the button for your Outlook add-in.

In Visual Studio, you can then set break-points and interact with your add-in and step through the code in your HTML or JavaScript files.

You can also change your code and review the effects of those changes in your Outlook add-in without having to close the Office Add-in and start the project again. In Outlook, just open the shortcut menu for the Outlook add-in, and then choose  **Reload**.


### Modify code and continue to debug the add-in without having to start the project again

You can change your code and review the effects of those changes in your add-in without having to close the host application and start the project again. After you change and save your code, open the shortcut menu for the add-in, and then choose **Reload**.
    

## Next steps

- [Deploy and publish your Office Add-in](../publish/publish.md)
    
