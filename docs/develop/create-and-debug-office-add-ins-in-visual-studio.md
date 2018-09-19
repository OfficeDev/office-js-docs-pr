---
title: Create and debug Office Add-ins in Visual Studio
description: ''
ms.date: 03/14/2018
---


# Create and debug Office Add-ins in Visual Studio

This article describes how to use Visual Studio to create your first Office Add-in. The steps in this article based on Visual Studio 2015. If you're using another version of Visual Studio, the procedures might vary slightly.

> [!NOTE]
> To get started with an add-in for OneNote, see [Build your first OneNote add-in](../onenote/onenote-add-ins-getting-started.md).

## Create an Office Add-in project in Visual Studio


To get started, make sure you have the [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) installed, and a version of Microsoft Office. You can join the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program), or follow these instructions to get the [latest version](../develop/install-latest-office-version.md).


1. On the Visual Studio menu bar, choose  **File** > **New** > **Project**.
    
2. In the list of project types under  **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose  **Web Add-ins**, and then select one of the add-in projects.  
    
3. Name the project, and then choose  **OK** to create the project.
    
4. Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The default Home.html page opens in Visual Studio.
    
In Visual Studio 2015, some of the add-in project templates have been updated to reflect additional functionality:


- Content add-ins can appear in the body of Access and PowerPoint documents, in addition to Excel spreadsheets. You can also choose the Basic Project option to create a basic content add-in project with minimal starter code, or the Document Visualization Project option (for Access and Excel only) to create a more full-featured content add-in that includes starter code to visualize and bind to data.
    
- Outlook add-ins include options not just for including your add-in in email messages or appointments, but also for specifying whether the add-in is available when an email message or appointment is being composed as well as read.
    

> [!NOTE]
> In Visual Studio most options are understandable from their descriptions except for the  **Email Message** checkbox. Use that checkbox if you want to create an Outlook add-in that appears not just with mail items, but also with meeting requests, responses, and cancellations.

When you've completed the wizard, Visual Studio creates a solution for you that contains two projects.



|**Project**|**Description**|
|:-----|:-----|
|Add-in project|Contains only an XML manifest file, which contains all the settings that describe your add-in. These settings help the Office host determine when your add-in should be activated and where the add-in should appear. Visual Studio generates the contents of this file for you so that you can run the project and use your add-in immediately. You change these settings any time by using the Manifest editor.|
|Web application project|Contains the content pages of your add-in, including all the files and file references that you need to develop Office-aware HTML and JavaScript pages. While you develop your add-in, Visual Studio hosts the web application on your local IIS server. When you're ready to publish, you'll have to find a server to host this project.To learn more about ASP.NET web application projects, see [ASP.NET Web Projects](http://msdn.microsoft.com/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).|

## Modify your add-in settings


To modify the settings of your add-in, edit the XML manifest file of the project. In  **Solution Explorer**, expand the add-in project node, expand the folder that contains the XML manifest, and choose the XML manifest. You can point to any element in the file to view a tooltip that describes the purpose of the element. For more information about the manfiest file, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).


## Develop the contents of your add-in


While the add-in project lets you modify the settings that describe your add-in, the web application provides the content that appears in the add-in. 

The web application project contains a default HTML page and JavaScript file that you can use to get started. The project also contains a JavaScript file that is common to all pages that you add to your project. These files are convenient because they contain references to other JavaScript libraries including the JavaScript API for Office. 

As your add-in becomes more sophisticated, you can add more HTML and JavaScript files. You can use the contents of the default HTML and JavaScript files as examples of the types of references you might want to add to other pages in your project to make them work with your add-in. The following table describes default HTML and JavaScript files.



|**File**|**Description**|
|:-----|:-----|
|**Home.html**|Located in the  **Home** folder of the project, this is default HTML page of the add-in. This page appears as the first page inside of the add-in when it is activated in a document, email message or appointment item. This file is convenient because it contains all of the file references that you need to get started. When you are ready to create your first add-in, just add your HTML code to this file.|
|**Home.js**|Located in the  **Home** folder of the project, this is the JavaScript file associated with the Home.js page. You can place any code that is specific to the behavior of the Home.html page in the Home.js file. The Home.js file contains some example code to get you started.|
|**App.js**|Located in the  **Add-in** folder of the project, this is the default JavaScript file of the entire add-in. You can place code that is common to the behavior of multiple pages of your add-in in the App.js file. The App.js file contains some example code to get you started.|

> [!NOTE]
> You don't have to use these files. Feel free to add other files to the project and use those instead. If you want another HTML file to appear as the initial page of the add-in, open the manifest editor, and then point the  **SourceLocation** property to the name of the file.


## Debug your add-in


When you are ready to start your add-in, review build and debug related properties, and then start the solution.


### Review the build and debug properties

Before you start the solution, verify that Visual Studio will open the host application that you want. That information appears in the property pages of the project along with several other properties that relate to building and debugging the add-in.


### To open the property pages of a project


1. In  **Solution Explorer**, choose the basic add-in project (not the Web project).
    
2. On the menu bar, choose  **View**,  **Properties Window**.
    
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
    
5. In  **Solution Explorer**, open the shortcut menu for the project, and then choose  **Properties**.
    
    The property pages for the project appear.
    
6. In the  **Start Document** list, choose the document that you added to the project, and then choose the **OK** button to close the property pages.
    

### Start the solution


Visual Studio will automatically build the solution when you start it. You can start the solution from the  **Menu** bar by choosing **Debug**,  **Start**. 


> [!NOTE]
> If script debugging isn't enabled in Internet Explorer, you won't be able to start the debugger in Visual Studio. You can enable script debugging by opening the  **Internet Options** dialog box, choosing the **Advanced** tab, and then clearing the **Disable Script Debugging (Internet Explorer)** and **Disable Script Debugging (Other)** check boxes.

Visual Studio builds the project and does the following:


1. Creates a copy of the XML manifest file and adds it to  _ProjectName_\Output directory. The host application consumes this copy when you start Visual Studio and debug the add-in.
    
2. Creates a set of registry entries on your computer that enable the add-in to appear in the host application.
    
3. Builds the web application project, and then deploys it to the local IIS web server (http://localhost). 
    
Next, Visual Studio does the following:


1. Modifies the [SourceLocation](https://docs.microsoft.com/javascript/office/manifest/sourcelocation?view=office-js) element of the XML manifest file by replacing the ~remoteAppUrl token with the fully qualified address of the start page (for example, http://localhost/MyAgave.html).
    
2. Starts the web application project in IIS Express.
    
3. Opens the host application. 
    
Visual Studio doesn't show validation errors in the  **OUTPUT** window when you build the project. Visual Studio reports errors and warnings in the **ERRORLIST** window as they occur. Visual Studio also reports validation errors by showing wavy underlines (known as squiggles) of different colors in the code and text editor. These marks notify you of problems that Visual Studio detected in your code. For more information, see [Code and Text Editor](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx). For more information about how to enable or disable validation, see: 

- [Options, Text Editor, JavaScript, IntelliSense](https://docs.microsoft.com/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2015)
    
- [How to: Set Validation Options for HTML Editing in Visual Web Developer](https://msdn.microsoft.com/library/0byxkfet(v=vs.100).aspx)
    
- [CSS, see Validation, CSS, Text Editor, Options Dialog Box](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)
    
To review the validation rules of the XML manifest file in your project, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).


### Show an add-in in Excel, Word, or Project and step through your code


If you set the  **Start Document** property of the add-in project to Excel or Word, Visual Studio creates a new document and the add-in appears. If you set the **Start Document** property of the add-in project to use an existing document, Visual Studio opens the document, but you have to insert the add-in manually. If you set the **Start Document** to **Microsoft Project**, you also have to insert the add-in manually.


### To show an Office Add-in in Excel or Word


1. In Excel or Word, on the  **Insert** tab, choose **Office Add-ins**.
    
2. In the list that appears, choose your add-in.
    

### To show an Office Add-in in Project


1. In Project, on the  **Project** tab, choose **Office Add-ins**.
    
2. In the list that appears, choose your add-in.
    
In Visual Studio, you can then set break-points. Then, as you interact with your add-in and step through the code in your HTML, JavaScript, and C# or VB code files.


### Show the Outlook add-in in Outlook and step through your code


To view the add-in in Outlook, open an email message or appointment item.

Outlook activates the add-in for the item as long as the activation criteria are met. The add-in bar appears at the top of the Inspector window or Reading Pane, and your Outlook add-in appears as a button in the add-in bar. If your add-in has an add-in command, a button will appear in the ribbon, either in the default tab or a specified custom tab, and the add-in will not appear in the add-in bar.

To view your Outlook add-in, choose the button for your Outlook add-in.

In Visual Studio, you can set break-points. Then, as you interact with your Outlook add-in and step through the code in your HTML, JavaScript, and C# or VB code files. 

You can also change your code and review the effects of those changes in your Outlook add-in without having to close the Office Add-in and start the project again. In Outlook, just open the shortcut menu for the Outlook add-in, and then choose  **Reload**.


### Modify code and continue to debug the add-in without having to start the project again


You can change your code and review the effects of those changes in your add-in without having to close the host application and start the project again. After you change your code, open the shortcut menu for the add-in, and then choose  **Reload**. When you reload the add-in it becomes disconnected with the Visual Studio debugger. Therefore, you can view the effects of your change, but you cannot step through your code again until you attach the Visual Studio debugger to all of the available Iexplore.exe processes.


### To attach the Visual Studio debugger to all of the available Iexplore.exe processes


1. In Visual Studio, choose  **DEBUG**,  **Attach to Process**.
    
2. In the  **Attach to Process** dialog box, choose all of the available **Iexplore.exe** processes, and then choose the **Attach** button.
    

## Next steps

- [Deploy and publish your Office Add-in](../publish/publish.md)
    
