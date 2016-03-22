
# Create and debug Office Add-ins in Visual Studio




 >**Note**  These instructions are based on Visual Studio 2015. If you're using another version of Visual Studio, the procedures might vary slightly.



## Create an Office Add-in project in Visual Studio


To get started, use the  **Office Add-ins** project template in Visual Studio. The **Create Office Add-in** wizard asks you to choose the type of add-in that you want to create, and then provides selections for the default configuration of your add-in.


1. On the Visual Studio menu bar, choose  **File**,  **New**,  **Project**.
    
2. In the list of project types under  **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose  **Office Add-ins**, and then choose  **Office Add-in**.
    
3. Name the project, and then choose  **OK**.
    
4. The  **Create Add-in for Office** dialog box opens. Choose the type of add-in you want to create, and then choose the **Next** button
    
5. Select the default options for the add-in you want to create, and then choose  **Finish**.
    
    Visual Studio creates the project, and its files appear in  **Solution Explorer**. The default Home.html page opens in Visual Studio.
    
In Visual Studio 2015, some of the add-in project templates have been updated to reflect additional functionality:


- Content add-ins can appear in the body of Access and PowerPoint documents, in addition to Excel spreadsheets. You can also choose the Basic Project option to create a basic content add-in project with minimal starter code, or the Document Visualization Project option (for Access and Excel only) to create a more full-featured content add-in that includes starter code to visualize and bind to data.
    
- Outlook add-ins include options not just for including your add-in in email messages or appointments, but also for specifying whether the add-in is available when an email message or appointment is being composed as well as read.
    

 >**Note**  In Visual Studio most options are understandable from their descriptions except for the  **Email Message** checkbox. Use that checkbox if you want to create an Outlook add-in that appears not just with mail items, but also with meeting requests, responses, and cancellations.

When you've completed the wizard, Visual Studio creates a solution for you that contains two projects.



|**Project**|**Description**|
|:-----|:-----|
|Add-in project|Contains only an XML manifest file, which contains all the settings that describe your add-in. These settings help the Office host determine when your add-in should be activated and where the add-in should appear. Visual Studio generates the contents of this file for you so that you can run the project and use your add-in immediately. You change these settings any time by using the Manifest editor.|
|Web application project|Contains the content pages of your add-in, including all the files and file references that you need to develop Office-aware HTML and JavaScript pages. While you develop your add-in, Visual Studio hosts the web application on your local IIS server. When you're ready to publish, you'll have to find a server to host this project.To learn more about ASP.NET web application projects, see [ASP.NET Web Projects](http://msdn.microsoft.com/en-us/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).|

## Modify your add-in settings


To modify the settings of your add-in, open the  **Manifest Designer**. The Manifest Designer is a property page-like editor that enables you to modify the most common settings of your add-in in a more visual way. To open the  **Manifest Designer** from **Solution Explorer**, expand the "Office" add-in project node, choose the folder that contains the XML manifest, and then press the ENTER key. 

For advanced settings, such as the target locale of the add-in, edit the XML manifest file of the project directly. In  **Solution Explorer** expand the add-in project node, expand the folder that contains the XML manifest and choose the XML manifest. You can point to any element in the file to view a tooltip that describes the purpose of the element. For a complete list of descriptions, see [Office Add-ins XML manifest](../../docs/overview/add-in-manifests.md).

The  **Manifest Designer** has multiple tabs, including **General**, which contains general add-in settings such as add-in name and icon,  **Activation**, which enables you to specify the add-in's requirements such as required API sets and target applications, and  **App Domains**, which enables you to specify the domains of pages consumed by your add-in. For Outlook add-ins, there is no  **Activation** tab but there are two other tabs: **Read Form** and **Compose Form**.


### General tab settings

The following table describes the fields that appear in the  **General** tab of the manifest editor.



|**Property**|**Corresponding value in the XML manifest file**|**Description**|
|:-----|:-----|:-----|
|**Display Name**| `DefaultValue` attribute of the [DisplayName](http://msdn.microsoft.com/en-us/library/529159ca-53bf-efcf-c245-e572dab0ef57%28Office.15%29.aspx) element.|Name that appears in the UI of the host application. For example, in Excel, when a user chooses  **Insert**,  **Add-in**, this name appears in the list of available add-ins. This name can also appear as a title above the add-in.|
|**Add-in type**|[OfficeApp](http://msdn.microsoft.com/en-us/library/4537b0a6-a741-332d-9e8f-4341c8b50b6a%28Office.15%29.aspx) complexType.|Type of the add-in. A value of  **Task Pane Add-in** indicates that the add-in appears in the task pane of the Office application. A value of **Content Add-in** indicates that the add-in appears in the body of a document. A value of **Outlook Add-in** indicates that the add-in appears adjacent to a message or appointment item.|
|**Version**|[Version](http://msdn.microsoft.com/en-us/library/6a8bbaa5-ee8c-6824-4aba-cb1a804269f6%28Office.15%29.aspx) element.|Specifies the version of the add-in.|
|**Provider name**|[ProviderName](http://msdn.microsoft.com/en-us/library/0062693a-fafa-ea2d-051a-75dac0f6c323%28Office.15%29.aspx) element.|Specifies the name of the individual or company that developed the add-in.|
|**Description**| `DefaultValue` attribute of the [Description](http://msdn.microsoft.com/en-us/library/bcce6bad-23d0-7631-7d8c-1064b8453b5a%28Office.15%29.aspx) element.|Description that appears when a user points to the add-in name in the list of available add-ins.|
|**Icon**| `DefaultValue` attribute of the [IconUrl](http://msdn.microsoft.com/en-us/library/c7dac2d4-4fda-6fc7-3774-49f02b2d3e1e%28Office.15%29.aspx) element.|32 x 32 pixel image that appears for your add-in in the ribbon of the host application.|
|**High resolution icon**| `DefaultValue` attribute of the [IconUrl](http://msdn.microsoft.com/en-us/library/c7dac2d4-4fda-6fc7-3774-49f02b2d3e1e%28Office.15%29.aspx) element.|64 x 64 pixel image that appears for your add-in in the ribbon of the host application.|
|**Source location**| `DefaultValue` attribute of the [SourceLocation](http://msdn.microsoft.com/en-us/library/e6ea8cd4-7c8b-1da7-d8f8-8d3c80a088bc%28Office.15%29.aspx) element.|Location of the first page that appears in the add-in when it is activated in the host application. The default value of this property is the default HTML file of your project.|
|**Support Url**| `DefaultValue` attribute of the [SupportUrl](http://msdn.microsoft.com/en-us/library/61cff5aa-929f-7d6a-2ce9-0b92b2d6e0a7%28Office.15%29.aspx) element.|Specifies the URL of a page that provides support information for the add-in.|
|**Requested Height** (content and Outlook add-ins only)|[RequestedHeight](http://msdn.microsoft.com/en-us/library/fe949c28-a9ff-26dd-6a80-5f81abc330e8.aspx) element.|Number of pixels that the add-in requires as the height of the add-in pane.|
|**Requested Width** (content add-ins only)|[RequestedWidth](http://msdn.microsoft.com/en-us/library/29032529-6661-fb99-1ff3-c02cc474017f.aspx) element.|Number of pixels that the add-in requires as the width of the add-in pane.|
|**Permissions**|[Permissions](http://msdn.microsoft.com/en-us/library/d4cfe645-353d-8240-8495-f76fb36602fe%28Office.15%29.aspx) element.|Permissions required by this add-in.|
|**Entity Highlighting** (Outlook add-ins only)|[DisableEntityHighlighting element (MailApp complexType) (app manifest schema v1.1)](http://msdn.microsoft.com/library/bf67a8d6-cb8f-7a58-d09d-4d5c7679d10f%28Office.15%29.aspx) element.|Specifies whether entity highlighting should be turned off for this Outlook add-in.|

### Activation tab settings

This tab enables you to list the sets of Office JavaScript APIs that must be supported on a target Office application for your add-in to activate. For example, if your add-in binds to a table, you'd add the TableBindings API set to the list. This feature helps prevent users from getting ugly JavaScript errors by inadvertently activating your add-in on an Office host that doesn't support the functionality contained in your add-in.

The following table describes the fields that appear in the  **Activation** tab of the manifest editor. This tab appears only in content add-ins and task pane add-ins.



|**Property**|**Description**|
|:-----|:-----|
|**Required API Sets**|Enables you to specify the API set names and the minimum versions required by your add-in to activate properly. To add an API set, choose it in the drop-down list. To delete an API set, choose it in the list and then choose the  **Delete** key. As you specify API sets, the page displays the Office clients that support that combination of API sets.<br/> **Note**  Because some API sets support only certain Office clients, specifying more API sets decreases the number of Office clients on which your add-in can activate.|
|**Applications**|Enables you to choose the Office applications that you want your add-in to target. You can target any Office application that's available in Office 365 and Office 2013 SP1, or you can target specific Office applications.|
|**IntelliSense**|Enables you to choose whether IntelliSense displays syntax information for all Office JavaScript APIs, or only for APIs in the  **Required API Sets** list.|
|**Summary**|The summary section shows where the add-in will be activated based on your input in the  **Required API Sets** and the **Applications** sections.|

### App Domains tab settings

App Domains settings are used when the add-in needs to communicate with a remote domain (cross-domain communication). The following table describes the fields that appear in the  **Add-in Domains** tab of the manifest editor. These settings enable you to specify the domains that this add-in uses to load pages.



|**Property**|**Description**|
|:-----|:-----|
|**Enter the URL of a domain**|Enables you to enter a URL of a domain that is used by your add-in. Choose the  **Add** button to add it to the **AppDomains** list.|
|**AppDomains**|The list of add-in domains that you've specified. You can delete items from the list by choosing them and then choosing the  **Remove** button.|

### Read Form tab settings

Read Form settings are only available for Outlook add-ins. These settings specify when a read form add-in is activated and also its UI properties.


|**Property**|**Description**|
|:-----|:-----|
|Activation|Specifies the activation rules for the read form add-in. You select the appropriate rule or rules in the tree pane or you can add rules by choosing the  **Add** drop down list. You can optionally specify the message class of the mail or appointment item.|
|Source location|Sepcifies the source location for the add-in.|
|Requested height|Specifies the requested height in pixels|
|Enable the add-in to appear in items that are opened on a tablet|Check this box if the add-in will appear on a tablet. If selected, you must also provide the source location for the tablet code and the requested height when on a tablet.|
|Enable the add-in to appear in items that are opened on a phone|Check this box if the add-in will appear on a phone. If selected, you must also provide the source location for the phone code.|

### Compose Form tab settings

Compose Form settings are only available for Outlook add-ins. These settings specify when a compose form add-in is activated and also its UI properties.


|**Property**|**Description**|
|:-----|:-----|
|Activation|Specifies the item types that will activate a compose form add-in. You can choose either or both of  **email messages** or **appointments**.|
|Source location|Sepcifies the source location for the add-in.|
|High resolution icon:|Specifies the URL of the image that is used to represent the add-in on a high DPI display. It must be 128 x 128 pixels.|
|Enable the add-in to appear in items that are opened on a tablet|Check this box if the add-in will appear on a tablet. If selected, you must also provide the source location for the tablet code .|
|Enable the add-in to appear in items that are opened on a phone|Check this box if the add-in will appear on a phone. If selected, you must also provide the source location for the phone code.|

## Develop the contents of your add-in


While the add-in project lets you modify the settings that describe your add-in, the web application provides the content that appears in the add-in. 

The web application project contains a default HTML page and JavaScript file that you can use to get started. The project also contains a JavaScript file that is common to all pages that you add to your project. These files are convenient because they contain references to other JavaScript libraries including the JavaScript API for Office. 

As your add-in becomes more sophisticated, you can add more HTML and JavaScript files. You can use the contents of the default HTML and JavaScript files as examples of the types of references you might want to add to other pages in your project to make them work with your add-in. The following table describes default HTML and JavaScript files.



|**File**|**Description**|
|:-----|:-----|
|**Home.html**|Located in the  **Home** folder of the project, this is default HTML page of the add-in. This page appears as the first page inside of the add-in when it is activated in a document, email message or appointment item. This file is convenient because it contains all of the file references that you need to get started. When you are ready to create your first add-in, just add your HTML code to this file.|
|**Home.js**|Located in the  **Home** folder of the project, this is the JavaScript file associated with the Home.js page. You can place any code that is specific to the behavior of the Home.html page in the Home.js file. The Home.js file contains some example code to get you started.|
|**App.js**|Located in the  **Add-in** folder of the project, this is the default JavaScript file of the entire add-in. You can place code that is common to the behavior of multiple pages of your add-in in the App.js file. The App.js file contains some example code to get you started.|

 >**Note**  You don't have to use these files. Feel free to add other files to the project and use those instead. If you want another HTML file to appear as the initial page of the add-in, open the manifest editor, and then point the  **SourceLocation** property to the name of the file.


## Debug your add-in


When you are ready to start your add-in, review build and debug related properties, and then start the solution.


### Review the build and debug properties

Before you start the solution, verify that Visual Studio will open the host application that you want. That information appears in the property pages of the project along with several other properties that relate to building and debugging the add-in.


### To open the property pages of a project


1. In  **Solution Explorer**, choose the project name.
    
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
    
     >**Note**  Choose the add-in project and not the web application project.
2. On the  **Project** menu, choose **Add Existing Item**.
    
3. In the  **Add Existing Item** dialog box, locate and select the document that you want to add.
    
4. Choose the  **Add** button to add the document to your project.
    
5. In  **Solution Explorer**, open the shortcut menu for the project, and then choose  **Properties**.
    
    The property pages for the project appear.
    
6. In the  **Start Document** list, choose the document that you added to the project, and then choose the **OK** button to close the property pages.
    

### Start the solution


Visual Studio will automatically build the solution when you start it. You can start the solution from the  **Menu** bar by choosing **Debug**,  **Start**. 


 >**Note**  If script debugging isn't enabled in Internet Explorer, you won't be able to start the debugger in Visual Studio. You can enable script debugging by opening the  **Internet Options** dialog box, choosing the **Advanced** tab, and then clearing the **Disable Script Debugging (Internet Explorer)** and **Disable Script Debugging (Other)** check boxes.

Visual Studio builds the project and does the following:


1. Creates a copy of the XML manifest file and adds it to  _ProjectName_\Output directory. The host application consumes this copy when you start Visual Studio and debug the add-in.
    
2. Creates a set of registry entries on your computer that enable the add-in to appear in the host application.
    
3. Builds the web application project, and then deploys it to the local IIS web server (http://localhost). 
    
Next, Visual Studio does the following:


1. Modifies the [SourceLocation](http://msdn.microsoft.com/en-us/library/e6ea8cd4-7c8b-1da7-d8f8-8d3c80a088bc%28Office.15%29.aspx) element of the XML manifest file by replacing the ~remoteAppUrl token with the fully qualified address of the start page (for example, http://localhost/MyAgave.html).
    
2. Starts the web application project in IIS Express.
    
3. Opens the host application. 
    
Visual Studio doesn't show validation errors in the  **OUTPUT** window when you build the project. Visual Studio reports errors and warnings in the **ERRORLIST** window as they occur. Visual Studio also reports validation errors by showing wavy underlines (known as squiggles) of different colors in the code and text editor. These marks notify you of problems that Visual Studio detected in your code. For more information, see [Code and Text Editor](http://go.microsoft.com/fwlink/?LinkID=128497). For more information about how to enable or disable validation, see: 


- [Options, Text Editor, JavaScript, IntelliSense](http://go.microsoft.com/fwlink/?LinkID=238779)
    
- [How to: Set Validation Options for HTML Editing in Visual Web Developer](http://msdn.microsoft.com/en-us/library/vstudio/0byxkfet%28v=vs.100%29.aspx)
    
- [CSS, see Validation, CSS, Text Editor, Options Dialog Box](http://go.microsoft.com/fwlink/?LinkID=238780)
    
To review the validation rules of the XML manifest file in your project, see [Office Add-ins XML manifest](../../docs/overview/add-in-manifests.md).


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

- [Publish your Office Add-in](../publish/publish.md)
    
