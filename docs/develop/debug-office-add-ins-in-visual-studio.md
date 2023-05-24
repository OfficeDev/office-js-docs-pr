---
title: Debug Office Add-ins in Visual Studio
description: Use Visual Studio to debug Office Add-ins in the Office desktop client on Windows.
ms.date: 05/20/2023
ms.localizationpriority: medium
---

# Debug Office Add-ins in Visual Studio

This article describes how to debug client-side code in Office Add-ins that are created with one of the Office Add-in project templates in Visual Studio 2022.  For information about debugging server-side code in Office Add-ins, see [Overview of debugging Office Add-ins - Server-side or client-side?](../testing/debug-add-ins-overview.md#server-side-or-client-side).

> [!NOTE]
> You can't use Visual Studio to debug add-ins in Office on Mac. For information about debugging on a Mac, see [Debug Office Add-ins on a Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md).

## Review the build and debug properties

Before you start debugging, review the properties of each project to confirm that Visual Studio will open the desired Office application and that other build and debug properties are set appropriately.

### Add-in project properties

Open the **Properties** window for the add-in project to review project properties.

1. In **Solution Explorer**, choose the add-in project (*not* the web application project).

1. From the menu bar, choose **View** > **Properties Window**.

The following table describes the properties of the add-in project.

|Property|Description|
|:-----|:-----|
|**Start Action**|Specifies the debug mode for your add-in. This should be set to **Microsoft Edge** for an Outlook add-in. For all other Office applications, it should be set to **Office Desktop Client**.|
|**Start Document**<br/>(Excel, PowerPoint, and Word add-ins only)|Specifies what document to open when you start the project. In a new project, this is set to **[New Excel Workbook]**, **[New Word Document]**, or **[New PowerPoint Presentation]**. To specify a particular document, follow the steps in [Use an existing document to debug the add-in](#use-an-existing-document-to-debug-the-add-in).|
|**Web Project**|Specifies the name of the web project associated with the add-in.|
|**Email Address**<br/>(Outlook add-ins only)|Specifies the email address of the user account in Exchange Server or Exchange Online that you want to use to test your Outlook add-in. If left blank, you'll be prompted for the email address when you start debugging.|
|**EWS Url**<br/>(Outlook add-ins only)|Specifies the Exchange Web Services URL (For example: `https://www.contoso.com/ews/exchange.aspx`). This property can be left blank.|
|**OWA Url**<br/>(Outlook add-ins only)|Specifies the Outlook on the web URL (For example: `https://www.contoso.com/owa`). This property can be left blank.|
|**Use multi-factor auth**<br/>(Outlook add-ins only)|Specifies the boolean value that indicates whether multi-factor authentication should be used. The default is **false**, but the property has no practical effect. If you normally have to provide a second factor to login to the email account, you'll be prompted to when you start debugging. |
|**User Name**<br/>(Outlook add-ins only)|Specifies the name of the user account in Exchange Server or Exchange Online that you want to use to test your Outlook add-in. This property can be left blank.|
|**Project File**|Specifies the name of the file containing build, configuration, and other information about the project.|
|**Project Folder**|Specifies the location of the project file.|

> [!NOTE]
> For an Outlook add-in, you may choose to specify values for one or more of the *Outlook add-in only* properties in the **Properties** window, but doing so isn't required.

### Web application project properties

Open the **Properties** window for the web application project to review project properties.

1. In **Solution Explorer**, choose the web application project.

1. From the menu bar, choose **View** > **Properties Window**.

The following table describes the properties of the web application project that are most relevant to Office Add-in projects.

|Property|Description|
|:-----|:-----|
|**SSL Enabled**|Specifies whether SSL is enabled on the site. This property should be set to **True** for Office Add-in projects.|
|**SSL URL**|Specifies the secure HTTPS URL for the site. Read-only.|
|**URL**|Specifies the HTTP URL for the site. Read-only.|
|**Project File**|Specifies the name of the file containing build, configuration, and other information about the project.|
|**Project Folder**|Specifies the location of the project file. Read-only. The manifest file that Visual Studio generates at runtime is written to the `bin\Debug\OfficeAppManifests` folder in this location.|

## Debug an Excel, PowerPoint, or Word add-in project

This section describes how to start and debug an Excel, PowerPoint, or Word add-in.

### Start the Excel, PowerPoint, or Word add-in project

Start the project by choosing **Debug** > **Start Debugging** from the menu bar or press the F5 button. Visual Studio will automatically build the solution and start the Office host application.

When Visual Studio builds the project, it performs the following tasks:

1. Creates a copy of the XML manifest file and adds it to the  `_ProjectName_\bin\Debug\OfficeAppManifests` directory. The Office application that hosts your add-in consumes this copy when you start Visual Studio and debug the add-in.

1. Creates a set of registry entries on your Windows computer that enables the add-in to appear in the Office application.

1. Builds the web application project, and then deploys it to the local IIS web server (`https://localhost`).

1. If this is the first add-in project that you have deployed to the local IIS web server, you may be prompted to install a Self-Signed Certificate to the current user's Trusted Root Certificate store. This is required for IIS Express to display the content of your add-in correctly.

> [!NOTE]
> If Office uses the Edge Legacy webview control (EdgeHTML) to run add-ins on your Windows computer, Visual Studio may prompt you to add a local network loopback exemption. This is required for the webview control to be able to access the website deployed to the local IIS web server. You can also change this setting anytime in Visual Studio under **Tools** > **Options** > **Office Tools (Web)** > **Web Add-In Debugging**. To find out what webview control is used on your Windows computer, see [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

Next, Visual Studio does the following:

1. Modifies the [SourceLocation](/javascript/api/manifest/sourcelocation) element of the XML manifest file (that was copied to the `_ProjectName_\bin\Debug\OfficeAppManifests` directory) by replacing the `~remoteAppUrl` token with the fully qualified address of the start page (for example, `https://localhost:44302/Home.html`).

1. Starts the web application project in IIS Express.

1. Validates the manifest.  

   > [!IMPORTANT]
   > The Office manifest XSD files that Visual Studio installs are out-of-date. If you get validation errors for the manifest, your first troubleshooting step should be to replace one or more of these files with the latest versions. For detailed instructions, see [Manifest schema validation errors in Visual Studio projects](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects).

1. Opens the Office application and sideloads your add-in.

### Debug the Excel, PowerPoint, or Word add-in

The best method for debugging an add-in in Visual Studio 2022 depends on whether the add-in is running in WebView2, which is the webview control that is associated with Microsoft Edge (Chromium), or an older webview control. To determine which webview control is being used, see [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). If your computer is using WebView2, see [Use the built-in Visual Studio debugger](#use-the-built-in-visual-studio-debugger). For any other webview control, see [Use the browser developer tools](#use-the-browser-developer-tools).

#### Use the built-in Visual Studio debugger

1. Set breakpoints, as needed, in the source JavaScript or TypeScript files. You can do this either before or after you start the add-in as described in the earlier section [Start the Excel, PowerPoint, or Word add-in project](#start-the-excel-powerpoint-or-word-add-in-project).

1. When the add-in is running, use the add-in's UI to run the code that contains your breakpoints. 

> [!IMPORTANT]
> Breakpoints set in `Office.initialize` or `Office.onReady` aren't hit. To debug these methods, see [Debug the initialize and onReady functions](../testing/debug-initialize-onready.md).

> [!TIP]
> If you encounter any problems, there's more information at [Debug a JavaScript or TypeScript app in Visual Studio](/visualstudio/javascript/debug-nodejs?view=vs-2022&preserve-view=true).

#### Use the browser developer tools

1. Follow the steps in the earlier section [Start the Excel, PowerPoint, or Word add-in project](#start-the-excel-powerpoint-or-word-add-in-project).

1. Launch the add-in in the Office application if it isn't already open. For example, if it's a task pane add-in, it will have added a button to the **Home** ribbon (for example, a **Show Taskpane** button). Select the button in the ribbon. 

   > [!NOTE]
   > If your add-in isn't sideloaded by Visual Studio, you can sideload it manually. In Excel, PowerPoint, or Word, choose the **Insert** tab, and then choose the down-arrow located to the right of **My Add-ins**.
   >
   > ![Screenshot showing Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted.](../images/excel-cf-register-add-in-1b.png)
   >
   > In the list of available add-ins, find the **Developer Add-ins** section and select the your add-in to register it.

   > [!TIP]
   > The task pane may appear blank when it first opens. If so, it should render correctly when you launch the debugging tools in a later step.

1. Open the [personality menu](../design/task-pane-add-ins.md#personality-menu) and then choose **Attach a debugger**. This will open the debugging tools for the webview control that Office is using to run add-ins on your Windows computer. You can set breakpoints and step through code as described in one of the following articles:

    - [Debug add-ins using developer tools for Internet Explorer](../testing/debug-add-ins-using-f12-tools-ie.md)
    - [Debug add-ins using developer tools for Edge Legacy](../testing/debug-add-ins-using-devtools-edge-legacy.md)
    - [Debug add-ins using developer tools in Microsoft Edge (Chromium-based)](../testing/debug-add-ins-using-devtools-edge-chromium.md)

1. To make changes to your code, first stop the debugging session in Visual Studio and close the Office application. Make your changes, and start a new debugging session.

## Debug an Outlook add-in project

This section describes how to start and debug an Outlook add-in. Debugging is done in Outlook on the web, not Outlook on Windows.

### Start the Outlook add-in project

Start the project by choosing **Debug** > **Start Debugging** from the menu bar or press the F5 button. Visual Studio will automatically build the solution and launch the Outlook page of your Microsoft 365 tenancy.

When Visual Studio builds the project it performs the following tasks.

1. Prompts you for login credentials. If you're asked to sign in repeatedly or if you receive an error that you are unauthorized, then Basic Auth may be disabled for accounts on your Microsoft 365 tenant. In this case, try using a Microsoft account instead. You can also try setting the property **Use multi-factor auth** to **True** in the Outlook Web Add-in project properties pane. See [Add-in project properties](#add-in-project-properties).

1. Creates a copy of the XML manifest file and adds it to the `_ProjectName_\bin\Debug\OfficeAppManifests` directory. Outlook consumes this copy when you start Visual Studio and debug the add-in.

1. Builds the web application project, and then deploys it to the local IIS web server (`https://localhost`).

1. If this is the first add-in project that you have deployed to the local IIS web server, you may be prompted to install a Self-Signed Certificate to the current user's Trusted Root Certificate store. This is required for IIS Express to display the content of your add-in correctly.

> [!NOTE]
> If Office uses the Edge Legacy webview control (EdgeHTML) to run add-ins on your Windows computer, Visual Studio may prompt you to add a local network loopback exemption. This is required for the webview control to be able to access the website deployed to the local IIS web server. You can also change this setting anytime in Visual Studio under **Tools** > **Options** > **Office Tools (Web)** > **Web Add-In Debugging**. To find out what webview control is used on your Windows computer, see [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

Next, Visual Studio does the following:

1. Modifies the [SourceLocation](/javascript/api/manifest/sourcelocation) element of the XML manifest file (that was copied to the `_ProjectName_\bin\Debug\OfficeAppManifests` directory) by replacing the `~remoteAppUrl` token with the fully qualified address of the start page (for example, `https://localhost:44302/Home.html`).

1. Starts the web application project in IIS Express.

1. Validates the manifest.  

   > [!IMPORTANT]
   > The Office manifest XSD files that Visual Studio installs are out-of-date. If you get validation errors for the manifest, your first troubleshooting step should be to replace one or more of these files with the latest versions. For detailed instructions, see [Manifest schema validation errors in Visual Studio projects](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects).

1. Opens the Outlook page of your Microsoft 365 tenancy in Microsoft Edge.

### Debug the Outlook add-in

The best method for debugging an add-in in Visual Studio 2022 depends on whether the add-in is running in WebView2, which is the webview control that is associated with Microsoft Edge (Chromium), or an older webview control. To determine which webview control is being used, see [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). If your computer is using WebView2, see [Use the built-in Visual Studio debugger](#use-the-built-in-visual-studio-debugger). For any other webview control, see [Use the browser developer tools](#use-the-browser-developer-tools).

#### Use the built-in Visual Studio debugger

1. Set breakpoints, as needed, in the source JavaScript or TypeScript files. You can do this either before or after you start the add-in as described in the earlier section [Start the Outlook add-in project](#start-the-outlook-add-in-project).

1. When the add-in is running, use the add-in's UI to run the code that contains your breakpoints. 

> [!TIP]
>
> - Sometimes in Outlook on the web, the Visual Studio debugger doesn't attach. If you get errors by the breakpoints that indicate they won't be hit, use the browser developer tools to attach to the Visual Studio debugger: After you have pressed F5 to start debugging and Outlook on the web has opened, follow the first four steps in the [Use the browser developer tools in Outlook on the web](#use-the-browser-developer-tools-in-outlook-on-the-web). (Use the instructions for Microsoft Edge (Chromium-based).) After you set a breakpoint in the browser tools and it's hit, execution will pause on the breakpoint in *both* the browser tools *and* in Visual Studio. This indicates that the Visual Studio debugger is attached. At this point, you can close the browser tools and add breakpoints in Visual Studio as you normally would.
> - If you encounter any problems, there's more information at [Debug a JavaScript or TypeScript app in Visual Studio](/visualstudio/javascript/debug-nodejs?view=vs-2022&preserve-view=true).

#### Use the browser developer tools in Outlook on the web

1. In the Outlook page, select an email message or appointment item to open it in its own window. 

1. Press F12 to open the Edge debugging tool.

1. After the tool is open, launch the add-in. For example, in the toolbar at the top of a message, select the **More apps** button, and then select your add-in from the callout that opens.

   ![Screenshot showing the More apps button and the callout that it opens with the add-in's name and icon visible along with other app icons.](../images/outlook-more-apps-button.png)

1. Use the instructions in one of the following articles to set breakpoints and step through code. They each have a link to more detailed guidance.

   - [Debug add-ins using developer tools for Edge Legacy](../testing/debug-add-ins-using-devtools-edge-legacy.md)
   - [Debug add-ins using developer tools in Microsoft Edge (Chromium-based)](../testing/debug-add-ins-using-devtools-edge-chromium.md)

   > [!TIP]
   > To debug code that runs in the `Office.initialize` function or an `Office.onReady` function that runs when the add-in opens, set your breakpoints, and then close and reopen the add-in. For more information about these functions, see [Initialize your Office Add-in](../develop/initialize-add-in.md).

1. To make changes to your code, first stop the debugging session in Visual Studio and close the Outlook pages. Make your changes, and start a new debugging session.

## Use an existing document to debug the add-in

If you have a document that contains test data you want to use while debugging your Excel, PowerPoint, or Word add-in, Visual Studio can be configured to open that document when you start the project. To specify an existing document to use while debugging the add-in, complete the following steps.

1. In **Solution Explorer**, choose the add-in project (*not* the web application project).

1. From the menu bar, choose **Project** > **Add Existing Item**.

1. In the **Add Existing Item** dialog box, locate and select the document that you want to add.

1. Choose the **Add** button to add the document to your project.

1. In **Solution Explorer**, choose the add-in project (*not* the web application project).

1. From the menu bar, choose **View** > **Properties Window**.

1. In the **Properties** window, choose the **Start Document** list, and then select the document that you added to the project. The project is now configured to start the add-in in that document.

## Next steps

After your add-in is working as desired, see [Deploy and publish your Office Add-in](../publish/publish.md) to learn about the ways you can distribute the add-in to users.
