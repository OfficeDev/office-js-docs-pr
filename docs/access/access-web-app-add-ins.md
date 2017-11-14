---
title: Create add-ins for Access web apps
description: Develop an Office Add-in that targets Access web apps by using Visual Studio.
ms.date: 11/20/2017 
---

# Create add-ins for Access web apps

> [!IMPORTANT]
> We no longer recommend that you create and use Access web apps and databases in SharePoint. As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.

This article shows you how to use Visual Studio 2015 to develop an Office Add-in that targets Access web apps.

> [!NOTE]
> For information about developing solutions for Access using VBA, see [Access](https://msdn.microsoft.com/en-us/library/fp179695.aspx) on MSDN.

## Prerequisites

To create an Office Add-in that targets Access web apps, you need:

- Visual Studio 2015

- A SharePoint Online site (included in many Office 365 subscriptions). This site must have an add-in catalog. For more information, see [Set up an add-in catalog on SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md#set-up-an-add-in-catalog).

> [!NOTE]
> Office Add-ins work with Access web apps hosted on SharePoint Online or Office 365. The Access 2013 desktop application doesn't support Office Add-ins. Office Add-ins targeting Access web apps are supported by version 1.1 and later of Office.js.


## Create a project in Visual Studio

1. Open Visual Studio, and in the menu, select **File** > **New** > **Project**.

2. In the **New Project** dialog box, in the left pane, go to **Installed** > **Templates** > **Visual C#** > **Office/SharePoint** > **Office Add-ins**.

   > [!NOTE]
   > If you don't have this template installed, see the [Latest Microsoft Office Developer Tools for Visual Studio 2015](https://blogs.msdn.microsoft.com/visualstudio/2015/11/23/latest-microsoft-office-developer-tools-for-visual-studio-2015/) blog post for information.

3. In the **New Project** dialog box, in the center pane, select **Office Add-in**.

4. At the bottom of the dialog box, enter a name for your project, and then select **OK**.

5. In the **Create Office Add-in** dialog box, select **Content**, and then select **Next**.

6. On the next screen of the **Create Office Add-in** dialog box, select either **Basic Add-in** or **Document Visualization Add-in**, and make sure that the check box for **Access** is selected.

7. When done, select **Finish**. Visual Studio creates a starter project for you to base your work on.

8. In **Solution Explorer**, select the project's web project (**project_name>Web**). In the properties pane, find the entry for **SSL URL**. This should look something like `https://localhost:44314/`. Select this URL, and copy it to your clipboard. You will need it shortly.

9. Right-click the name of your project in **Solution Explorer**. In the context menu, select **Publish**.

10. In the Publish Your Add-in Wizard, select the drop-down list next to **Current profile**. In this list, select **new**.

11. In the **Publish Office and SharePoint Add-ins** dialog box, select **Create new profile**, enter a recognizable name for the profile, and then select **Finish**. The **Publish Office and SharePoint Add-ins** dialog box closes, returning you to the Publish Your Add-in Wizard.

12. In the wizard, select **Package the add-in**. This finalizes your add-in so that it can be published to an add-in catalog in SharePoint.

13. On the next page, for **Where is your website hosted?**, enter the URL for the host of your website. This can be the **SSL URL** value that you copied in step 8. Select **Finish**.

14. In **Solution Explorer**, right-click the project's manifest node (directly under the project name), and select **Open Folder in File Explorer**. Make a note of the path to this file. You will need this value later.

> [!NOTE]
> You can't debug the add-in without deploying it with an Access web app.

## Review the manifest and the Home.html file

1. In your Visual Studio project, open the **Home.html** file, and find the lines that reference the office.js script library.

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

    > [!NOTE]
    > This script tag references version 1.1 (and later) of Office.js. Access uses API elements introduced in version 1.1.

2. Open the manifest file associated with your project. This file is named after the name of your project, and has the extension `.xml`.

3. In the manifest file, find the **Hosts** section and look for a **Host** entry.

    ```xml
    <Hosts> <Host Name="Database" /> </Hosts>
    ```
    
    > [!NOTE]
    > This is where the applications that can use the add-in are listed. Because you selected **Access** in the **Create Office Add-in** dialog box, **Database** is listed. If you included Excel, there is an entry for **Workbook** as well.

Office and SharePoint Add-ins are web-based. The code for the add-in must be hosted on a web server. For this example, the web server is your development computer. The server must be running to serve the add-in for testing, which means that Visual Studio must be running the add-in at the time that you view and debug it in SharePoint.

For a user to find and use the add-in, it needs to be registered with an Add-in Catalog in SharePoint. The information that the Add-in Catalog needs is contained in the manifest file.

> [!NOTE]
> You need to create an Access web app to host your Office Add-in.

## Publish your add-in to a SharePoint Online catalog

1.  Sign in to SharePoint Online or Office 365, and then go to the **SharePoint admin center** by selecting **Admin** in the Office 365 toolbar at the top of the page.

2. On the **SharePoint admin center** page, in the link bar on the left, select **add-ins**. This takes you to the add-ins view.

3. In the center pane of the page, select **Add-in Catalog**.

4. On the **Catalog** page, select **Distribute Office Add-ins**. This takes you to a directory page called **Office Add-ins** that lists all installed Office Add-ins.

5. At the top of the **Office Add-ins** page, select **new add-in**.

6. In the **Add a document** dialog box, select **Browse**, and then go to the location of the manifest file in your Visual Studio project. If you copied the address of your manifest file earlier, you can paste it into this dialog.

7. Select the manifest file in your project, and then select **OK**. SharePoint adds your add-in to the local SharePoint library.

> [!NOTE]
> This procedure assumes that you have created a test site on SharePoint. If you haven't, you can do so on the **Sites** tab at the top of the SharePoint window. You can use an existing Access web apps if you have one available.

## Create an Access web app to host your add-in

1. Go to your test site. In the left link bar, select **Site Contents**.

2. On the **Site Contents** page, select **Add an add-in**.

3. In the **Site Contents - Your Add-ins** page, use the search bar at the top of the page to search for **Access App**.

4. You should see a tile for **Access App**.

    > [!NOTE]
    > Remember that this is not your Office Add-in, it is a new Access web app. This Access web app hosts your Office Add-in.

5. Selecting this tile brings up the **Adding an Access app** dialog box. Enter a unique name for your Access web app, and select **Create**. It might take a while for SharePoint to create your app. When it is finished, you see your Access web app listed on the **Site Contents** page with a **new** label by it.

6. The Access web app now requires you to open it in the desktop version of Microsoft Access 2013 and add data to it before it can be opened and viewed in SharePoint.

## Add your add-in to an Access web app

1. Open an Access web app.

2. On the SharePoint tab bar, select the gear icon in the upper-left corner. A menu appears. Select the **Office Add-ins** menu item. This opens the **Office Add-ins** dialog box.

3. Select the **My Organization** view and wait a moment for SharePoint to fill the dialog box with the Office Add-ins that are available to you.

4. One of the add-ins in the dialog box should be the Office Add-in that you registered in a previous procedure. Select that add-in to insert it in your Access web app. Remember that the app must be running in Visual Studio to be detected and displayed on your Access web apps page.

## Debug your add-in for Office

To debug your add-in, in Internet Explorer, select F12 or the gear icon in the browser's tab bar (not the gear icon on the SharePoint page). This brings up the F12 debug tools provided by Internet Explorer 11. If you are using another browser, check your browser documentation to determine how to enter debug mode.

At this point, you can set breakpoints, step through your JavaScript code, explore the DOM, and modify the code to confirm that your changes appear in the Office Add-in targeting Access web apps. For more information, see [Using the F12 developer tools](http://msdn.microsoft.com/library/ie/bg182326%28v=vs.85%29).

## Next steps

To learn more about how to implement an Office Add-in that manipulates data in an Access web app, download the sample [Office 365: Bind and manipulate data in an Access web app](https://code.msdn.microsoft.com/officeapps/Office-365-Bind-and-4876274e).

## Additional resources

- [Understanding the JavaScript API for add-ins](../develop/understanding-the-javascript-api-for-office.md)
- [JavaScript API for Office](../../reference/javascript-api-for-office.md)
