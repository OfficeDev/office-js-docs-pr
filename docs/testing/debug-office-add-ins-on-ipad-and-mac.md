---
title: Debug Office Add-ins on iPad and Mac
description: ''
ms.date: 03/21/2018
---

# Debug Office Add-ins on iPad and Mac

You can use Visual Studio to develop and debug add-ins on Windows, but you can't use it to debug add-ins on the iPad or Mac. Because add-ins are developed using HTML and Javascript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on an iPad or Mac. 

## Debugging with Safari Web Inspector on a Mac

If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.

To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later. If you don't have an Office Mac build, you can get one by joining the [Office 365 Developer program](https://aka.ms/o365devprogram).

To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

Then, open the Office application and insert your add-in. Right-click the add-in and you should see an **Inspect Element** option in the context menu.  Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.

> [!NOTE]
> Please note that this is an experimental feature and there are no guarantees that we will preserve this functionality in future versions of Office applications.
>
> If you are trying to use the inspector and the dialog flickers, try the following workaround:
> 1. Reduce the size of the dialog.
> 2. Choose **Inspect Element**, which opens in a new window.
> 3. Resize the dialog to its original size.
> 4. Use the inspector as required.

## Debugging with Vorlon.JS on a iPad or Mac

To debug an add-in on iPad or Mac, you can use Vorlon.JS, a debugger for web pages that is similar to the F12 tools. It is designed to work remotely and it enables you to debug web pages across different devices. For more information, see the [Vorlon website](http://www.vorlonjs.com).  


### Install and set up Vorlon.JS  

1.	Log on to the device as an administrator.

2.	Install [Node.js](https://nodejs.org) if it isn't already installed.

3.	Open a **Terminal** window and enter the command `npm i -g vorlon`. The tool is installed to `/usr/local/lib/node_modules/vorlon`.


### Configure Vorlon.JS to use HTTPS

To debug an application using Vorlon.JS, you add a `<script>` tag to the opening page of the application that loads a Vorlon.JS script from a well-known location (for details, see the following procedure). If an add-in is SSL-secured (HTTPS), any scripts that it uses must be hosted from an HTTPS server, including the Vorlon.JS script. Therefore, you must configure Vorlon.JS to use SSL in order to use Vorlon.JS with add-ins.

> [!IMPORTANT]
> [!include[HTTPS guidance](../includes/https-guidance.md)]

1.	In **Finder**, go to `/usr/local/lib/node_modules/vorlon`, open the context menu for (right-click) the `/Server` folder, and then select **Get Info**.

2.	Choose the padlock icon in the lower right corner of the **Server info** window to unlock the folder.

3. In the **Sharing and Permissions** section of the window, set the **Privilege** for the **staff** group to **Read & Write**.

4. Choose the padlock icon again to ***relock*** the folder.

5. Back in **Finder**, expand the `/Server` subfolder, right-click the file `config.json`, and then select **Get Info**.

6. In the **config.json info** window, change the privileges of the file exactly the way you did for its parent `/Server` folder. Be sure to relock and close the window.

7. Back in **Finder**, right-click the file `config.json`, select **Open with**, and then select **TextEdit**. The file opens in a text editor.

8. Change the value of the **useSSL** property to `true`.

9. In the **plugins** section, find the plugin with the **id** of `OFFICE` and the **name** of `Office Addin`. If the **enabled** property for the plug-in is not already `true`, set it to `true`.

10. Save the file and close the editor.

11.	In **Finder**, navigate to `/usr/local/lib/node_modules/vorlon`, right-click the `Server` subfolder, and select **New terminal at folder**.

12.	In the **Terminal** window, enter `sudo vorlon`. You will be prompted to enter your administrator password. The Vorlon server starts. Leave the **Terminal** window open.

13.	Open a browser window and go to `https://localhost:1337`, which is the Vorlon.JS interface. When prompted, choose **Always** to trust the security certificate.

    > [!NOTE]
    > If you are not prompted, you might need to trust the certificate manually. The certificate file is `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt`. Try the following steps. If you have trouble, consult Macintosh or iPad help.
    >
    > 1. Close the browser window and in the **Terminal** window that is running the Vorlon server, use Control-C to stop the server.
    > 2. In **Finder**, right-click the `server.crt` file and select **Keychain Access**. The **Keychain Access** window opens.
    > 3. In the **Keychains** list on the left, select **login** if it is not already selected, and then select **Certificates** in the **Category** section. The certificate **localhost** is listed.
    > 4. Right-click the certificate **localhost** and select **Get Info**. A **localhost** window opens.
    > 5. In the **Trust** section, open the selector labeled **When using this certificate** and select **Always Trust**. 
    > 6. Close the **localhost** window. If the action was successful, the **localhost** certificate in the **Keychain Access** window has a white cross in a blue circle on its icon.


### Configure the add-in for Vorlon.JS debugging

1. Add the following script tag to the `<head>` section of the home.html file (or main HTML file) of your add-in:

    ```html
    <script src="https://localhost:1337/vorlon.js"></script>
    ```  

2. Deploy the add-in web application to a web server that is accessible from the Mac or iPad, such as an Azure website.

3. Update the URL of the add-in in all the places where the URL appears in the add-in manifest.

4. Copy the add-in manifest to the following folder on the Mac or iPad: `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`, where *{host_name}* is Word, Excel, PowerPoint, or Outlook.


### Inspect an add-in in Vorlon.JS

1. If the Vorlon server is not running, in **Finder**, navigate to `/usr/local/lib/node_modules/vorlon`, right-click the `Server` subfolder, and select **New terminal at folder**. 

2.	In the **Terminal** window, enter `sudo vorlon`. You will be prompted to enter your administrator password. The Vorlon server starts. Leave the **Terminal** window open.

3.	Open a browser window and go to `https://localhost:1337`, which is the Vorlon.JS interface.

4. Sideload the add-in. If it is for Excel, PowerPoint, or Word, sideload it as described in [Sideload an Office Add-in on iPad and Mac](sideload-an-office-add-in-on-ipad-and-mac.md). If it is an Outlook add-in, sideload it as described in [Sideload Outlook add-ins for testing](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing). If the add-in does not use add-in commands, it will open immediately. Otherwise, choose the button to open the add-in. Depending on the build of the Office host application, the button will be on either the **Home** tab or an **Add-in** tab.

The add-in will show up in the list of Clients in Vorlon.JS (on the left side of the Vorlon.JS interface) as **{OS} - n**, for some number *n*, and where *{OS}* is the device type, such as "Macintosh".

![Screenshot that shows the Vorlon.js interface](../images/vorlon-interface.png)

The Vorlon tool has a variety of plug-ins. The ones that are currently enabled appear as tabs at the top of the tool. (You can enable more plug-ins by choosing the gears icon on the left.) These plug-ins are  similar to the functions in F12 tools. For example, you can highlight DOM elements, execute commands, and more. For more details, see [Vorlon Documentation Core Plugins](http://vorlonjs.com/documentation/#console).

An **Office Addin** plug-in adds extra capabilities for Office.js, such as exploring the object model, executing Office.js calls, and reading the values of object properties. For instructions, see [VorlonJS plugin for debugging Office Add-in](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/).

> [!NOTE]
> There is no way to set break points in Vorlon.JS.


## Clearing the Office application's cache on a Mac or iPad

Add-ins are cached often in Office for Mac, for performance reasons. Normally, the cache is cleared by reloading the add-in. If  more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.

On a Mac, you can clear the cache manually by deleting everything in the `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.

On an iPad, you can call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.
