---
title: Sideload Office Add-ins in Office on the web for testing
description: Test your Office Add-in in Office on the web by sideloading.
ms.date: 02/11/2022
ms.localizationpriority: medium
---

# Sideload Office Add-ins in Office on the web for testing

When you sideload an add-in, you're able to install the add-in without first putting it in the add-in catalog. This is useful when testing and developing your add-in because you can see how your add-in will appear and function.

When you sideload an add-in, the add-in's manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.

Sideloading varies between host applications (for example, Excel).

> [!NOTE]
> Sideloading as described in this article is supported on Excel, OneNote, PowerPoint, and Word. To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).

## Sideload an Office Add-in in Office on the web

This process is supported for **Excel**, **OneNote**, **PowerPoint**, and **Word** only. For other host applications, see the manual sideloading instructions in the following section. This example project assumes that you are using a project created with the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md).

1. Open [Office on the web](https://office.live.com/). Using the **Create** option, make a document in **Excel**, **OneNote**, **PowerPoint**, or **Word**. In this new document, select **Share** in the ribbon, select **Copy Link**, and copy the URL.

1. In the root directory of your yo office project files, open the **package.json** file. Within the **config** section of this file, create a `"document"` property. Paste the URL you copied as the value for the `"document"` property. For example, yours will look something like this:

    ```json
      "config": {
        "document": "<YOUR URL>",
        ...
      }
    ```

    > [!TIP]
    > If you are creating an add-in not using our Yeoman generator, you can add query parameters to your document's URL, by appending the following to the existing URL.
    >
    > - The dev server port, such as `&wdaddindevserverport=3000`.
    > - The manifest file name, such as `&wdaddinmanifestfile=manifest1.xml`.
    > - The manifest GUID, such as `&wdaddinmanifestguid=05c2e1c9-3e1d-406e-9a91-e9ac64854143`.
    >
    > If you are using the Yeoman generator, adding this information is not necessary as the Yeoman tooling appends this information automatically.
    > Note that in both cases, however, you can only load manifests from localhost.

1. In the command line starting at the root directory of your project, run the following command. Replace "{url}" with the URL of an Office document on your OneDrive or a SharePoint library to which you have permissions.

    [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

1. The first time you use this method to sideload an add-in on the web, you'll see a dialog asking you to enable developer mode. Select the checkbox for **Enable Developer Mode now** and select **OK**.

1. You will see a second dialog box, asking if you wish to register an Office Add-in manifest from your computer. You should select **Yes**.

1. Your add-in is installed. If it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the task pane should appear.

## Sideload an Office Add-in in Office on the web manually

This method doesn't use the command line and can be accomplished using commands only within the host application (such as Excel).

1. Open [Office on the web](https://office.com/). Open a document in **Excel**, **OneNote**, **PowerPoint**, or  **Word**. On the **Insert** tab on the ribbon in the **Add-ins** section, choose **Office Add-ins**.

1. On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.

    ![The Office Add-ins dialog with a drop-down in the upper right reading "Manage my add-ins" and a drop-down below it with the option "Upload My Add-in".](../images/office-add-ins-my-account.png)

1. **Browse** to the add-in manifest file, and then select **Upload**.

    ![The upload add-in dialog with buttons for browse, upload, and cancel.](../images/upload-add-in.png)

1. Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the pane should appear.

> [!NOTE]
> To test your Office Add-in with Microsoft Edge with the original WebView (EdgeHTML), an additional configuration step is required. In a Windows Command Prompt, run the following line: `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes`. This is not required when Office is using the Chromium-based Edge WebView2. For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

[!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

## Sideload an Office Add-in

1. Sign in to your Microsoft 365 account.

1. Open the App Launcher on the left end of the toolbar and select **Excel**, **PowerPoint**, or **Word**, and then create a new document.

1. Steps 3 - 6 are the same as in the preceding section **Sideload an Office Add-in in Office on the web**.

## Sideload an add-in when using Visual Studio

If you're using Visual Studio to develop your add-in, the process to sideload is similar to manual sideloading to the web. The only difference is that you must update the value of the **SourceURL** element in your manifest to include the full URL where the add-in is deployed.

> [!NOTE]
> Although you can sideload add-ins from Visual Studio to Office on the web, you cannot debug them from Visual Studio. To debug you will need to use the browser debugging tools. For more information, see [Debug add-ins in Office on the web](debug-add-ins-in-office-online.md).

1. In Visual Studio, show the **Properties** window by choosing **View** > **Properties Window**.
1. In the **Solution Explorer**, select the web project. This will display properties for the project in the **Properties** window.
1. In the Properties window, copy the **SSL URL**.
1. In the add-in project, open the manifest XML file. Be sure you are editing the source XML. For some project types Visual Studio will open a visual view of the XML which will not work for the next step.
1. Search and replace all instances of **~remoteAppUrl/** with the SSL URL you just copied. You will see several replacements depending on the project type, and the new URLs will appear similar to `https://localhost:44300/Home.html`.
1. Save the XML file.
1. Right click the web project and choose **Debug** > **Start new instance**. This will run the web project without launching Office.
1. From Office on the web, sideload the add-in using steps previously described in [Sideload an Office Add-in in Office on the web](#sideload-an-office-add-in-in-office-on-the-web).

## Remove a sideloaded add-in

You can remove a previously sideloaded add-in by clearing your browser's cache. If you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you may need to clear your browser's cache and then re-sideload the add-in using the updated manifest. Doing so will allow Office on the web to render the add-in as it's described by the updated manifest.

## See also

- [Sideload Office Add-ins on Mac](sideload-an-office-add-in-on-mac.md)
- [Sideload Office Add-ins on iPad](sideload-an-office-add-in-on-ipad.md)
- [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md)
- [Clear the Office cache](clear-cache.md)
