---
title: Sideload Office Add-ins for testing from a network share
description: Learn how to sideload an Office Add-in for testing from a network share.
ms.date: 05/21/2025
ms.localizationpriority: medium
---

# Sideload Office Add-ins for testing from a network share

You can test an Office Add-in in an Office client that's on Windows by publishing the manifest to a network file share (instructions follow). This deployment option is intended to be used when you've completed development and testing on a localhost and want to test the add-in from a non-local server or cloud account.

> [!IMPORTANT]
> Deployment by network share isn't supported for production add-ins. This method has the following limitations.
>
> - The add-in can only be installed on Windows computers.
> - Add-ins that use the [unified manifest for Microsoft 365](../develop/unified-manifest-overview.md) aren't supported when published to a network share.
> - If a new version of an add-in changes the ribbon, such as by adding a custom tab or custom button to it, each user will have to reinstall the add-in.

> [!NOTE]
> If your add-in project was created with a sufficiently recent version of the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md), the add-in will automatically sideload in the Office desktop client when you run `npm start`.

This article applies only to testing Word, Excel, PowerPoint, and Project add-ins and only on Windows. If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in.

- [Sideload Office Add-ins in Office on the web for testing](sideload-office-add-ins-for-testing.md)
- [Sideload Office Add-ins on Mac for testing](sideload-an-office-add-in-on-mac.md)
- [Sideload Office Add-ins on iPad for testing](sideload-an-office-add-in-on-ipad.md)
- [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md)

The following video walks you through the process of sideloading your add-in in Office on the web or desktop using a shared folder catalog.  

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## Share a folder

1. In File Explorer on the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.

1. Open the context menu for the folder you want to use as your shared folder catalog (for example, right-click the folder) and choose **Properties**.

1. Within the **Properties** dialog window, open the **Sharing** tab and then choose the **Share** button.

    :::image type="content" source="../images/sideload-windows-properties-dialog.png" alt-text="Folder Properties dialog with the Sharing tab and Share button highlighted.":::

1. Within the **Network access** dialog window, add yourself and any other users and/or groups with whom you want to share your add-in. You'll need at least **Read/Write** permission to the folder. After you've finished choosing people to share with, choose the **Share** button.

1. When you see the **Your folder is shared** confirmation, make note of the full network path that's displayed immediately following the folder name. (You'll need to enter this value as the **Catalog Url** when you [specify the shared folder as a trusted catalog](#specify-the-shared-folder-as-a-trusted-catalog), as described in the next section of this article.) Choose the **Done** button to close the **Network access** dialog window.

   :::image type="content" source="../images/sideload-windows-network-access-dialog.png" alt-text="Network access dialog with the share path highlighted.":::

1. Choose the **Close** button to close the **Properties** dialog window.

## Specify the shared folder as a trusted catalog

There are two options for how you specify this trust. Follow the instructions for the option that works better for your setup.

- [Configure the trust manually.](#configure-the-trust-manually)
- [Configure the trust with a Registry script.](#configure-the-trust-with-a-registry-script)

### Configure the trust manually

1. Open a new document in Excel, Word, PowerPoint, or Project.

1. Choose the **File** tab, and then choose **Options**.

1. Choose **Trust Center**, and then choose the **Trust Center Settings** button.

1. Choose **Trusted Add-in Catalogs**.

1. In the **Catalog Url** box, enter the full network path to the folder that you [shared](#share-a-folder) previously. If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.

    :::image type="content" source="../images/sideload-windows-properties-dialog-2.png" alt-text="Folder Properties dialog with the Sharing tab and network path highlighted.":::

1. After you've entered the full network path of the folder into the **Catalog Url** box, choose the **Add catalog** button.

1. Select the **Show in Menu** check box for the newly-added item, and then choose the **OK** button to close the **Trust Center** dialog window.

    :::image type="content" source="../images/sideload-windows-trust-center-dialog.png" alt-text="Trust Center dialog with catalog selected.":::

1. Choose the **OK** button to close the **Options** dialog window.

1. Close and reopen the Office application so your changes will take effect.

### Configure the trust with a Registry script

1. In a text editor, create a file named **TrustNetworkShareCatalog.reg**.

1. Add the following content to the file.

    ```text
    Windows Registry Editor Version 5.00

    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{-random-GUID-here-}]
    "Id"="{-random-GUID-here-}"
    "Url"="\\\\-share-\\-folder-"
    "Flags"=dword:00000001
    ```

1. Use one of the many online GUID generation tools, such as [GUID Generator](https://guidgenerator.com/), to generate a random GUID, and within the TrustNetworkShareCatalog.reg file, replace the string "-random-GUID-here-" *in both places* with the GUID. (The enclosing `{}` symbols should remain.)

1. Replace the `Url` value with the full network path to the folder that you [shared](#share-a-folder) previously. (Note that any `\` characters in the URL must be doubled.) If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.

    :::image type="content" source="../images/sideload-windows-properties-dialog-2.png" alt-text="Folder Properties dialog with the Sharing tab and network path highlighted.":::

1. The file should now look like the following. Save it.

    ```text
    Windows Registry Editor Version 5.00

    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{01234567-89ab-cedf-0123-456789abcedf}]
    "Id"="{01234567-89ab-cedf-0123-456789abcedf}"
    "Url"="\\\\TestServer\\OfficeAddinManifests"
    "Flags"=dword:00000001
    ```

1. Close *all* Office applications.

1. Run the TrustNetworkShareCatalog.reg just as you would any executable, such as double-clicking it.

## Sideload your add-in

1. Put the manifest XML file of any add-in that you're testing into the shared folder catalog. Note that you deploy the web application itself to a web server. Be sure to specify the URL in the `<SourceLocation>` element of the manifest file.

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

    > [!NOTE]
    > For Visual Studio projects, use the manifest built by the project in the `{projectfolder}\bin\Debug\OfficeAppManifests` folder.

1. In Excel, Word, or PowerPoint, select **Home** > **Add-ins** from the ribbon, then select **Advanced**. In Project, select **My Add-ins** on the **Project** tab of the ribbon.

1. Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.

1. Select the name of the add-in and choose **Add** to insert the add-in.

## Remove a sideloaded add-in

You can remove a previously sideloaded add-in by clearing the Office cache on your computer. Details on how to clear the cache on Windows can be found in the article [Clear the Office cache](clear-cache.md#clear-the-office-cache-on-windows).

## See also

- [Validate an Office Add-in's manifest](troubleshoot-manifest.md)
- [Clear the Office cache](clear-cache.md)
- [Publish your Office Add-in](../publish/publish.md)
