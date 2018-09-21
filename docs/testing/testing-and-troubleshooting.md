---
title: Troubleshoot user errors with Office Add-ins
description: ''
ms.date: 01/23/2018
---

# Troubleshoot user errors with Office Add-ins

At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in. 

You can also use [Fiddler](http://www.telerik.com/fiddler) to identify and debug issues with your add-ins.

After you resolve the user's issue, you can [respond directly to customer reviews in AppSource](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).

## Common errors and troubleshooting steps

The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.



|**Error message**|**Resolution**|
|:-----|:-----|
|App error: Catalog could not be reached|Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.|
|APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.|Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).|
|Error: Object doesn't support property or method 'defineProperty'|Confirm that Internet Explorer is not running in Compatibility Mode. Go to Tools >  **Compatibility View Settings**.|
|Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.|Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).|


## Outlook add-in doesn't work correctly

If an Outlook add-in running on Windows is not working correctly, try turning on script debugging in Internet Explorer. 


- Go to Tools >  **Internet Options** > **Advanced**.
    
- Under  **Browsing**, uncheck  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)**.
    
We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.


## Add-in doesn't activate in Office 2013

If the add-in doesn't activate when the user performs the following steps:


1. Signs in with their Microsoft account in Office 2013.
    
2. Enables two-step verification for their Microsoft account.
    
3. Verifies their identity when prompted when they try to insert an add-in.
    
Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).


## Add-in doesn't load in task pane or other issues with the add-in manifest

See [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md) to debug add-in manifest issues.


## Add-in dialog box cannot be displayed

When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:

"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."

![A screen shot of the dialog box error message](http://i.imgur.com/3mqmlgE.png)

|**Affected browsers**|**Affected platforms**|
|:--------------------|:---------------------|
|Internet Explorer, Microsoft Edge|Office Online|

To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.

> [!IMPORTANT]
> Do not add the URL for an add-in to your list of trusted sites if you don't trust the add-in.

To add a URL to your list of trusted sites:

1. In Internet Explorer, choose the Tools button, and go to **Internet options** > **Security**.
2. Select the **Trusted sites** zone, and choose **Sites**.
3. Enter the URL that appears in the error message, and choose **Add**.
4. Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.

This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.

```js

Office.context.ui.displayDialogAsync(startAddress, {displayInFrame:true}, callback);
```

## Changes to add-in commands including ribbon buttons and menu items do not take effect
Sometimes changes to add-in commands such as the icon for a ribbon button or the text of a menu item do not seem to take effect. Clear the Office cache of the old versions.

#### For Windows:
Delete the content of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.

#### For Mac:
Delete the content of the folder `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.

#### For iOS:
Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.

## See also

- [Debug add-ins in Office Online](debug-add-ins-in-office-online.md) 
- [Sideload an Office Add-in on iPad and Mac](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Debug Office Add-ins on iPad and Mac](debug-office-add-ins-on-ipad-and-mac.md)  
- [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md)
    
