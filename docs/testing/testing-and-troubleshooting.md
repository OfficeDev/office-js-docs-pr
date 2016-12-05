
# Troubleshoot user errors with Office Add-ins

At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in. 

You can also use [Fiddler](http://www.telerik.com/fiddler) to identify and debug issues with your add-ins.

After you resolve the user's issue, you can [respond directly to customer reviews in the Office Store](https://msdn.microsoft.com/library/jj635874.aspx).

## Common errors and troubleshooting steps

The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.



|**Error message**|**Resolution**|
|:-----|:-----|
|App error: Catalog could not be reached|Verify firewall settings."Catalog" refers to the Office Store. This message indicates that the user cannot access the Office Store.|
|APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.|Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/en-us/kb/2986156/).|
|Error: Object doesn't support property or method 'defineProperty'|Confirm that Internet Explorer is not running in Compatibility Mode. Go to Tools >  **Compatibility View Settings**.|
|Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.|Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).|

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
    
Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/en-us/kb/2986156/).


## Additional resources



- [Debug add-ins in Office Online](../testing/debug-add-ins-in-office-online.md)
    
- [Sideload an Office Add-in on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
    
- [Debug Office Add-ins on iPad and Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)
    
- [Create and debug Office Add-ins in Visual Studio](../../docs/get-started/create-and-debug-office-add-ins-in-visual-studio.md)
    
- [Deploy and install Outlook add-ins for testing](../outlook/testing-and-tips.md)
    
