
# Deploy and install Outlook add-ins for testing


As part of the process of developing an Outlook add-in, you will probably find yourself iteratively deploying and installing the add-in for testing, which involves the following steps:


1. Creating a manifest file that describes the add-in.
    
2. Deploying the add-in UI file(s) to a web server.
    
3. Installing the add-in in your mailbox.
    
4. Test the add-in, making appropriate changes to the UI or manifest files, and repeating steps 2 and 3 to test the changes.
    

## Creating a manifest file for the add-in

Each add-in is described by an XML manifest, a document that gives the server information about the add-in, provides descriptive information about the add-in for the user, and identifies the location of the add-in UI HTML file. You can store the manifest in a local folder or server, as long as the location is accessible by the Exchange server of the mailbox that you are testing with. We'll assume that you store your manifest in a local folder. For information about how to create a manifest file, see [Outlook add-in manifests](../outlook/manifests/manifests.md). 


## Deploying an add-in to a web server

You can use HTML and JavaScript to create the add-in UI. The resulting source file is stored on a web server that can be accessed by the Exchange server that hosts the add-in. The source file is identified by the  **SourceLocation** child element in the [DesktopSettings](http://msdn.microsoft.com/en-us/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c%28Office.15%29.aspx) element, [TabletSettings](http://msdn.microsoft.com/en-us/library/5c89cc7c-7ae0-49c9-fdd5-4c52118228f6%28Office.15%29.aspx) element, and/or [PhoneSettings](http://msdn.microsoft.com/en-us/library/13e4eae3-8e8c-fd55-a1c2-3297b485f327%28Office.15%29.aspx) element specified in the add-in manifest file.

After initially deploying the UI files for the add-in, you can update the add-in UI and behavior by replacing the HTML file stored on the web server with a new version of the HTML file.


## Installing the add-in


After preparing the add-in manifest file and deploying the add-in UI to a web server that can be accessed, you can install the add-in for a mailbox on an Exchange server by using an Outlook rich client, Outlook Web App, or OWA for Devices, or by running remote Windows PowerShell cmdlets.


### Installing an add-in in an Outlook rich client

You can install an add-in if your mailbox is on Exchange Online, Exchange 2013 or a later release. In Outlook for Windows, you can install add-ins through the Office Fluent Backstage view. Choose **File** and **Manage add-ins**. This allows you to sign in to the Exchange Admin Center. After signing in, continue the installation process with step 4 in the next section.

In Outlook for Mac, choose **Manage add-ins** at the right end of the add-in bar and then sign in to the Exchange Admin Center. Continue with step 4 in the next section.


### Installing an add-in by using Outlook Web App or Outlook.com

To use Outlook Web App (OWA) to install an Outlook add-in, follow these steps:


1. Browse to the OWA URL for your organization or Outlook.com and login.
    
2. Choose the gear icon in the upper-right corner and choose **Manage add-ins**.
    
3. Select the plus sign ( **+**) to add a new add-in.
    
4. From the drop-down list, select **Add from file**, assuming you have stored the manifest on a local folder.
    
5. Browse to the file path of the manifest, and then select **Install**.
    
6. Select the user name in the upper-right corner of the window and select **My Mail** to switch to your email to test the add-in.
    

>**Note**  If you are not using any of the following to develop your add-in: 
- Office 365 developer tenant
- Napa Office 365 Development Tools
- Visual Studio

And, if you do not have at minimum the "My Custom add-ins" role for your Exchange Server, then you can install add-ins only from the Office Store. In order to test your add-in, or install add-ins in general by specifying a URL or file name for the add-in manifest, you should request your Exchange administrator to provide the necessary permissions.

The Exchange administrator can run the following PowerShell cmdlet to assign a single user the necessary permissions. In this example, wendyri is the user's email alias.

**New-ManagementRoleAssignment -Role "My Custom add-ins" -User "wendyri"** 

If necessary, the administrator can run the following cmdlet to assign multiple users the similar necessary permissions:

**$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom add-ins" -User $_.Alias}**

For more information about the My Custom add-ins role, see [My Custom add-ins role](http://technet.microsoft.com/en-us/library/aa0321b3-2ec0-4694-875b-7a93d3d99089%28EXCHG.150%29.aspx). 

Using Office 365, Napa, or Visual Studio to develop add-ins assigns you the organization administrator role which allows you to install add-ins by file or URL in the EAC, or by Powershell cmdlets.


### Installing an add-in by using remote PowerShell

After you create a remote Windows PowerShell session on your Exchange server, you can install an Outlook add-in by using the  **New-App** cmdlet with the following PowerShell command.


```
New-App -URL:"http://<fully-qualified URL">
```

The fully qualified URL is the location of the add-in manifest file that you prepared for your add-in.

You can use the following additional PowerShell cmdlets to manage the add-ins for a mailbox:


-  **Get-App** - Lists the add-ins that are enabled for a mailbox.
    
-  **Set-App** - Enables or disables a add-in on a mailbox.
    
-  **Remove-App** - Removes a previously installed add-in from an Exchange server.
    

## Additional resources



- [Outlook add-ins](../outlook/outlook-add-ins.md)
    
- [Troubleshoot user errors with Office Add-ins](../testing/testing-and-troubleshooting.md)
    
