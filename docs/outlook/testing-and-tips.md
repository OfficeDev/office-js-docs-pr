---
title: Deploy and install Outlook add-ins for testing
description: Create a manifest file, deploy the add-in UI file to a web server, install the add-in in your mailbox, and then test the add-in.
ms.date: 05/20/2020
localization_priority: Priority
---

# Deploy and install Outlook add-ins for testing

As part of the process of developing an Outlook add-in, you will probably find yourself iteratively deploying and installing the add-in for testing, which involves the following steps.

1. Creating a manifest file that describes the add-in.
1. Deploying the add-in UI file(s) to a web server.
1. Installing the add-in in your mailbox.
1. Testing the add-in, making appropriate changes to the UI or manifest files, and repeating steps 2 and 3 to test the changes.

> [!NOTE]
> [Custom panes have been deprecated](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/) so please ensure that you're using [a supported add-in extension point](outlook-add-ins-overview.md#extension-points).

## Create a manifest file for the add-in

Each add-in is described by an XML manifest, a document that gives the server information about the add-in, provides descriptive information about the add-in for the user, and identifies the location of the add-in UI HTML file. You can store the manifest in a local folder or server, as long as the location is accessible by the Exchange server of the mailbox that you are testing with. We'll assume that you store your manifest in a local folder. For information about how to create a manifest file, see [Outlook add-in manifests](manifests.md).

## Deploy an add-in to a web server

You can use HTML and JavaScript to create the add-in. The resulting source files are stored on a web server that can be accessed by the Exchange server that hosts the add-in. After initially deploying the source files for the add-in, you can update the add-in UI and behavior by replacing the HTML files or JavaScript files stored on the web server with a new version of the HTML file.

## Install the add-in

After preparing the add-in manifest file and deploying the add-in UI to a web server that can be accessed, you can sideload the add-in for a mailbox on an Exchange server by using an Outlook client, or install the add-in by running remote Windows PowerShell cmdlets.

### Sideload the add-in

You can install an add-in if your mailbox is on Exchange Online, Exchange 2013 or a later release. Sideloading add-ins requires at minimum the **My Custom Apps** role for your Exchange Server. In order to test your add-in, or install add-ins in general by specifying a URL or file name for the add-in manifest, you should request your Exchange administrator to provide the necessary permissions.

The Exchange administrator can run the following PowerShell cmdlet to assign a single user the necessary permissions. In this example, `wendyri` is the user's email alias.

```powershell
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
```

If necessary, the administrator can run the following cmdlet to assign multiple users the similar necessary permissions.

```powershell
$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
```

For more information about the My Custom Apps role, see [My Custom Apps role](/exchange/my-custom-apps-role-exchange-2013-help).

Using Microsoft 365 or Visual Studio to develop add-ins assigns you the organization administrator role which allows you to install add-ins by file or URL in the EAC, or by Powershell cmdlets.

### Install an add-in by using remote PowerShell

After you create a remote Windows PowerShell session on your Exchange server, you can install an Outlook add-in by using the `New-App` cmdlet with the following PowerShell command.

```powershell
New-App -URL:"http://<fully-qualified URL">
```

The fully qualified URL is the location of the add-in manifest file that you prepared for your add-in.

You can use the following additional PowerShell cmdlets to manage the add-ins for a mailbox.

- `Get-App` - Lists the add-ins that are enabled for a mailbox.
- `Set-App` - Enables or disables a add-in on a mailbox.
- `Remove-App` - Removes a previously installed add-in from an Exchange server.

## Client versions

Deciding what versions of the Outlook client to test depends on your development requirements.

- If you are developing an add-in for private use, or only for members of your organization, then it is important to test the versions of Outlook that your company uses. Keep in mind that some users may use Outlook on the web, so testing your company's standard browser versions is also important.

- If you are developing an add-in to list in [AppSource](https://appsource.microsoft.com), you must test the required versions as specified in the [Commercial marketplace certification policies 1120.3](/legal/marketplace/certification-policies#11203-functionality). This includes:
  - The latest version of Outlook on Windows and the version prior to the latest.
  - The latest version of Outlook on Mac.
  - The latest version of Outlook on iOS and Android (if your add-in [supports mobile form factor](add-mobile-support.md)).
  - The browser versions specified in the Commercial marketplace validation policy 1120.3.

> [!NOTE]
> If your add-in does not support one of the above clients due to [requesting an API requirement set](apis.md) that the client does not support, that client would be removed from the list of required clients.

## Outlook on the web and Exchange server versions

Consumer and Microsoft 365 account users see the modern UI version when they access Outlook on the web and no longer see the classic version which has been deprecated. However, on-premises Exchange servers continue to support classic Outlook on the web. Therefore, during the validation process, your submission may receive a warning that the add-in is not compatible with classic Outlook on the web. In that case, you should consider testing your add-in in an on-premises Exchange environment. This warning won't block your submission to AppSource but your customers may experience a sub-optimal experience if they use Outlook on the web in an on-premises Exchange environment.

To mitigate this, we recommend you test your add-in in Outlook on the web connected to your own private on-premises Exchange environment. For more information, see guidance on how to [Establish an Exchange 2016 or Exchange 2019 test environment](/Exchange/plan-and-deploy/plan-and-deploy?view=exchserver-2019&preserve-view=true#establish-an-exchange-2016-or-exchange-2019-test-environment) and how to manage [Outlook on the web in Exchange Server](/exchange/clients/outlook-on-the-web/outlook-on-the-web?view=exchserver-2019&preserve-view=true).

Alternatively, you can opt to pay for and use a service that hosts and manages on-premises Exchange servers. A couple of options are:

- [Rackspace](https://www.rackspace.com/email-hosting/exchange-server)
- [Hostway](https://hostway.com/products-services-2/hosted-microsoft-exchange/)

Furthermore, if you don't want your add-ins to be available for users who are connected to on-premises Exchange, you can set the [requirement set](../reference/requirement-sets/outlook-api-requirement-sets.md#exchange-server-support) in the add-in manifest to be 1.6 or higher. Such add-ins will not be tested or validated on the classic Outlook on the web UI.

## See also

- [Troubleshoot user errors with Office Add-ins](../testing/testing-and-troubleshooting.md)
