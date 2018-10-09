# Add-in commands requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. For more information, see [Add-in commands for Excel, Word, and PowerPoint](https://docs.microsoft.com/office/dev/add-ins/design/add-in-commands) and [Add-in commands for Outlook](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook).

The initial release of add-in commands doesn't have a corresponding requirement set (that is, there isn't an AddinCommands 1.0 requirement set). The following table lists the Office host applications that support the initial release version, and the build versions or number for those applications.  

| Release   |  Office 2013 for Windows | Office 2016 for Windows (non-subscription) | Office 365 for Windows   |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Add-in commands (initial release, no requirement set) | N/A | 16.0.4678.1000 *Supported in Outlook only* |Version 1603 (Build 6769.0000) or later | N/A | 15.33 or later| January 2016 | |

The add-in commands 1.1 requirement set introduces the ability to [autoopen a task pane with documents](https://docs.microsoft.com/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document).

The following table lists the add-in commands 1.1 requirement set, the Office host applications that support that requirement set, and the build or version numbers for the Office application. 

|  Requirement set  |  Office 2013 for Windows | Office 2016 for Windows (non-subscription) | Office 365 for Windows   |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddinCommands 1.1  | N/A | 16.0.4678.1000 *Supported in Outlook only*  | Version 1705 (Build 8121.1000) or later | N/A | 15.34 or later| May 2017 | |

To find out more about versions, build numbers, and Office Online Server, see:

- [Version and build numbers of update channel releases for Office 365 clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [What version of Office am I using?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Where you can find the version and build number for an Office 365 client application](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server overview](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## Office common API requirement sets

For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).

## See also

- [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office hosts and API requirements](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
