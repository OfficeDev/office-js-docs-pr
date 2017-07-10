
# Office Add-in commands requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md).

Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. For more information, see [Add-in commands for Excel, Word, and PowerPoint](../../docs/design/add-in-commands.md). 

The initial release of add-in commands doesn't have a corresponding requirement set (that is, there isn't an AddInCommands 1.0 requirement set). The following table lists the Office host applications that support the initial release version, and the build versions or number for those applications.  

| Release   |  Office 2013 for Windows | Office 2016 for Windows*   |  Office 2016 for iPad  |  Office 2016 for Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|
| Add-in commands (initial release, no requirement set) | Not applicable | Version 1603 (Build 6769.0000) or later | Not applicable | 15.33 or later| January 2016 | |

A revision of the add-in commands feature includes a new feature which is the ability to autoopen taskpane with documents. The autoopen feature is included in the add-in commands 1.1 requirement set. For more information about the autoopen taskpane feature, see [Automatically open a task pane with a document](../../docs/add-ins/design/automatically-open-a-task-pane-with-a-document). 

The following table lists the Office Add-in commands 1.1 requirement set, the Office host applications that support that requirement set, and the build or version numbers for the Office application. 

|  Requirement set  |  Office 2013 for Windows | Office 2016 for Windows*   |  Office 2016 for iPad  |  Office 2016 for Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|
| AddInCommands 1.1  | Not applicable | Version 1705 (Build 8121.1000) or later | Not applicable | 15.34 or later| May 2017 | |

>**\*Note:** The build number for Office 2016 installed via MSI is 16.0.4266.1001. In order to use the Office Add-in commands, please run Office update to get the latest version. 

To find out more about versions, build numbers, and Office Online Server, see:

- [Version and build numbers of update channel releases for Office 365 clients](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [What version of Office am I using?](https://support.office.com/en-us/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-US&rs=en-US&ad=US&fromAR=1)
- [Where you can find the version and build number for an Office 365 client application](https://technet.microsoft.com/en-us/library/mt592918.aspx#Anchor_1)
- [Office Online Server overview](https://technet.microsoft.com/en-us/library/jj219437(v=office.16).aspx)

## Office common API requirement sets

For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).

## Additional resources

- [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../docs/overview/add-in-manifests.md)
