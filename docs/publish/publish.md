
# Deploy and publish your Office Add-in


You can use one of several methods to deploy your Office Add-in for testing or distribution to users:

- [Sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) - Use as part of your development process to test your add-in running on Windows, Office Online, iPad, or Mac.
- [SharePoint catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) - Use as part of your development process to test your add-in, or to distribute your add-in to users in your organization.
- [Office 365 admin center preview](https://support.office.com/en-ie/article/Deploy-Office-Add-Ins-in-Office-365-737e8c86-be63-44d7-bf02-492fa7cd9c3f?ui=en-US&rs=en-IE&ad=IE) - Use to distribute your add-in to users in your organization.
- [Office Store] - Use to distribute your add-in publicly to users.

The options that are available depend on the Office host that you're targeting and the type of add-in you create.

### Deployment Options for Word, Excel, and PowerPoint Add-ins

| Extension point            | Sideloading | SharePoint catalog | Office 365 admin center preview | Office Store |
|:----------------|:-----------:|:------------------:|:-------------------------------:|:------------:|
| Content         | X           | X                  | X                               | X            |
| Task pane       | X           | X                  | X                               | X            |
| Command 		  | X           |                    | X                               | X            |

> **NOTE:** SharePoint catalogs are not supported for Office 2016 for Mac. To deploy Office Add-ins to Mac clients, you must submit them to the [Office Store].    

### Deployment Options for Outlook Add-ins

| Extension point     | Sideloading | Exchange server | Office Store |
|:---------|:-----------:|:---------------:|:------------:|
| Mail App | X           | X               | X            |
| Command  | X           | X               | X            |

To broaden the reach of your add-in, make sure that it works across platforms. Office Add-ins are supported on Windows, Mac, Web, iOS and Android. For an overview of which features are supported by each platform, see [Office Add-in host and platform availability].   

For information about licensing your Office Store add-ins, see [License your add-ins](https://msdn.microsoft.com/EN-US/library/office/jj163257.aspx).

For information about how end users acquire, insert, and run add-ins, see [Start using your Office Add-in](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).

## Additional resources

- [Office Add-in host and platform availability]
- [Deploy and install Outlook add-ins for testing](../outlook/testing-and-tips.md) 
- [Submit add-ins and web apps to the Office Store][Office Store]
- [Design guidelines for Office Add-ins](../design/add-in-design)
- [Created effective Office Store add-ins](https://msdn.microsoft.com/en-us/library/jj635874.aspx)
- [Troubleshoot user errors with Office Add-ins](../testing/testing-and-troubleshooting.md)

[Office Store]: http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx
[Office Add-in host and platform availability]: http://dev.office.com/add-in-availability
