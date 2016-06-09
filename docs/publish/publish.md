
# Deploy and publish your Office Add-in


You can use one of several methods to deploy your Office Add-in for testing or distribution to users:

- [Sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) - Use as part of your development process to test your add-in running on Windows, Office Online, iPad, or Mac.
- [SharePoint catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) - Use as part of your development process to test your add-in, or to distribute your add-in to users in your organization.
- [Office 365 admin center preview](https://support.office.com/en-ie/article/Deploy-Office-Add-Ins-in-Office-365-737e8c86-be63-44d7-bf02-492fa7cd9c3f?ui=en-US&rs=en-IE&ad=IE) - Use to distribute your add-in to users in your organization.
- [Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx) - Use to distribute your add-in publicly to users.

The options that are available depend on the type of add-in you create. 

**Options to deploy and publish Word, Excel, and PowerPoint add-ins**

|**Extension point**|**Sideloading**|**SharePoint catalog**|**Office 365 admin center preview**|**Office Store**|
|:-----|:-----|:-----|:-----|:-----|
|Command|X||X|X|
|Content|X|X|X|X|
|Task pane|X|X|X|X|

**Options to deploy and publish Outlook add-ins**

|**Extension point**|**Exchange server**|**Office Store**|
|:-----|:-----|:-----|
|Command|X|X|
|Read/compse panes|X|X|

To broaden the reach of your add-in, make sure that it works across platforms. The Office.js version 1.1 includes support for Office Online, and the Office Store validation process verifies add-in support for Office Online. 

For information about how end users acquire, insert, and run add-ins, see [Start using your Office Add-in](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).

## Additional resources

- [Deploy and install Outlook add-ins for testing](../outlook/testing-and-tips.md) 
- [Submit add-ins and web apps to the Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)
- [Create effective Office Store add-ins](https://msdn.microsoft.com/library/jj635874.asp) 
- [Troubleshoot user errors with Office Add-ins](../testing/testing-and-troubleshooting.md)


    


