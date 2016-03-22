
# Design Office Add-ins for the iPad


The following table lists the tasks to perform to design an Office Add-in to run in Office for iPad.


|**Task**|**Description**|**Resources**|
|:-----|:-----|:-----|
|Update your add-in to support Office.js version 1.1.|Update the JavaScript files (Office.js and app-specific .js files) and the add-in manifest validation file used in your Office Add-in project to version 1.1.|[What's changed in the JavaScript API for Office](../../reference/what's-changed-in-the-javascript-api-for-office.md)|
|Apply UI design best practices.|Integrate your add-in UI seamlessly with the iOS experience.|[Designing for iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)|
|Apply add-in design best practices.|Ensure that your add-in provides clear value, is engaging, and performs consistently.|[Best practices for developing Office Add-ins](../../docs/design/add-in-development-best-practices.md)|
|Optimize your add-in for touch.|Make your UI responsive to touch inputs in addition to mouse and keyboard.|[Apply UX design principles](https://msdn.microsoft.com/EN-US/library/mt590883.aspx#Anchor_3)|
|Make your add-in free.|Office on iPad is a channel through which you can reach more users and promote your services. These new users have the potential to become your customers.|[Validation policy 10.8](http://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)|
|Make your add-in commerce free.|Your add-in must be free of in-app purchases, trial offers, UI that aims to upsell to paid or links to any online stores where users can purchase or acquire other content, apps, or add-ins.Your Privacy Policy and Terms of Use pages must also be free of any commerce UI or Store links.|[Validation policy 3.4](http://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)|
|Resubmit your add-in to the Office Store.|In the Seller Dashboard, select the  **Make this add-in available in the Office Add-in Catalog on iPad** check box, and provide your Apple developer ID in the Apple ID box. Review the[Office Store Application Provider Agreement](https://sellerdashboard.microsoft.com/Assets/Content/Agreements/en-US/Office_Store_Seller_Agreement_20120927.md) to make sure you understand agreement.|[Submit Office and SharePoint Add-ins and Office 365 web apps to the Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)|
Your add-in can remain as-is for Office applications that are running on other platforms. You can also serve a different UI based on the browser/device that your add-in is running on. To detect whether your add-in is running on an iPad, you can use the following APIs: 

- var isTouchEnabled = [Office.context.touchEnabled](../../reference/shared/office.context.touchenabled.md)
    
- var allowCommerce = [Office.context.commerceAllowed](../../reference/shared/office.context.commerceallowed.md)
    

## Best practices for developing Office Add-ins for iOS and Mac

Apply the following best practices for developing add-ins that run on iOS:


-  **Use Visual Studio to develop your add-in.**
    
    If you develop your add-in with Visual Studio, you can [set breakpoints and debug its code](../../docs/get-started/create-and-debug-office-add-ins-in-visual-studio.md#Test) in an Office host application running on Windows, before sideloading your add-in on the iPad or Mac. Because an add-in that runs in Office for iOS or Office for Mac supports the same APIs as an add-in running in Office for Windows, your add-in's code should run the same way on both platforms.
    
-  **Specify API requirements in your add-in's manifest or with runtime checks.**
    
    When you specify API requirements in your add-in's manifest, Office will determine if the host application supports those API members. If the API members are available in the host, then your add-in will be available in that host application. Alternatively, you can perform a runtime check to determine if a method is available in the host before using it in your add-in. Runtime checks ensure that your add-in is always available in the host, and provides additional functionality if the methods are available. For more information, see [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md).
    
For general add-in development best practices, see [Best practices for developing Office Add-ins](../../docs/design/add-in-development-best-practices.md).


## Additional resources
<a name="bk_addresources"> </a>


- [Sideload an Office Add-in on iPad and Mac](../../docs/testing/sideload-an-office-add-in-on-ipad-and-mac.md)
    
- [Debug Office Add-ins on iPad and Mac](../../docs/testing/debug-office-add-ins-on-ipad-and-mac.md)
    
