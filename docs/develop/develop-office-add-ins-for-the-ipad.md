---
title: Develop Office Add-ins for the iPad
description: ''
ms.date: 01/23/2018
---


# Develop Office Add-ins for the iPad


The following table lists the tasks to perform to develop an Office Add-in to run in Office for iPad.


|**Task**|**Description**|**Resources**|
|:-----|:-----|:-----|
|Update your add-in to support Office.js version 1.1.|Update the JavaScript files (Office.js and app-specific .js files) and the add-in manifest validation file used in your Office Add-in project to version 1.1.|[What's changed in the JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/what's-changed-in-the-javascript-api-for-office?view=office-js)|
|Apply UI design best practices.|Integrate your add-in UI seamlessly with the iOS experience.|[Designing for iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)|
|Apply add-in design best practices.|Ensure that your add-in provides clear value, is engaging, and performs consistently.|[Best practices for developing Office Add-ins](../concepts/add-in-development-best-practices.md)|
|Optimize your add-in for touch.|Make your UI responsive to touch inputs in addition to mouse and keyboard.|[Apply UX design principles](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|Make your add-in free.|Office on iPad is a channel through which you can reach more users and promote your services. These new users have the potential to become your customers.|[Validation policy 10.8](https://docs.microsoft.com/office/dev/store/validation-policies#10-apps-and-add-ins-utilize-supported-capabilities)|
|Make your add-in commerce free.|Your add-in must be free of in-app purchases, trial offers, UI that aims to upsell to paid or links to any online stores where users can purchase or acquire other content, apps, or add-ins. Your Privacy Policy and Terms of Use pages must also be free of any commerce UI or AppSource links.|[Validation policy 3.4](https://docs.microsoft.com/office/dev/store/validation-policies#3-apps-and-add-ins-can-sell-additional-features-or-content-through-purchases-within-the-app-or-add-in)|
|Resubmit your add-in to AppSource.|In the Seller Dashboard, select the **Make this add-in available in the Office Add-in Catalog on iPad** check box, and provide your Apple developer ID in the Apple ID box. Review the [Application Provider Agreement](https://sellerdashboard.microsoft.com/Assets/Content/Agreements/en-US/Office_Store_Seller_Agreement_20120927.htm) to make sure you understand agreement.|[Make your solutions available in AppSource and within Office](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)|

Your add-in can remain as-is for Office applications that are running on other platforms. You can also serve a different UI based on the browser/device that your add-in is running on. To detect whether your add-in is running on an iPad, you can use the following APIs:
- var isTouchEnabled = [Office.context.touchEnabled](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#touchenabled)
- var allowCommerce = [Office.context.commerceAllowed](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#commerceallowed)
    

## Best practices for developing Office Add-ins for iOS and Mac

Apply the following best practices for developing add-ins that run on iOS:


-  **Use Visual Studio to develop your add-in.**
    
    If you develop your add-in with Visual Studio, you can [set breakpoints and debug its code](../develop/create-and-debug-office-add-ins-in-visual-studio.md) in an Office host application running on Windows, before sideloading your add-in on the iPad or Mac. Because an add-in that runs in Office for iOS or Office for Mac supports the same APIs as an add-in running in Office for Windows, your add-in's code should run the same way on both platforms.
    
-  **Specify API requirements in your add-in's manifest or with runtime checks.**
    
    When you specify API requirements in your add-in's manifest, Office will determine if the host application supports those API members. If the API members are available in the host, then your add-in will be available in that host application. Alternatively, you can perform a runtime check to determine if a method is available in the host before using it in your add-in. Runtime checks ensure that your add-in is always available in the host, and provides additional functionality if the methods are available. For more information, see [Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md).
    
For general add-in development best practices, see [Best practices for developing Office Add-ins](../concepts/add-in-development-best-practices.md).


## See also

- [Sideload an Office Add-in on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Debug Office Add-ins on iPad and Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)
    
