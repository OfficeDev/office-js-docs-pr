---
title: Develop Office Add-ins for the iPad
description: 'Get an overview and best practices for creating an Office Add-in that runs on an iPad.'
ms.date: 03/18/2020
localization_priority: Normal
---


# Develop Office Add-ins for the iPad


The following table lists the tasks to perform to develop an Office Add-in to run in Office on iPad.


|**Task**|**Description**|**Resources**|
|:-----|:-----|:-----|
|Update your add-in to support Office.js version 1.1.|Update the JavaScript files (Office.js and app-specific .js files) and the add-in manifest validation file used in your Office Add-in project to version 1.1.|[Update API and manifest version](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|Apply UI design best practices.|Integrate your add-in UI seamlessly with the iOS experience.|[Designing for iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)|
|Apply add-in design best practices.|Ensure that your add-in provides clear value, is engaging, and performs consistently.|[Best practices for developing Office Add-ins](../concepts/add-in-development-best-practices.md)|
|Optimize your add-in for touch.|Make your UI responsive to touch inputs in addition to mouse and keyboard.|[Apply UX design principles](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|Make your add-in free.|Office on iPad is a channel through which you can reach more users and promote your services. These new users have the potential to become your customers.|[Certification policy 1120.2](/legal/marketplace/certification-policies#11202-acquisition-pricing-and-terms)|
|Make your add-in commerce free.|Your add-in must be free of in-app purchases, trial offers, UI that aims to upsell to paid or links to any online stores where users can purchase or acquire other content, apps, or add-ins. Your Privacy Policy and Terms of Use pages must also be free of any commerce UI or AppSource links.|[Certification policy 1100.3](/legal/marketplace/certification-policies#11003-selling-additional-features)|
|Resubmit your add-in to AppSource.|In Partner Center, on the **Product setup** page, select the **Make my product available on iOS and Android (if applicable)** check box, and provide your Apple developer ID in Account settings. Review the [Application Provider Agreement](https://go.microsoft.com/fwlink/?linkid=715691) to make sure you understand the terms.|[Make your solutions available in AppSource and within Office](/office/dev/store/submit-to-appsource-via-partner-center)|

Your add-in can remain as-is for Office applications that are running on other platforms. You can also serve a different UI based on the browser/device that your add-in is running on. To detect whether your add-in is running on an iPad, you can use the following APIs:
- var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#touchenabled)
- var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#commerceallowed)


## Best practices for developing Office Add-ins for iOS and Mac

Apply the following best practices for developing add-ins that run on iOS:


-  **Use Visual Studio to develop your add-in.**

    If you develop your add-in with Visual Studio, you can [set breakpoints and debug its code](../develop/debug-office-add-ins-in-visual-studio.md) in an Office client application running on Windows, before sideloading your add-in on the iPad or Mac. Because an add-in that runs in Office on iOS or Mac supports the same APIs as an add-in running in Office on Windows, your add-in's code should run the same way on both platforms.

-  **Specify API requirements in your add-in's manifest or with runtime checks.**

    When you specify API requirements in your add-in's manifest, Office will determine if the Office client application supports those API members. If the API members are available in the application, then your add-in will be available. Alternatively, you can perform a runtime check to determine if a method is available in the application before using it in your add-in. Runtime checks ensure that your add-in is always available in the application, and provides additional functionality if the methods are available. For more information, see [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md).

For general add-in development best practices, see [Best practices for developing Office Add-ins](../concepts/add-in-development-best-practices.md).


## See also

- [Sideload an Office Add-in on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)  
- [Debug Office Add-ins on iPad and Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)
