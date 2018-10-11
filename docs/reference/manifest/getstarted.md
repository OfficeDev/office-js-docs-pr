# GetStarted element

Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote hosts. The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).

## Child elements

| Element                       | Required | Description                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Title](#title)               | Yes      | Defines where an add-in exposes functionality.     |
| [Description](#description)   | Yes      | A URL to a file that contains JavaScript functions.|
| [LearnMoreUrl](#learnmoreurl) | No       | A URL to a page that explains the add-in in detail.   |

### Title 

Required. The title used for the top of the callout. The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section.

### Description

Required. The description / body content for the callout. The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section.

### LearnMoreUrl

Required. The URL to a page where the user can learn more about your add-in. The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section.

> [!NOTE]
> **LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients. We recommend that you add this URL for all clients so that the URL will render when it becomes available. 

## See also

The following code samples use the **GetStarted** element:

* [Excel Web Add-in for Manipulating Table and Chart Formatting](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [Word Add-in JavaScript SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
