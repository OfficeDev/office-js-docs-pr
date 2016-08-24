# GetStarted element

Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote hosts. The **GetStarted** element is a child element of [FormFactor](./formfactor.md).

 ## Child elements

| Element                       | Required | Description                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Title](#title)               | Yes      | Defines where an add-in exposes functionality.     |
| [Description](#description)   | Yes      | A URL to a file that contains JavaScript functions.|
| [LearnMoreUrl](#learnmoreurl) | No       | A URL to a page that explains the add-in in detail.   |


## Title 
Required. The title used for the top of the callout. The **resid** attribute references a valid ID in the [ShortStrings](./resources.md#shortstrings) element in the [Resources](./resources.md) section.

## Description
Required. The description / body content for the callout. The **resid** attribute references a valid ID in the [LongStrings](./resources.md#longstrings) element in the [Resources](./resources.md) section.

## LearnMoreUrl
Required. The URL to a page where the user can learn more about your add-in. The **resid** attribute references a valid ID in the [Urls](./resources.md#urls) element in the [Resources](./resources.md) section.

> **NOTE:** **LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients. We recommend that you add this URL for all clients so that the URL will render when it becomes available. 
