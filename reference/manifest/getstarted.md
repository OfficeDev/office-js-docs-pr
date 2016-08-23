# GetStarted element

 GetStarted provides information used by the callout that appears when installing the add-in on Word, Excel and PowerPoint hosts. The **GetStarted** element is a child element of [FormFactor](./formfactor.md).

 ## Child elements

| Element                       | Required | Description                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Title](#title)               | Yes      | Defines where an add-in exposes functionality.     |
| [Description](#description)   | Yes      | A URL to a file that contains JavaScript functions.|
| [LearnMoreUrl](#learnmoreurl) | No       | A URL to a page explaining the add-in in detail.   |


## Title 
Required. The title used for the top of the callout. The **resid** attribute reference a valid ID in the [ShortStrings](./resources.md#shortstrings) from the [Resources](./resources.md) section.

## Description
Required. The description / body content for the callout. The **resid** attribute reference a valid ID in the [LongStrings](./resources.md#longstrings) from the [Resources](./resources.md) section.

## LearnMoreUrl
Required. The URL to a page where the end-user can learn more about your add-in. The **resid** attribute reference a valid ID in the [Urls](./resources.md#urls) from the [Resources](./resources.md) section.

> **NOTE:** At this time LearnMoreUrl is not rendered in Word, Excel or PowerPoint clients. We do however recommend adding this URL so that it will automatically begin rendering when this feature is rolled out to customers. 
