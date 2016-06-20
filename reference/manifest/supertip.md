## Supertip
Defines a rich tooltip (both Title and Description). It is used by both [Button](./button.md) and [Menu](./menu-control.md) controls. 

## Child elements
|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Title](#title)        | Yes |   The text for the supertip.         |
|  [Description](#description)  | Yes |  The description for the supertip.    |

## Title
Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the [ShortStrings](./resources.md#shortstrings) element in the [Resources](./resources.md) element.

## Description
Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the [LongStrings](./resources.md#longstrings) element in the [Resources](./resources.md) element.

```xml
 <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
```