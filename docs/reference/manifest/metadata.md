# Metadata element

Defines the metadata settings used by a custom function in Excel.

## Attributes

None

## Child elements

|  Element  |  Required  |  Description  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Yes  | String with the resource id of the JSON file used by custom functions. |

## Example

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
