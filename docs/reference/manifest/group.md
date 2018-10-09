# Group element

Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  Yes  | A unique ID for the group.|

### id attribute

Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.

## Child elements
|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Label](#label)      | Yes |  The label for the CustomTab or a group.  |
|  [Control](#control)    | Yes |  Collection of one or more Control objects.  |

### Label 

Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.

### Control
A group requires at least one control.

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```