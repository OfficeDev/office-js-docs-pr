# CustomPane

The CustomPane [extension point](./extensionpoint.md) defines an add-in that activates when specified rules are satisfied. It is only for read form and it displays in a horizontal pane. 

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [RequestedHeight](#requestedheight) | No |  The requested height in pixels.  |
|  [SourceLocation](#sourcelocation)  | Yes |  The URL for the source code file of the add-in.  |
|  [Rule](#rule)  | Yes |  The rule or collection of rules that specify when the add-in activates.  |
|  [DisableEntityHighlighting](#disableentityhighlighting)  | No |  Specifies whether entity highlighting should be turned off. |

## RequestedHeight
Optional. The requested height, in pixels, for the display pane when it is running on a desktop computer. This can be from 32 to 450 pixels. It is the same as in read add-ins (see [RequestedHeight element](../reference/requestedheight.md)

## SourceLocation
Required. The URL for the source code file of the add-in. This refers to a  **Url** element in the [Resources](./resources.md)  element.

## Rule
Required. The rule or collection of rules that specify when the add-in activates. It is the same as defined in [Outlook add-in manifests](../../outlook/manifests/manifests.md), except the ItemIs rule has the following changes: **ItemType** is either "Message" or "AppointmentAttendee", and there is no **FormType** attribute. For more information, see [Custom pane Outlook add-ins](../../outlook/custom-pane-outlook-add-ins.md) and [Activation rules for Outlook add-ins](../../outlook/manifests/activation-rules.md).

## DisableEntityHighlighting
Optional. Specifies whether entity highlighting should be turned off for this Outlook add-in. 

## CustomPane example
```xml
<ExtensionPoint xsi:type="CustomPane">
   <RequestedHeight>100< /RequestedHeight> 
   <SourceLocation resid="residReadTaskpaneUrl"/>
   <Rule xsi:type="RuleCollection" Mode="Or">
     <Rule xsi:type="ItemIs" ItemType="Message"/>
     <Rule xsi:type="ItemHasAttachment"/>
     <Rule xsi:type="ItemHasKnownEntity" EntityType="Address"/>
   </Rule>
</ExtensionPoint>
```