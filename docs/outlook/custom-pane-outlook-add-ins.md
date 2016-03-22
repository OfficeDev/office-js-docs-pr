
# Custom pane Outlook add-ins

A custom pane is an extension point for an add-in that activates when specific conditions are satisfied on the currently selected item. It is defined in the add-in manifest in the  **VersionOverrides** element, along with any add-in commands implemented by the add-in. For more information, see [Define add-in commands in your Outlook add-in manifest](../outlook/manifests/define-add-in-commands.md).
A custom pane can only appear in a message read or appointment attendee views. It displays an entry in the add-in bar. When the user clicks the entry, the custom pane shows with a horizontal orientation above the body of the item. The appearance and behavior is the same as with read mode add-ins that do not implement add-in commands.

**An add-in with a custom pane in read mode**

![Shows a custom pane in a message read form.](../../images/c585ab0a-6c33-42d0-a20f-5deb8b54f480.png)

The following example defines a custom pane for items that are messages or have an attachment or include an address. 



```
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



-  **RequestedHeight** specifies the desired height, in pixels, of this mail add-in when running on a desktop computer. It is ignored otherwise. It can be a value between 32 and 450. If it is not set, the default is 350 px. Optional.
    
-  **SourceLocation** specifies the HTML page that provides the UI for the custom pane. The **resid** attribute is set to the value of the **id** attribute of a **Url** element in the **Resources** element. Required.
    
-  **Rule** specifies the rule or collection of rules that specify when the add-in activates. It is the same as defined in [Outlook add-in manifests](../outlook/manifests/manifests.md), except the [ItemIs](http://msdn.microsoft.com/en-us/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx) rule has the following changes: **ItemType** is either "Message" or "AppointmentAttendee", and there is no **FormType** attribute. For more information, see [Activation rules for Outlook add-ins](../outlook/manifests/activation-rules.md).
    

## Additional resources



- [Get Started with Outlook add-ins for Office 365](https://dev.outlook.com/MailAppsGettingStarted/GetStarted.aspx)
    
- [Activation rules for Outlook add-ins](../outlook/manifests/activation-rules.md)
    
- [Outlook add-in manifests](../outlook/manifests/manifests.md)
    
- [Define add-in commands in your Outlook add-in manifest](../outlook/manifests/define-add-in-commands.md)
    
