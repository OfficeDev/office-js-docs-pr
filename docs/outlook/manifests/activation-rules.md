
# Activation rules for Outlook add-ins


Outlook activates some types of add-ins if the message or appointment that the user is reading or composing satisfies the activation rules of the add-in. This is true for all add-ins that use the 1.1 manifest schema and for custom pane addins. The user can then choose the add-in from the Outlook UI to start it for the current item.

The following figure shows Outlook add-ins activated in the add-in bar for the message in the Reading Pane. 

![App bar showing activated read mail apps](../../../images/mod_off15_MailAppAppBar.png)


## Specify activation rules in a manifest


To have Outlook activate an add-in for specific conditions, specify activation rules in the add-in manifest by using one of the following **Rule** elements:

- [Rule element (MailApp complexType)](http://msdn.microsoft.com/en-us/library/56dfc32e-2b8c-1724-05be-5595baf38aa3%28Office.15%29.aspx) - Specifies an individual rule.
- [Rule element (RuleCollection complexType)](http://msdn.microsoft.com/en-us/library/c6ce9d52-4b53-c6a6-de7e-c64106135c81%28Office.15%29.aspx) - Combines multiple rules using logical operations.
    

 >**Note**  The [Rule](http://msdn.microsoft.com/en-us/library/56dfc32e-2b8c-1724-05be-5595baf38aa3%28Office.15%29.aspx) element that you use to specify an individual rule is of the abstract [Rule](http://msdn.microsoft.com/en-us/library/bcd7a3a7-9cd4-a270-89e0-5386d1c6df01%28Office.15%29.aspx) complex type. Each of the following types of rules extends this abstract **Rule** complex type. So when you specify an individual rule in a manifest, you must use the [xsi:type](http://www.w3.org/TR/xmlschema-1/) attribute to further define one of the following types of rules. For example, the following rule defines an [ItemIs](http://msdn.microsoft.com/en-us/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx) rule: `<Rule xsi:type="ItemIs" ItemType="Message" />`Note: The  **FormType** attribute applies to activation rules in the manifest v1.1 but is not defined in **VersionOverrides** v1.0. So it can't be used when [ItemIs](http://msdn.microsoft.com/en-us/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx) is used in the **VersionOverrides** node.

The following table lists the types of rules that are available. You can find more information following the table and in the specified articles under [Create Outlook add-ins for read forms](../../outlook/read-scenario.md).


|**Rule name**|**Applicable forms**|**Description**|
|:-----|:-----|:-----|
|[ItemIs](http://msdn.microsoft.com/en-us/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx)|Read ,Compose, Custom pane|Checks to see whether the current item is of the specified type (message or appointment). Can also check the item class and form type.and optionally, item message class.|
|[ItemHasAttachment](http://msdn.microsoft.com/en-us/library/031db7be-8a25-5185-a9c3-93987e10c6c2%28Office.15%29.aspx)|Read, Custom pane|Checks to see whether the selected item contains an attachment.|
|[ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx)|Read, Custom pane|Checks to see whether the selected item contains one or more well-known entities. More information: [Match strings in an Outlook item as well-known entities](../../outlook/match-strings-in-an-item-as-well-known-entities.md).|
|[ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx)|Read, Custom pane|Checks to see whether the sender's email address, the subject, and/or the body of the selected item contains a match to a regular expression.More information: [Use regular expression activation rules to show an Outlook add-in](../../outlook/use-regular-expressions-to-show-an-outlook-add-in.md).|
|[RuleCollection](http://msdn.microsoft.com/en-us/library/926249ab-2d2f-39f5-1d73-fab1c989966f%28Office.15%29.aspx)|Read, Compose, Custom pane|Combines a set of rules so that you can form more complex rules.|

## ItemIs rule


The  **ItemIs** complex type defines a rule that evaluates to **true** if the current item matches the item type, and optionally the item message class if it's stated in the rule.

Specify one of the following item types in the  **ItemType** attribute of an **ItemIs** rule. You can specify more than one **ItemIs** rule in a manifest. The [ItemType](http://msdn.microsoft.com/en-us/library/5a890b98-3d83-77ef-ef03-9b513d35b79f%28Office.15%29.aspx) simpleType defines the types of Outlook items that support Outlook add-ins.



|**ItemType value**|**Description**|
|:-----|:-----|
|**Appointment**|Specifies an item in an Outlook calendar. This includes a meeting item that has been responded to and has an organizer and attendees, or an appointment that does not have an organizer or attendee and is simply an item on the calendar.This corresponds to the IPM.Appointment message class in Outlook.|
|**Message**|Specifies one of the following items received in typically the Inbox: <ul><li><p>An email message. This corresponds to the IPM.Note message class in Outlook.</p></li><li><p>A meeting request, response, or cancellation. This corresponds to the following  message classes in Outlook:</p><p>IPM.Schedule.Meeting.Request</p><p>IPM.Schedule.Meeting.Neg</p><p>IPM.Schedule.Meeting.Pos</p><p>IPM.Schedule.Meeting.Tent</p><p>IPM.Schedule.Meeting.Canceled</p></li></ul>|
The  **FormType** attribute is used to specify the mode (read or compose) in which the add-in should activate.


 >**Note**  The [ItemIs](http://msdn.microsoft.com/en-us/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx) **FormType** attribute is defined in schema v1.1 and later but not in **VersionOverrides** v1.0. Do not include the **FormType** attribute when defining rules for a custom pane.

After an add-in is activated, you can use the [mailbox.item](../../../reference/outlook/Office.context.mailbox.item.md) property to obtain the currently selected item in Outlook, and the [item.itemType](../../../reference/outlook/Office.context.mailbox.item.md) property to obtain the type of the current item.

You can optionally use the  **ItemClass** attribute to specify the message class of the item, and the **IncludeSubClasses** attribute to specify whether the rule should be **true** when the item is a subclass of the specified class.

For more information about message classes, see [Item Types and Message Classes](http://msdn.microsoft.com/library/15b709cc-7486-b6c7-88a3-4a4d8e0ab292%28Office.15%29.aspx).

The following example is an  **ItemIs** rule that lets users see the add-in in the Outlook add-in bar when the user is reading a message:




```
<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
```

The following example is an  **ItemIs** rule that lets users see the add-in in the Outlook add-in bar when the user is reading a message or appointment.




```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
</Rule>
```


## ItemHasAttachment rule


The  **ItemHasAttachment** complex type defines a rule that checks if the selected item contains an attachment.


```XML
<Rule xsi:type="ItemHasAttachment" />
```


## ItemHasKnownEntity rule


Before an item is made available to an add-in, the server examines it to determine whether the subject and body contain any text that is likely to be one of the known entities. If any of these entities are found, it is placed in a collection of known entities that you access by using the  **getEntities** or **getEntitiesByType** method of that item.

You can specify a rule by using the [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) complex type that shows your add-in when an entity of the specified type is present in the item. You can specify the following known entities in the **EntityType** attribute of an **ItemHasKnownEntity** rule:


-  **Address**
-  **Contact**
-  **EmailAddress**
-  **MeetingSuggestion**
-  **PhoneNumber**
-  **TaskSuggestion**
-  **URL**
    
These entities are defined as enumerated values in the [KnownEntityType](http://msdn.microsoft.com/en-us/library/432d413b-9fcc-eb50-cfea-0ed10a43bd52%28Office.15%29.aspx) simple type.

You can optionally include a regular expression in the  **RegularExpression** attribute so that your add-in is only shown when an entity that matches the regular expression in present. To obtain matches to regular expressions specified in **ItemHasKnownEntity** rules, you can use the **getRegExMatches** or **getFilteredEntitiesByName** method for the currently selected Outlook item.

The following example shows a collection of  **Rule** elements that show the add-in when one of the specified well-known entities is present in the message.


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="TaskSuggestion" />
</Rule>
```

The following example shows an  **ItemHasKnownEntity** rule with a **RegularExpression** attribute that activates the add-in when a URL that contains the word "contoso" is present in a message.




```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegularExpression="contoso" />
```

For more information about entities in activation rules, see [Match strings in an Outlook item as well-known entities](../../outlook/match-strings-in-an-item-as-well-known-entities.md).


## ItemHasRegularExpressionMatch rule


The  **ItemHasRegularExpressionMatch** complex type defines a rule that uses a regular expression to match the contents of the specified property of an item. If text that matches the regular expression is found in the specified property of the item, Outlook activates the add-in bar and displays the add-in. You can use the **getRegExMatches** or **getRegExMatchesByName** method of the object that represents the currently selected item to obtain matches for the specified regular expression.

The following example shows an  **ItemHasRegularExpressionMatch** that activates the add-in when the body of the selected item contains "apple", "banana", or "coconut", ignoring case.




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" pPropertyName="BodyAsPlaintext" IgnoreCase="true" />
```

For more information about using the  **ItemHasRegularExpressionMatch** rule, see [Use regular expression activation rules to show an Outlook add-in](../../outlook/use-regular-expressions-to-show-an-outlook-add-in.md).


## RuleCollection rule


The  **RuleCollection** complex type combines multiple rules into a single rule. You can specify whether the rules in the collection should be combined with a logical OR or a logical AND by using the **Mode** attribute.

When a logical AND is specified, an item must match all the specified rules in the collection to show the add-in. When a logical OR is specified, an item that matches any of the specified rules in the collection will show the add-in.


 >**Note**  The [Rule](http://msdn.microsoft.com/en-us/library/c6ce9d52-4b53-c6a6-de7e-c64106135c81%28Office.15%29.aspx) element that you use to specify a collection of rules is of the abstract [Rule](http://msdn.microsoft.com/en-us/library/bcd7a3a7-9cd4-a270-89e0-5386d1c6df01%28Office.15%29.aspx) complex type. The **RuleCollection** complex type extends this abstract **Rule** complex type. So when you specify a rule collection in a manifest, you must use the **xsi:type** attribute to specify the **RuleCollection** complex type.

You can combine  **RuleCollection** rules to form complex rules. The following example activates the add-in when the user is viewing an appointment or message item and the subject or body of the item contains an address.




```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read"/>
  </Rule>
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

The following example activates the add-in when the user is composing a message, or when the user is viewing an appointment and the subject or body of the appointment contains an address.




```XML
<Rule xsi:type="RuleCollection" Mode="Or"> 
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" /> 
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
  </Rule> 
</Rule>
```


## Limits for rules and regular expressions


To provide a satisfactory experience with Outlook add-ins, you should adhere to the activation and API usage guidelines. The following table shows general limits for regular expressions and rules but there are specific rules for different hosts. For more information, see [Limits for activation and JavaScript API for Outlook add-ins](../../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md) and [Troubleshoot Outlook add-in activation](../../outlook/troubleshoot-outlook-add-in-activation.md).

|**Add-in element**|**Guidelines**|
|:-----|:-----|
|manifest size|No larger than 256 KB.|
|rules|No more than 15 rules.|
|[ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx)|An Outlook rich client will apply the rule against the first 1 MB of the body, and not to the rest of the body.|
|regular expressions|For [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) or [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) rules for all Outlook hosts:<br><ul><li>Specify no more than 5 regular expressions in activation rules for an Outlook add-in. You cannot install an add-in if you exceed the that limit.</li><li>Specify regular expressions whose anticipated results are returned by the <b>getRegExMatches</b> method call within the first 50 matches. </li><li>Specify look-ahead assertions in regular expressions, but not look-behind, (?<=text), and negative look-behind (?<!text).</li><li>Specify regular expressions whose match does not exceed the limits in the table below.<br/><br/><table><tr><th>Limit on length of a regex match</th><th>Outlook rich clients</th><th>Outlook Web App for Devices</th></tr><tr><td>Item body is plain text</td><td>1.5 KB</td><td>3 KB</td></tr><tr><td>Item body it HTML</td><td>3 KB</td><td>3 KB</td></tr></table>|

## Additional resources

- [Outlook add-ins](../../outlook/outlook-add-ins.md)
- [Create Outlook add-ins for compose forms](../../outlook/compose-scenario.md)
- [Limits for activation and JavaScript API for Outlook add-ins](../../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Item Types and Message Classes](http://msdn.microsoft.com/library/15b709cc-7486-b6c7-88a3-4a4d8e0ab292%28Office.15%29.aspx)
- [Use regular expression activation rules to show an Outlook add-in](../../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
- [Match strings in an Outlook item as well-known entities](../../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
