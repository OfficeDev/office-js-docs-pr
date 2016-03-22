
# Rule element
Specifies the activation rule(s) that should be evaluated for this mail add-in.

 **Add-in type:** Mail


## Syntax:

 **ItemIs Rule** - Defines a rule that evaluates to true if the selected item is of the specified type.


```XML
<Rule xsi:type="ItemIs" 
   ItemType= ["Appointment" | "Message"]
   FormType=["Read" | "Edit" | "ReadOrEdit"] 
   ItemClass = "string " 
   IncludeSubClasses=["true" | "false"] />
```

 **ItemHasAttachment Rule** - Defines a rule that evaluates to true if the item contains an attachment.




```XML
<Rule xsi:type="ItemHasAttachment"  />
```

 **ItemHasKnownEntity** - Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.




```XML
<Rule xsi:type="ItemHasKnownEntity" 
  EntityType=["MeetingSuggestion" | "TaskSuggestion" |"Address" | "Url" | "PhoneNumber" | "EmailAddress" | "Contact" ]
  RegExFilter="string "
  FilterName="string "
  IgnoreCase=["true | false"]/>
```

 **ItemHasRegularExpressionMatch Rule** - Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    RegExName="string " 
    RegExValue="string " 
    PropertyName=["Subject" | "BodyAsPlainText" | "BodyAsHtml" | "SenderSTMPAddress"]
    IgnoreCase=["true" | "false"]
/>
```

 **RuleCollection Rule** - Defines a collection of rules and the logical operator to use when evaluating them.




```XML
<Rule xsi:type="RuleCollection" Mode=["And" | "Or"]>
   ...
</Rule>
```


## Contained in:

 _[OfficeApp](../../reference/manifest/officeapp.md)_


## Attributes:

 **ItemIs Rule attributes**



|**Attribute**|**Type**|**Required**|**Description**|
|:-----|:-----|:-----|:-----|
|ItemType|ItemType (string)|required|Specifies the item type to match. Can be one of the following:

|**ItemType**|**Corresponding ItemClass**|
|:-----|:-----|
|Appointment|IPM.Appointment|
|Message|Includes email messages, meeting requests, responses, and cancellations. The following are the corresponding message classes:IPM.NoteIPM.Schedule.Meeting.RequestIPM.Schedule.Meeting.NegIPM.Schedule.Meeting.PosIPM.Schedule.Meeting.TentIPM.Schedule.Meeting.Canceled|
|
|FormType|ItemFormType (string)|required|Specifies whether the app should appear in read or edit form for the item. Can be one of the following:
|
|
|**FormType**|**Description**|
|:-----|:-----|
|Read|Specifies to activate the mail add-in only in read forms (of the specified  **ItemType**).|
|Edit|Specifies to activate the mail add-in only in compose forms (of the specified  **ItemType**).|
|ReadOrEdit|Specifies to activate the mail add-in in both read and compose forms (of the specified  **ItemType**).|
|
|ItemClass|string|optional|Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](http://msdn.microsoft.com/library/f464a152-2dff-4fb3-bf98-c1a3639c3e80%28Office.15%29.aspx).|
|IncludeSubClasses|boolean|optional|Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is false.|
 **ItemHasAttachment Rule attributes**

None.

 **ItemHasKnownEntity Rule attributes**



|**Attribute**|**Type**|**Required**|**Description**|
|:-----|:-----|:-----|:-----|
|EntityType|KnownEntityType (string)|required|Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following:

|**KnownEntityType**|**Descripition**|
|:-----|:-----|
|MeetingSuggestion|Text that is identified by pattern recognition to reference an event or a meeting.|
|TaskSuggestion| Text that is identified by pattern recognition to contain an actionable phrase.|
|Address|Text that is identified by pattern recognition to reference a postal address in the United States.|
|Url|Text that is identified by pattern recognition to contain a file name or web address URL.|
|PhoneNumber| A series of digits that is identified by pattern recognition as a telephone number in North America.|
|EmailAddress|Text that is identified by pattern recognition to contain an SMTP format email address.|
|Contact|Text that is identified by pattern recognition to contain contact information.|
|
|RegExFilter|string|optional|Specifies a regular expression to run against this entity for activation.|
|FilterName|string|optional|Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.|
|IgnoreCase|boolean|optional|Specifies to ignore case when running the regular expression specified by the  **RegExFilter** attribute.|
 **ItemHasRegularExpressionMatch Rule attributes**



|**Attribute**|**Type**|**Required**|**Description**|
|:-----|:-----|:-----|:-----|
|RegExName|string|required|Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.|
|RegExValue|string|required|Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown. |
|PropertyName|PropertyName (string)|required|Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following:

|**PropertyName**|**Description**|
|:-----|:-----|
|Subject|Evaluates the regular expression against the item subject.|
|BodyAsPlainText|Evaluates the regular expression against the item body in plain text.|
|BodyAsHtml|Evaluates the regular expression against the item body if the body is available in HTML.|
|SenderSTMPAddress|Evaluates the regular expression against the SMTP address of the item sender.|
|
|IgnoreCase|boolean|optional|Specifies to ignore the case when executing the regular expression.|
 **RuleCollection Rule attributes**



|**Attribute**|**Type**|**Required**|**Description**|
|:-----|:-----|:-----|:-----|
|Mode|string|required|Specifies the logical operator to use when evaluating this rule collection. Can be either: "And" or "Or".|

## Additional resources



- [Activate a mail add-in in Outlook for a specific message class](http://msdn.microsoft.com/library/f464a152-2dff-4fb3-bf98-c1a3639c3e80%28Office.15%29.aspx) and [Activation rules for Outlook add-ins](../../docs/outlook/manifests/activation-rules.md#MailAppDefineRules_ItemIs)
    
- [Match strings in an Outlook item as well-known entities](../../docs/outlook/match-strings-in-an-item-as-well-known-entities.md)
    
- [Use regular expression activation rules to show an Outlook add-in](../../docs/outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
