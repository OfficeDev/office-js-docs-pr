# Rule element

Specifies the activation rule(s) that should be evaluated for this contextual mail add-in.

**Add-in type:** Mail contextual add-in

## Contained in

- [OfficeApp](officeapp.md)
- [ExtensionPoint](extensionpoint.md)

## Attributes

| Attribute | Required | Description |
|:-----|:-----|:-----|
| **xsi:type** | Yes | The type of rule being defined. |

The type of rule can be one of the following.

- [ItemIs](#itemis-rule)
- [ItemHasAttachment](#itemhasattachment-rule)
- [ItemHasKnownEntity](#itemhasknownentity-rule)
- [ItemHasRegularExpressionMatch](#itemhasregularexpressionmatch-rule)
- [RuleCollection](#rulecollection)

## ItemIs rule

Defines a rule that evaluates to true if the selected item is of the specified type.

### Attributes

| Attribute | Required | Description |
|:-----|:-----|:-----|
| **ItemType** | Yes | Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations. |
| **FormType** | No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md)) | Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`. |
| **ItemClass** | No | Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](https://docs.microsoft.com/outlook/add-ins/activation-rules). |
| **IncludeSubClasses** | No | Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`. |

### Example

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## ItemHasAttachment rule

Defines a rule that evaluates to true if the item contains an attachment.

### Example

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## ItemHasKnownEntity rule

Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.

### Attributes

| Attribute | Required | Description |
|:-----|:-----|:-----|
| **EntityType** | Yes | Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`. |
| **RegExFilter** | No | Specifies a regular expression to run against this entity for activation. |
| **FilterName** | No | Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code. |
| **IgnoreCase** | No | Specifies to ignore case when running the regular expression specified by the  **RegExFilter** attribute. |
| **Highlight** | No | **Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`. |

### Example

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## ItemHasRegularExpressionMatch rule

Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.

### Attributes

| Attribute | Required | Description |
|:-----|:-----|:-----|
| **RegExName** | Yes | Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in. |
| **RegExValue** | Yes | Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown. |
| **PropertyName** | Yes | Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHtml`, or `SenderSTMPAddress`. |
| **IgnoreCase** | No | Specifies to ignore the case when executing the regular expression. |
| **Highlight** | No | **Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching text. Can be one of the following: `all` or `none`. If not specified, the default value is `all`. |

### Example

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHtml" IgnoreCase="true" />
```

## RuleCollection

Defines a collection of rules and the logical operator to use when evaluating them.

### Attributes

| Attribute | Required | Description |
|:-----|:-----|:-----|
| **Mode** | Yes | Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`. |

### Example

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## See also

- [Activation rules for Outlook add-ins](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [Match strings in an Outlook item as well-known entities](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [Use regular expression activation rules to show an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)