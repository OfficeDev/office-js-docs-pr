---
title: Use regular expression activation rules to show an add-in
description: Learn how to use regular expression activation rules for Outlook contextual add-ins.
ms.date: 10/03/2022
ms.localizationpriority: medium
---

# Use regular expression activation rules to show an Outlook add-in

You can specify regular expression rules to have a [contextual add-in](contextual-outlook-add-ins.md) activated when a match is found in specific fields of the message. Contextual add-ins activate only in read mode. Outlook doesn't activate contextual add-ins when the user is composing an item. There are also other scenarios where Outlook doesn't activate add-ins, for example, digitally signed items. For more information, see [Activation rules for Outlook add-ins](activation-rules.md).

[!include[Unified manifest for Microsoft 365 does not support contextual add-ins](../includes/json-manifest-outlook-contextual-not-supported.md)]

You can specify a regular expression as part of an [ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule) rule or [ItemHasKnownEntity](/javascript/api/manifest/rule#itemhasknownentity-rule) rule in the add-in XML manifest. The rules are specified in a [DetectedEntity](/javascript/api/manifest/extensionpoint#detectedentity) extension point.

Outlook evaluates regular expressions based on the rules for the JavaScript interpreter used by the browser on the client computer. Outlook supports the same list of special characters that all XML processors also support. The following table lists these special characters. You can use these characters in a regular expression by specifying the escape sequence of the corresponding character, as described in the following table.

|Character|Description|Escape sequence to use|
|:-----|:-----|:-----|
|`"`|Double quotation mark|`&quot;`|
|`&`|Ampersand|`&amp;`|
|`'`|Apostrophe|`&apos;`|
|`<`|Less-than sign|`&lt;`|
|`>`|Greater-than sign|`&gt;`|

## ItemHasRegularExpressionMatch rule

An  `ItemHasRegularExpressionMatch` rule is useful in controlling activation of an add-in based on specific values of a supported property. The `ItemHasRegularExpressionMatch` rule has the following attributes.

|Attribute name|Description|
|:-----|:-----|
|`RegExName`|Specifies the name of the regular expression so that you can refer to the expression in the code for your add-in.|
|`RegExValue`|Specifies the regular expression that will be evaluated to determine whether the add-in should be shown.|
|`PropertyName`|Specifies the name of the property that the regular expression will be evaluated against. The allowed values are `BodyAsHTML`, `BodyAsPlaintext`, `SenderSMTPAddress`, and `Subject`.<br/><br/>If you specify `BodyAsHTML`, Outlook only applies the regular expression if the item body is HTML. Otherwise, Outlook returns no matches for that regular expression.<br/><br/>If you specify `BodyAsPlaintext`, Outlook always applies the regular expression on the item body.<br/><br/>**Important:** If you need to specify the **Highlight** attribute for the **\<Rule\>** element, you must set the **PropertyName** attribute to `BodyAsPlaintext`. |
|`IgnoreCase`|Specifies whether to ignore case when matching the regular expression specified by `RegExName`.|
| `Highlight` | Specifies how the client should highlight matching text. This element can only be applied to `Rule` elements within `ExtensionPoint` elements. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.<br/><br/>**Important:** To specify the **Highlight** attribute in the **\<Rule\>** element, you must set the **PropertyName** attribute to `BodyAsPlaintext`. |

### Best practices for using regular expressions in rules

Pay special attention to the following when you use regular expressions.

- If you specify an `ItemHasRegularExpressionMatch` rule on the body of an item, the regular expression should further filter the body and shouldn't attempt to return the entire body of the item. Using a regular expression such as `.*` to attempt to obtain the entire body of an item doesn't always return the expected results.
- The plain text body returned on one browser can be different in subtle ways on another. If you use an `ItemHasRegularExpressionMatch` rule with `BodyAsPlaintext` as the `PropertyName` attribute, test your regular expression on all the browsers that your add-in supports.

    Because different browsers use different ways to obtain the text body of a selected item, you should make sure that your regular expression supports the subtle differences that can be returned as part of the body text. For example, some browsers such as Internet Explorer 9 uses the `innerText` property of the DOM, and others such as Firefox uses the `.textContent()` method to obtain the text body of an item. Also, different browsers may return line breaks differently: a line break is `\r\n` on Internet Explorer, and `\n` on Firefox and Chrome. For more information, se [W3C DOM Compatibility - HTML](https://quirksmode.org/dom/html/).

- The HTML body of an item is slightly different between an Outlook rich client, and Outlook on the web or Outlook on mobile devices. Define your regular expressions carefully.

- Depending on the Outlook client, type of device, or property that a regular expression is being applied on, there are other best practices and limits for each of the clients that you should be aware of when designing regular expressions as activation rules. See [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) for details.

### Examples

The following `ItemHasRegularExpressionMatch` rule activates the add-in whenever the sender's SMTP email address matches `@contoso`, regardless of uppercase or lowercase characters.

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]"
    PropertyName="SenderSMTPAddress"
/>
```

The following is another way to specify the same regular expression using the  `IgnoreCase` attribute.

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@contoso"
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

The following `ItemHasRegularExpressionMatch` rule activates the add-in whenever a stock symbol is included in the body of the current item.

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    PropertyName="BodyAsPlaintext"
    RegExName="TickerSymbols"
    RegExValue="\b(NYSE|NASDAQ|AMEX):\s*[A-Za-z]+\b"/>

```

## ItemHasKnownEntity rule

An `ItemHasKnownEntity` rule activates an add-in based on the existence of an entity in the subject or body of the selected item. The [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) type defines the supported entities. Applying a regular expression on an `ItemHasKnownEntity` rule provides the convenience where activation is based on a subset of values for an entity (for example, a specific set of URLs, or telephone numbers with a certain area code).

> [!NOTE]
> Outlook can only extract entity strings in English regardless of the default locale specified in the manifest. Only messages support the `MeetingSuggestion` entity type; appointments don't support this. You can't extract entities from items in the **Sent Items** folder, nor can you use an `ItemHasKnownEntity` rule to activate an add-in for items in the **Sent Items** folder.

The `ItemHasKnownEntity` rule supports the attributes in the following table. Note that while specifying a regular expression is optional in an `ItemHasKnownEntity` rule, if you choose to use a regular expression as an entity filter, you must specify both the `RegExFilter` and `FilterName` attributes.

|Attribute name|Description|
|:-----|:-----|
|`EntityType`|Specifies the type of entity that must be found for the rule to evaluate to `true`. Use multiple rules to specify multiple types of entities.|
|`RegExFilter`|Specifies a regular expression that further filters instances of the entity specified by `EntityType`.|
|`FilterName`|Specifies the name of the regular expression specified by `RegExFilter`, so that it is subsequently possible to refer to it by code.|
|`IgnoreCase`|Specifies whether to ignore case when matching the regular expression specified by `RegExFilter`.|

### Examples

The following `ItemHasKnownEntity` rule activates the add-in whenever there is a URL in the subject or body of the current item, and the URL contains the string `youtube`, regardless of the case of the string.

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="Url"
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

## Using regular expression results in code

You can obtain matches to a regular expression by using the following methods on the current item.

- [getRegExMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) returns matches in the current item for all regular expressions specified in `ItemHasRegularExpressionMatch` and `ItemHasKnownEntity` rules of the add-in.

- [getRegExMatchesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) returns matches in the current item for the identified regular expression specified in an `ItemHasRegularExpressionMatch` rule of the add-in.

- [getFilteredEntitiesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) returns entire instances of entities that contain matches for the identified regular expression specified in an `ItemHasKnownEntity` rule of the add-in.

When the regular expressions are evaluated, the matches are returned to your add-in in an array object. For `getRegExMatches`, that object has the identifier of the name of the regular expression.

> [!NOTE]
> Outlook doesn't return matches in any particular order in the array. Also, you shouldn't assume that matches are returned in the same order in this array even when you run the same add-in on each of these clients on the same item in the same mailbox.

### Examples

The following is an example of a rule collection that contains an  `ItemHasRegularExpressionMatch` rule with a regular expression named `videoURL`.

```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="videoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="BodyAsPlaintext"/>
</Rule>
```

The following example uses `getRegExMatches` of the current item to set a variable `videos` to the results of the preceding `ItemHasRegularExpressionMatch` rule.

```js
const videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

Multiple matches are stored as array elements in that object. The following code example shows how to iterate over the matches for a regular expression named  `reg1` to build a string to display as HTML.

```js
function initDialer()
{
    let myEntities;
    let myString;
    let myCell;
    myEntities = Office.context.mailbox.item.getRegExMatches();

    myString = "";
    myCell = document.getElementById('dialerholder');
    // Loop over the myEntities collection.
    for (let i in myEntities.reg1) {
        myString += "<p><a href='callto:tel:" + myEntities.reg1[i] + "'>" + myEntities.reg1[i] + "</a></p>";
    }

    myCell.innerHTML = myString;
}
```

The following is an example of an `ItemHasKnownEntity` rule that specifies the `MeetingSuggestion` entity and a regular expression named `CampSuggestion`. Outlook activates the add-in if it detects that the currently selected item contains a meeting suggestion, and the subject or body contains the term `WonderCamp`.

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

The following code example uses `getFilteredEntitiesByName` on the current item to set a variable `suggestions` to an array of detected meeting suggestions for the preceding `ItemHasKnownEntity` rule.

```js
const suggestions = Office.context.mailbox.item.getFilteredEntitiesByName("CampSuggestion");
```

## See also

- [Outlook add-in: Contoso Order Number](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) - A sample contextual add-in that activates based on a regular expression match.
- [Create Outlook add-ins for read forms](read-scenario.md)
- [Activation rules for Outlook add-ins](activation-rules.md)
- [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md)
- [Best practices for regular expressions in the .NET framework](/dotnet/standard/base-types/best-practices)
