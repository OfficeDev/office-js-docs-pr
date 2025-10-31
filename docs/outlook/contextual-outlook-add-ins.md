---
title: Contextual Outlook add-ins
description: Initiate tasks related to a message without leaving the message itself to result in an easier and richer user experience.
ms.date: 10/30/2025
ms.localizationpriority: medium
---

# Contextual Outlook add-ins

Contextual add-ins are Outlook add-ins that activate based on text in a message or appointment. By using contextual add-ins, a user can initiate tasks related to a mail item without leaving the item itself. For example, a contextual add-in can choose a string in the body of a mail item that opens a meeting suggestion add-in.

You can specify regular expression rules to activate a contextual add-in when a match is found in specific fields of the message. Contextual add-ins only activate in read mode. Outlook doesn't activate contextual add-ins when the user is composing an item.

> [!IMPORTANT]
> Entity-based contextual Outlook add-ins are now retired. As an alternative solution, implement regular expression rules in your contextual add-in.

## Configure the manifest

[!INCLUDE [Unified manifest for Microsoft 365 does not support contextual add-ins](../includes/json-manifest-outlook-contextual-not-supported.md)]

A contextual add-in's manifest must include an [ExtensionPoint](/javascript/api/manifest/extensionpoint#detectedentity) element with its `xsi:type` attribute set to `DetectedEntity`. Within the `<ExtensionPoint>` element, the add-in must then specify a regular expression rule using the [Rule](/javascript/api/manifest/rule) element with its `xsi:type` attribute set to [ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule).

The following example activates an add-in whenever a stock symbol is included in the body of the current mail item.

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="Context.Label" />
  <SourceLocation resid="DetectedEntity.URL" />
  <Rule xsi:type="ItemHasRegularExpressionMatch" PropertyName="BodyAsPlaintext" RegExName="TickerSymbols" RegExValue="\b(NYSE|NASDAQ|AMEX):\s*[A-Za-z]+\b" />
</ExtensionPoint>
```

### Supported characters in regular expression rules

Outlook evaluates regular expressions based on the rules for the JavaScript interpreter used by the browser or webview control on the client computer. For brevity, this article uses "browser" to refer to "browser or webview control". Outlook supports the same list of special characters that all XML processors also support. The following table lists these special characters. You can use these characters in a regular expression by specifying the escape sequence of the corresponding character.

|Character|Description|Escape sequence to use|
|:-----|:-----|:-----|
|`"`|Double quotation mark|`&quot;`|
|`&`|Ampersand|`&amp;`|
|`'`|Apostrophe|`&apos;`|
|`<`|Less-than sign|`&lt;`|
|`>`|Greater-than sign|`&gt;`|

### Best practices for using regular expressions in rules

Be mindful of the following when you use regular expressions.

- If you specify an `ItemHasRegularExpressionMatch` rule on the body of an item, the regular expression should further filter the body and shouldn't attempt to return the entire body of the item. Using a regular expression such as `.*` to attempt to obtain the entire body of an item doesn't always return the expected results.
- The plain text body returned on one browser can be different in subtle ways on another. If you use an `ItemHasRegularExpressionMatch` rule with `BodyAsPlaintext` as the `PropertyName` attribute, test your regular expression on all the browsers that your add-in supports.

    Because different browsers use different ways to obtain the text body of a selected item, you should make sure that your regular expression supports the subtle differences that can be returned as part of the body text. For example, browsers may return line breaks differently. For more information, see [W3C DOM Compatibility - HTML](https://quirksmode.org/dom/html/).

- The HTML body of an item is slightly different between classic Outlook on Windows or Outlook on Mac, and Outlook on the web, on mobile devices, or [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627). Define your regular expressions carefully.

- Depending on the Outlook client, type of device, or property that a regular expression is being applied on, there are other best practices and limits for each of the clients that you should be aware of when designing regular expressions as activation rules. For details, see [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md).

## Use regular expression results in your JavaScript code

In the JavaScript code of your add-in, you can obtain matches to a regular expression by using the following methods on the current item.

- [getRegExMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) returns matches in the current item for all regular expressions specified in an `ItemHasRegularExpressionMatch` rule of the add-in.

- [getRegExMatchesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) returns matches in the current item for the identified regular expression specified in an `ItemHasRegularExpressionMatch` rule of the add-in.

- [getSelectedRegExMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) returns highlighted matches in the current item for the regular expression specified in an `ItemHasRegularExpressionMatch` rule of the add-in.

When the regular expressions are evaluated, the matches are returned to your add-in in an array object. For `getRegExMatches`, that object has the identifier of the name of the regular expression.

> [!NOTE]
> Outlook doesn't return matches in any particular order in the array. Also, you shouldn't assume that matches are returned in the same order in this array even when you run the same add-in on each of these clients on the same item in the same mailbox.

The following is an example of a rule collection that contains an `ItemHasRegularExpressionMatch` rule with a regular expression named `videoURL`.

```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="videoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="BodyAsPlaintext"/>
</Rule>
```

The `getRegExMatches` method is then called on the current message to set a variable `videos` to the results of specified `ItemHasRegularExpressionMatch` rule.

```js
const videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

## See also

- [Outlook add-in: Contoso Order Number](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) (sample contextual add-in that activates based on a regular expression match)
- [Build your first Outlook add-in](../quickstarts/outlook-quickstart-yo.md)
