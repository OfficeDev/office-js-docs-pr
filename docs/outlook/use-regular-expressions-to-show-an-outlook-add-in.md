
# Use regular expression activation rules to show an Outlook add-in

You can specify regular expression rules to have an Outlook add-in activated in read scenarios - when the user views a message or appointment in the Reading Pane or inspector, Outlook evaluates regular expression rules to determine if it should activate your add-in. Outlook does not evaluate these rules when the user is composing an item. There are also other scenarios where Outlook does not activate add-ins, for example, items protected by Information Rights Management (IRM) or in the Junk Email folder. For more information, see [Activation rules for Outlook add-ins](../outlook/manifests/activation-rules.md#MailAppDefineRules_Activation).

You can specify a regular expression as part of an [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) rule or [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) rule in the add-in XML manifest. Outlook evaluates regular expressions based on the rules for the JavaScript interpreter used by the browser on the client computer. Outlook supports the same list of special characters that all XML processors also support. The following table lists these special characters. You can use these characters in a regular expression by specifying the escaped sequence for the corresponding character, as described in the following table.



|**Character**|**Description**|**Escaped sequence to use**|
|:-----|:-----|:-----|
|"|Double quotation mark|&amp;quot;|
|&amp;|Ampersand|&amp;amp;|
|'|Apostrophe|&amp;apos;|
|<|Less-than sign|&amp;lt;|
|>|Greater-than sign|&amp;gt;|

## ItemHasRegularExpressionMatch rule


An  **ItemHasRegularExpressionMatch** rule is useful in controlling activation of an add-in based on specific values of a supported property. The **ItemHasRegularExpressionMatch** rule has the following attributes.



|**Attribute name**|**Description**|
|:-----|:-----|
|**RegExName**|Specifies the name of the regular expression so that you can refer to the expression in the code for your add-in.|
|**RegExValue**|Specifies the regular expression that will be evaluated to determine whether the add-in should be shown.|
|**PropertyName**|Specifies the name of the property that the regular expression will be evaluated against. The allowed values are  **BodyAsHTML**,  **BodyAsPlaintext**,  **SenderSMTPAddress**, and  **Subject**. If you specify  **BodyAsHTML**, Outlook applies the regular expression only if the item body is HTML, and otherwise Outlook returns no matches for that regular expression. Because appointments are always saved in Rich Text Format, a regular expression that specifies  **BodyAsHTML** does not match any strings in the body of appointment items.If you specify  **BodyAsPlaintext**, Outlook always applies the regular expression on the item body.|
|**IgnoreCase**|Specifies whether to ignore case when matching the regular expression specified by  **RegExName**.|

### Best practices for using regular expressions in rules

Pay special attention to the following when you use regular expressions:


- If you specify an  **ItemHasRegularExpressionMatch** rule on the body of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to attempt to obtain the entire body of an item does not always return the expected results.
    
- The plain text body returned on one browser can be different in subtle ways on another. If you use an [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) rule with **BodyAsPlaintext** as the **PropertyName** attribute, test your regular expression on all the browsers that your add-in supports.
    
    Because different browsers use different ways to obtain the text body of a selected item, you should make sure that your regular expression supports the subtle differences that can be returned as part of the body text. For example, some browsers such as Internet Explorer 9 uses the  **innerText** property of the DOM, and others such as Firefox uses the **.textContent()** method to obtain the text body of an item. Also, different browsers may return line breaks differently: a line break is "\r\n" on Internet Explorer, and "\n" on Firefox and Chrome. For more information, se [W3C DOM Compatibility - HTML](http://www.quirksmode.org/dom/w3c_html.mdl#t07).
    
- The HTML body of an item is slightly different between an Outlook rich client, and Outlook Web App or OWA for Devices. Define your regular expressions carefully. As an example, consider the following regular expression used in an  **ItemHasRegularExpressionMatch** rule with **BodyAsHTML** as the **PropertyName** attribute value:
    
    ```
      http.*\.contoso\.com
    ```


    A rule with this regular expression would match the string "http-equiv="Content-Type" which exists in the HTML body of an item in an Outlook rich client, as part of the following  **META** tag:
    

    ```HTML
      <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=us-ascii">
    ```


The same rule does not return this match in Outlook Web App and OWA for Devices because the HTML body in these hosts does not include that  **META** tag. This can affect whether the add-in is activated appropriately for the various Outlook clients. In this example, use the following regular expression instead:
    

    ```
      http://.*\.contoso\.com/
    ```

- Depending on the host application, type of device, or property that a regular expression is being applied on, there are other best practices and limits for each of the hosts that you should be aware of when designing regular expressions as activation rules. See [Limits for activation and JavaScript API for Outlook add-ins](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md) for details.
    

### Examples

The following  **ItemHasRegularExpressionMatch** rule activates the add-in whenever the sender's SMTP email address matches "@contoso", regardless of uppercase or lowercase characters.


```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    RegExName="addressMatches" 
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]" 
    PropertyName="SenderSMTPAddress"
/>
```

The following is another way to specify the same regular expression using the  **IgnoreCase** attribute.




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    RegExName="addressMatches" 
    RegExValue="@contoso" 
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

The following  **ItemHasRegularExpressionMatch** rule activates the add-in whenever a stock symbol is included in the body of the current item.




```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" 
    PropertyName="BodyAsPlaintext" 
    RegExName="TickerSymbols" 
    RegExValue="\b(NYSE|NASDAQ|AMEX):\s*[A-Za-z]+\b"/>

```


## ItemHasKnownEntity rule


An  **ItemHasKnownEntity** rule activates a add-in based on the existence of an entity in the subject or body of the selected item. The [KnownEntityType](http://msdn.microsoft.com/en-us/library/432d413b-9fcc-eb50-cfea-0ed10a43bd52%28Office.15%29.aspx) type defines the supported entities. Applying a regular expression on an **ItemHasKnownEntity** rule provides the convenience where activation is based on a subset of values for an entity (for example, a specific set of URLs, or telephone numbers with a certain area code).


 >**Note**  Outlook can extract entity strings in only English regardless of the default locale specified in the manifest. Only messages but not appointments support the  **MeetingSuggestion** entity type.You cannot extract entities from items in the Sent Items folder, nor can you use an [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) rule to activate an add-in for items in the Sent Items folder.

The  **ItemHasKnownEntity** rule supports the attributes in the following table. Note that while specifying a regular expression is optional in an **ItemHasKnownEntity** rule, if you choose to use a regular expression as an entity filter, you must specify both the **RegExFilter** and **FilterName** attributes.



|**Attribute name**|**Description**|
|:-----|:-----|
|**EntityType**|Specifies the type of entity that must be found for the rule to evaluate to  **true**. Use multiple rules to specify multiple types of entities.|
|**RegExFilter**|Specifies a regular expression that further filters instances of the entity specified by  **EntityType**.|
|**FilterName**|Specifies the name of the regular expression specified by  **RegExFilter**, so that it is subsequently possible to refer to it by code.|
|**IgnoreCase**|Specifies whether to ignore case when matching the regular expression specified by  **RegExFilter**.|

### Examples

The following  **ItemHasKnownEntity** rule activates the add-in whenever there is a URL in the subject or body of the current item, and the URL contains the string "youtube", regardless of the case of the string.


```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="Url" 
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```


## Using regular expression results in code


You can obtain matches to a regular expression by using the following methods on the current item:


- [getRegExMatches](../../reference/outlook/Office.context.mailbox.item.md) returns matches in the current item for all regular expressions specified in **ItemHasRegularExpressionMatch** and **ItemHasKnownEntity** rules of the add-in.
    
- [getRegExMatchesByName](../../reference/outlook/Office.context.mailbox.item.md) returns matches in the current item for the identified regular expression specified in an **ItemHasRegularExpressionMatch** rule of the add-in.
    
- [getFilteredEntitiesByName](../../reference/outlook/Office.context.mailbox.item.md) returns entire instances of entities that contain matches for the identified regular expression specified in an **ItemHasKnownEntity** rule of the add-in.
    
When the regular expressions are evaluated, the matches are returned to your add-in in an array object. For  **getRegExMatches**, that object has the identifier of the name of the regular expression. 


 >**Note**  An Outlook rich client does not return matches in any particular order in the array. Also, you should not assume the Outlook rich client to return matches in the same order in this array as Outlook Web App or OWA for Devices, even when you run the same add-in on each of these clients on the same item in the same mailbox. For other differences in processing regular expressions between an Outlook rich client and Outlook Web App or OWA for Devices, see [Limits for activation and JavaScript API for Outlook add-ins](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md).


### Examples

The following is an example of a rule collection that contains an  **ItemHasRegularExpressionMatch** rule with a regular expression named `videoURL`.


```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="VideoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="Body"/>
</Rule>
```

The following example uses  **getRegExMatches** of the current item to set a variable `videos` to the results of the preceding **ItemHasRegularExpressionMatch** rule.




```
var videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

Multiple matches are stored as array elements in that object. The following code example shows how to iterate over the matches for a regular expression named  `reg1` to build a string to display as HTML.




```js
function initDialer() 
{
    var myEntities;
    var myString;
    var myCell;
    myEntities = _Item.getRegExMatches();

    myString = "";
    myCell = document.getElementById('dialerholder');
    // Loop over the myEntities collection.
    for (var i in myEntities.reg1) {
        myString += "<p><a href='callto:tel:" + myEntities.reg1[i] + "'>" + myEntities.reg1[i] + "</a></p>";
    }
    myCell.innerHTML = myString;
}

```

The following is an example of an  **ItemHasKnownEntity** rule that specifies the **MeetingSuggestion** entity and a regular expression named `CampSuggestion`. Outlook activates the add-in if it detects that the currently selected item contains a meeting suggestion, and the subject or body contains the term "WonderCamp".




```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

The following code example uses  **getFilteredEntitiesByName(name)** of the current item to set a variable `suggestions` to get an array of detected meeting suggestions for the preceding **ItemHasKnownEntity** rule.




```
var suggestions = Office.context.mailbox.item.getFilteredEntitiesByName(CampSuggestion);
```


## Additional resources



- [Create Outlook add-ins for read forms](../outlook/read-scenario.md)
    
- [Activation rules for Outlook add-ins](../outlook/manifests/activation-rules.md)
    
- [Limits for activation and JavaScript API for Outlook add-ins](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    
- [Match strings in an Outlook item as well-known entities](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
- [Best Practices for Regular Expressions in the .NET Framework](http://msdn.microsoft.com/en-us/library/618e5afb-3a97-440d-831a-70e4c526a51c%28Office.15%29.aspx)
    
