
# Override element
Provides a way to specify the value of a setting for an additional locale.

 **Add-in type:** Content, Task pane, Mail


## Syntax:


```XML
<Override Locale="string " Value="string " />
```


## Contained in:


||
|:-----|
|[CitationText](../../reference/manifest/citationtext.md)|
|[Description](../../reference/manifest/description.md)|
|[DictionaryName](../../reference/manifest/dictionaryname.md)|
|[DictionaryHomePage](../../reference/manifest/dictionaryhomepage.md)|
|[DisplayName](../../reference/manifest/displayname.md)|
|[HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md)|
|[IconUrl](../../reference/manifest/iconurl.md)|
|[QueryUri](../../reference/manifest/queryuri.md)|
|[SourceLocation](../../reference/manifest/sourcelocation.md)|
|[SupportUrl](../../reference/manifest/supporturl.md)|

## Attributes



|**Attribute**|**Type**|**Required**|**Description**|
|:-----|:-----|:-----|:-----|
|Locale|string|required|Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.|
|Value|string|required|Specifies value of the setting expressed for the specified locale.|

## Additional resources



- [Localization for Office Add-ins](../../docs/develop/localization.md#off15wecon_LocalesManifest)
    
