
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
|[CitationText](https://dev.office.com/reference/add-ins/manifest/citationtext)|
|[Description](https://dev.office.com/reference/add-ins/manifest/description)|
|[DictionaryName](https://dev.office.com/reference/add-ins/manifest/dictionaryname)|
|[DictionaryHomePage](https://dev.office.com/reference/add-ins/manifest/dictionaryhomepage)|
|[DisplayName](https://dev.office.com/reference/add-ins/manifest/displayname)|
|[HighResolutionIconUrl](https://dev.office.com/reference/add-ins/manifest/highresolutioniconurl)|
|[IconUrl](https://dev.office.com/reference/add-ins/manifest/iconurl)|
|[QueryUri](https://dev.office.com/reference/add-ins/manifest/queryuri)|
|[SourceLocation](https://dev.office.com/reference/add-ins/manifest/sourcelocation)|
|[SupportUrl](https://dev.office.com/reference/add-ins/manifest/supporturl)|

## Attributes



|**Attribute**|**Type**|**Required**|**Description**|
|:-----|:-----|:-----|:-----|
|Locale|string|required|Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.|
|Value|string|required|Specifies value of the setting expressed for the specified locale.|

## Additional resources



- [Localization for Office Add-ins](../../docs/develop/localization.md#off15wecon_LocalesManifest)
    
