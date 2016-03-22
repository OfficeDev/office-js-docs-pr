
# OfficeApp element
The root element in the manifest of an Office Add-in.

 **Add-in type:** Content, Task pane, Mail


## Syntax:


```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```


## Contained in:

 _none_


## Must contain:



|**Element**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Id](../../reference/manifest/id.md)|x|x|x|
|[Version](../../reference/manifest/version.md)|x|x|x|
|[ProviderName](../../reference/manifest/providername.md)|x|x|x|
|[DefaultLocale](../../reference/manifest/defaultlocale.md)|x|x|x|
|[DefaultSettings](../../reference/manifest/defaultsettings.md)|x|x|x|
|[DisplayName](../../reference/manifest/displayname.md)|x|x|x|
|[Description](../../reference/manifest/description.md)|x|x|x|
|[FormSettings](../../reference/manifest/formsettings.md)||x||
|[Permissions](../../reference/manifest/permissions.md)|x||x|
|[Rule](../../reference/manifest/rule.md)||x||

## Can contain:



|**Element**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[AlternateId](../../reference/manifest/alternateid.md)|x|x|x|
|[IconUrl](../../reference/manifest/iconurl.md)|x|x|x|
|[HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md)|x|x|x|
|[SupportUrl](../../reference/manifest/supporturl.md)|x|x|x|
|[AppDomains](../../reference/manifest/appdomains.md)|x|x|x|
|[Hosts](../../reference/manifest/hosts.md)|x|x|x|
|[Requirements](../../reference/manifest/requirements.md)|x|x|x|
|[AllowSnapshot](../../reference/manifest/allowsnapshot.md)|x|||
|[Permissions](../../reference/manifest/permissions.md)||x||
|[DisableEntityHighlighting](../../reference/manifest/disableentityhighlighting.md)||x||
|[Dictionary](http://msdn.microsoft.com/library/c2563502-f020-4d12-a55e-dad35d59b9ac%28Office.15%29.aspx)|||x|

## Attributes


|||
|:-----|:-----|
|xmlns|Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`|
|xmlns:xsi|Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`|
|xsi:type|Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`|
