
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
|[Id](https://dev.office.com/reference/add-ins/manifest/id)|x|x|x|
|[Version](https://dev.office.com/reference/add-ins/manifest/version)|x|x|x|
|[ProviderName](https://dev.office.com/reference/add-ins/manifest/providername)|x|x|x|
|[DefaultLocale](https://dev.office.com/reference/add-ins/manifest/defaultlocale)|x|x|x|
|[DefaultSettings](https://dev.office.com/reference/add-ins/manifest/defaultsettings)|x|x|x|
|[DisplayName](https://dev.office.com/reference/add-ins/manifest/displayname)|x|x|x|
|[Description](https://dev.office.com/reference/add-ins/manifest/description)|x|x|x|
|[FormSettings](https://dev.office.com/reference/add-ins/manifest/formsettings)||x||
|[Permissions](https://dev.office.com/reference/add-ins/manifest/permissions)|x||x|
|[Rule](https://dev.office.com/reference/add-ins/manifest/rule)||x||

## Can contain:



|**Element**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[AlternateId](https://dev.office.com/reference/add-ins/manifest/alternateid)|x|x|x|
|[IconUrl](https://dev.office.com/reference/add-ins/manifest/iconurl)|x|x|x|
|[HighResolutionIconUrl](https://dev.office.com/reference/add-ins/manifest/highresolutioniconurl)|x|x|x|
|[SupportUrl](https://dev.office.com/reference/add-ins/manifest/supporturl)|x|x|x|
|[AppDomains](https://dev.office.com/reference/add-ins/manifest/appdomains)|x|x|x|
|[Hosts](https://dev.office.com/reference/add-ins/manifest/hosts)|x|x|x|
|[Requirements](https://dev.office.com/reference/add-ins/manifest/requirements)|x|x|x|
|[AllowSnapshot](https://dev.office.com/reference/add-ins/manifest/allowsnapshot)|x|||
|[Permissions](https://dev.office.com/reference/add-ins/manifest/permissions)||x||
|[DisableEntityHighlighting](https://dev.office.com/reference/add-ins/manifest/disableentityhighlighting)||x||
|[Dictionary](https://dev.office.com/reference/add-ins/manifest/dictionary)|||x|
|[VersionOverrides](https://dev.office.com/reference/add-ins/manifest/versionoverrides)|X|X|X|

## Attributes


|||
|:-----|:-----|
|xmlns|Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`|
|xmlns:xsi|Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`|
|xsi:type|Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`|
