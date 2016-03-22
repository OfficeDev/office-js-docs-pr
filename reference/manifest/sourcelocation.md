
# SourceLocation element
Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long.

 **Add-in type:** Content, Task pane, Mail


## Syntax:


```XML
<SourceLocation DefaultValue="string " />
```


## Contained in:

[DefaultSettings](../../reference/manifest/defaultsettings.md) (Content and task pane add-ins)

[FormSettings](../../reference/manifest/formsettings.md) (Mail add-ins)


## Can contain:

[Override](../../reference/manifest/override.md)


## Attributes



|**Attribute**|**Type**|**Required**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|required|Specifies the default value for this setting for the locale specified in the [DefaultLocale](../../reference/manifest/defaultlocale.md) element.|
