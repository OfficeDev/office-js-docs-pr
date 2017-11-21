
# DefaultSettings element
Specifies the default source location and other default settings for your content or task pane add-in .

 **Add-in type:** Content, Task pane


## Syntax:


```XML
<DefaultSettings>
  ...
</DefaultSettings>
```


## Contained in:

[OfficeApp](https://dev.office.com/reference/add-ins/manifest/officeapp)


## Can contain:



|**Element**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](https://dev.office.com/reference/add-ins/manifest/sourcelocation)|x||x|
|[RequestedWidth](https://dev.office.com/reference/add-ins/manifest/requestedwidth)|x|||
|[RequestedHeight](https://dev.office.com/reference/add-ins/manifest/requestedheight)|x|||

## Remarks

The source location and other settings in the  **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the [FormSettings](https://dev.office.com/reference/add-ins/manifest/formsettings) element.

