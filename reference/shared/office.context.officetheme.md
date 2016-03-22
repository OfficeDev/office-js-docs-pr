
# Context.officeTheme property
Provides access to the properties for Office theme colors.

 **Important:** This API currently works only in Excel, Outlook, PowerPoint, and Word in [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) on Windows desktop.


|||
|:-----|:-----|
|**Hosts:**|Excel, Outlook, PowerPoint, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Not in a set|
|**Added in**|1.3|



```js
Office.context.officeTheme
```


## Members


**Properties**

|||
|:-----|:-----|
|Name|Description|
|[bodyBackgroundColor ](../../reference/shared/office.context.bodybackgroundcolor.md)|Gets the Office theme body background color.|
|[bodyForegroundColor](../../reference/shared/office.context.bodyforegroundcolor.md)|Gets the Office theme body foreground color.|
|[controlBackgroundColor](../../reference/shared/office.context.controlbackgroundcolor.md)|Gets the Office theme control background color.|
|[controlForegroundColor](../../reference/shared/office.context.controlforegroundcolor.md)|Gets the Office theme control foreground color.|

## Remarks

Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with  **File** > **Office Account** > **Office Theme** UI, which is applied across all Office host applications. Using Office theme colors is appropriate for Outlook and task pane add-ins.


## Example


```js
function applyOfficeTheme(){
    // Get office theme colors.
    var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
    var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
    var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
    var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

    // Apply body background color to a CSS class.
    $('.body').css('background-color', bodyBackgroundColor);
}
```


## Support details



|||
|:-----|:-----|
|**Minimum permission level**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Add-in types**|Content, task pane, Outlook|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history


|**Version**|**Changes**|
|:-----|:-----|
|1.3|Introduced|
