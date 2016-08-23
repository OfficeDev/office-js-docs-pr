# FormFactor element

Specifies the settings for an add-in for a given form factor. For example, defining a `Host` with the type `MailHost` and `DesktopFormFactor` will apply to Outlook for Desktop but  _not_ Outlook Web App or Outlook.com. It contains all the add-in information for that form factor except for the  **Resources** node.

Each FormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](./functionfile.md) and [ExtensionPoint element](./extensionpoint.md). 

The following FormFactors are supported:

- `DesktopFormFactor` (Office for Windows or Mac clients)

## Child elements

| Element                               | Required | Description  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](./extensionpoint.md) | Yes      | Defines where an add-in exposes functionality. |
| [FunctionFile](./functionfile.md)     | Yes      | A URL to a file that contains JavaScript functions.|
| [GetStarted](./getstarted.md)         | No       | Defines callout that appears when installing the add-in in Word, Excel and PowerPoint hosts |

## FormFactor example

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```