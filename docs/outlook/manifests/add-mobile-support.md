# Add support for add-in commands for Outlook Mobile

> **Note**: Add-in commands for Outlook Mobile are currently only supported by Outlook for iOS.

Using add-in commands in Outlook Mobile allows your users to access the same functionality (with some [limitations](#code-considerations)) that they already have in Outlook for Windows, Outlook for Mac, and Outlook on the web. Adding support for Outlook Mobile requires updating the add-in manifest and possibly changing your code for mobile scenarios.

## Updating the manifest

The first step to enabling add-in commands in Outlook Mobile is to define them in the add-in manifest. The **VersionOverrides** v1.1 schema defines a new form factor for mobile, [MobileFormFactor](../../reference/manifest/mobileformfactor.md).

This element contains all of the information for loading the add-in in mobile clients. This enables you to define completely different UI elements and JavaScript files for the mobile experience.

The following example shows a single taskpane button in a **MobileFormFactor** element.

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
  ...
  <MobileFormFactor>
    <FunctionFile resid="residUILessFunctionFileUrl" />
    <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
      <Control xsi:type="MobileButton" id="TaskPane1Btn">
        <Label resid="residTaskPaneButton0Name" />
        <Icon xsi:type="bt:MobileIconList">
          <bt:Image size="25" scale="1" resid="tp0icon" />
          <bt:Image size="25" scale="2" resid="tp0icon" />
          <bt:Image size="25" scale="3" resid="tp0icon" />

          <bt:Image size="32" scale="1" resid="tp0icon" />
          <bt:Image size="32" scale="2" resid="tp0icon" />
          <bt:Image size="32" scale="3" resid="tp0icon" />

          <bt:Image size="48" scale="1" resid="tp0icon" />
          <bt:Image size="48" scale="2" resid="tp0icon" />
          <bt:Image size="48" scale="3" resid="tp0icon" />
        </Icon>
        <Action xsi:type="ShowTaskpane">
          <SourceLocation resid="residTaskpaneUrl" />
        </Action>
      </Control>
    </ExtensionPoint>
  </MobileFormFactor>
  ...
</VersionOverrides>
```

This is very similar to the elements that appear in a [DesktopFormFactor](../../reference/manifest/desktopformfactor.md) element, with some notable differences.

- The [OfficeTab](../../reference/manifest/officetab.md) element is not used.
- The [ExtensionPoint](../../reference/manifest/exensionpoint.md) element must have only one child element. If the add-in only adds one button, the child element should be a [Control](../../reference/manifest/control.md) element. If the add-in adds more than one button, the child element should be a [Group](../../reference/manifest/group.md) element that contains multiple `Control` elements.
- There is no `Menu` type equivalent for the `Control` element.
- The [Supertip](../../reference/manifest/supertip.md) element is not used.
- The required icon sizes are different. Mobile add-ins minimally must support 25x25, 32x32 and 48x48 pixel icons.

## Code considerations

Designing an add-in for mobile introduces some additional considerations.

### Use REST instead of Exchange Web Services

The [Office.context.mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) method is not supported in Outlook Mobile. Add-ins should prefer to get information from the Office.js API when possible. If add-ins require information not exposed by the Office.js API, then they should use the [Outlook REST APIs](https://dev.outlook.com/restapi/reference) to access the user's mailbox. 

Mailbox requirement set 1.5 introduces a new version of [Office.context.mailbox.getCallbackTokenAsync](https://dev.outlook.com/reference/add-ins/1.5/Office.context.mailbox.html#getCallbackTokenAsync) that can request an access token compatible with the REST APIs, and a new [Office.context.mailbox.restUrl](https://dev.outlook.com/reference/add-ins/1.5/Office.context.mailbox.html#restUrl) property that can be used to find the REST API endpoint for the user.

### Pinch zoom

By default users can use the "pinch zoom" gesture to zoom in on taskpanes. If this does not make sense for your scenario, be sure to disable pinch zoom in your HTML.

### Closing taskpanes

In Outlook Mobile, taskpanes take up the entire screen and by default require the user to close them to return to the message. Consider using the [Office.context.ui.closeContainer](https://dev.outlook.com/reference/add-ins/1.5/Office.context.ui.html#closeContainer) method to close the taskpane when your scenario is complete.

### Compose mode and appointments

Currently add-ins in Outlook Mobile only support activation when reading messages. Add-ins are not activated when composing messages or when viewing or composing appointments.

### Unsupported APIs

The following APIs are not supported by Outlook Mobile.

  - [Office.context.officeTheme](../../reference/outlook/Office.context.md)
  - [Office.context.mailbox.ewsUrl](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.convertToEwsId](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.convertToRestId](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.displayAppointmentForm](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.displayMessageForm](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.displayNewAppointmentForm](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md)
  - [Office.context.mailbox.item.dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.resources](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.displayReplyAllForm](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.displayReplyForm](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getEntities](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getEntitiesByType](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getFilteredEntitiesByName](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getRegexMatches](../../reference/outlook/Office.context.mailbox.item.md)
  - [Office.context.mailbox.item.getRegexMatchesByName](../../reference/outlook/Office.context.mailbox.item.md)