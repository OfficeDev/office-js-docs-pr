---
title: Add mobile support to an Outlook add-in
description: Adding support for Outlook Mobile requires updating the add-in manifest and possibly changing your code for mobile scenarios.
ms.date: 12/10/2019
localization_priority: Normal
---

# Add support for add-in commands for Outlook Mobile

Using add-in commands in Outlook Mobile allows your users to access the same functionality (with some [limitations](#code-considerations)) that they already have in Outlook on the web, Windows, and Mac. Adding support for Outlook Mobile requires updating the add-in manifest and possibly changing your code for mobile scenarios.

## Updating the manifest

The first step to enabling add-in commands in Outlook Mobile is to define them in the add-in manifest. The [VersionOverrides](/office/dev/add-ins/reference/manifest/versionoverrides) v1.1 schema defines a new form factor for mobile, [MobileFormFactor](/office/dev/add-ins/reference/manifest/mobileformfactor).

This element contains all of the information for loading the add-in in mobile clients. This enables you to define completely different UI elements and JavaScript files for the mobile experience.

The following example shows a single task pane button in a `MobileFormFactor` element.

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
  ...
  <MobileFormFactor>
    <FunctionFile resid="residUILessFunctionFileUrl" />
    <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
      <Group id="mobileMsgRead">
        <Label resid="groupLabel" />
        <Control xsi:type="MobileButton" id="TaskPaneBtn">
          <Label resid="residTaskPaneButtonName" />
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
      </Group>
    </ExtensionPoint>
  </MobileFormFactor>
  ...
</VersionOverrides>
```

This is very similar to the elements that appear in a [DesktopFormFactor](/office/dev/add-ins/reference/manifest/desktopformfactor) element, with some notable differences.

- The [OfficeTab](/office/dev/add-ins/reference/manifest/officetab) element is not used.
- The [ExtensionPoint](/office/dev/add-ins/reference/manifest/extensionpoint) element must have only one child element. If the add-in only adds one button, the child element should be a [Control](/office/dev/add-ins/reference/manifest/control) element. If the add-in adds more than one button, the child element should be a [Group](/office/dev/add-ins/reference/manifest/group) element that contains multiple `Control` elements.
- There is no `Menu` type equivalent for the `Control` element.
- The [Supertip](/office/dev/add-ins/reference/manifest/supertip) element is not used.
- The required icon sizes are different. Mobile add-ins minimally must support 25x25, 32x32 and 48x48 pixel icons.

## Code considerations

Designing an add-in for mobile introduces some additional considerations.

### Use REST instead of Exchange Web Services

The [Office.context.mailbox.makeEwsRequestAsync](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox#methods) method is not supported in Outlook Mobile. Add-ins should prefer to get information from the Office.js API when possible. If add-ins require information not exposed by the Office.js API, then they should use the [Outlook REST APIs](/outlook/rest/) to access the user's mailbox.

Mailbox requirement set 1.5 introduced a new version of [Office.context.mailbox.getCallbackTokenAsync](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox#methods) that can request an access token compatible with the REST APIs, and a new [Office.context.mailbox.restUrl](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox#properties) property that can be used to find the REST API endpoint for the user.

### Pinch zoom

By default users can use the "pinch zoom" gesture to zoom in on task panes. If this does not make sense for your scenario, be sure to disable pinch zoom in your HTML.

### Close task panes

In Outlook Mobile, task panes take up the entire screen and by default require the user to close them to return to the message. Consider using the [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--) method to close the task pane when your scenario is complete.

### Compose mode and appointments

Currently add-ins in Outlook Mobile only support activation when reading messages. Add-ins are not activated when composing messages or when viewing or composing appointments.

### Unsupported APIs

APIs introduced in requirement set 1.6 or later are not supported by Outlook Mobile. The following APIs from earlier requirement sets are also not supported.

  - [Office.context.officeTheme](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context#officetheme-officetheme)
  - [Office.context.mailbox.ewsUrl](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox#properties)
  - [Office.context.mailbox.convertToEwsId](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox#methods)
  - [Office.context.mailbox.convertToRestId](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox#methods)
  - [Office.context.mailbox.displayAppointmentForm](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox#methods)
  - [Office.context.mailbox.displayMessageForm](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox#methods)
  - [Office.context.mailbox.displayNewAppointmentForm](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox#methods)
  - [Office.context.mailbox.makeEwsRequestAsync](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox#methods)
  - [Office.context.mailbox.item.dateTimeModified](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#properties)
  - [Office.context.mailbox.item.displayReplyAllForm](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#methods)
  - [Office.context.mailbox.item.displayReplyForm](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#methods)
  - [Office.context.mailbox.item.getEntities](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#methods)
  - [Office.context.mailbox.item.getEntitiesByType](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#methods)
  - [Office.context.mailbox.item.getFilteredEntitiesByName](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#methods)
  - [Office.context.mailbox.item.getRegexMatches](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#methods)
  - [Office.context.mailbox.item.getRegexMatchesByName](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/office.context.mailbox.item#methods)

## See also

[Requirement set support](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)