---
title: Add mobile support to an Outlook add-in
description: Learn how to add support for Outlook Mobile including how to update the add-in manifest and change your code for mobile scenarios, if necessary.
ms.date: 10/17/2022
ms.localizationpriority: medium
---

# Add support for add-in commands for Outlook Mobile

Using add-in commands in Outlook Mobile allows your users to access the same functionality (with some [limitations](#code-considerations)) that they already have in Outlook on the web, Windows, and Mac. Adding support for Outlook Mobile requires updating the add-in manifest and possibly changing your code for mobile scenarios.

## Updating the manifest

[!INCLUDE [Teams manifest not supported on mobile devices](../includes/no-mobile-with-json-note.md)]

The first step to enabling add-in commands in Outlook Mobile is to define them in the add-in manifest. The [VersionOverrides](/javascript/api/manifest/versionoverrides) v1.1 schema defines a new form factor for mobile, [MobileFormFactor](/javascript/api/manifest/mobileformfactor).

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

This is very similar to the elements that appear in a [DesktopFormFactor](/javascript/api/manifest/desktopformfactor) element, with some notable differences.

- The [OfficeTab](/javascript/api/manifest/officetab) element is not used.
- The [ExtensionPoint](/javascript/api/manifest/extensionpoint) element must have only one child element. If the add-in only adds one button, the child element should be a [Control](/javascript/api/manifest/control) element. If the add-in adds more than one button, the child element should be a [Group](/javascript/api/manifest/group) element that contains multiple `Control` elements.
- There is no `Menu` type equivalent for the `Control` element.
- The [Supertip](/javascript/api/manifest/supertip) element is not used.
- The required icon sizes are different. Mobile add-ins minimally must support 25x25, 32x32 and 48x48 pixel icons.

## Code considerations

Designing an add-in for mobile introduces some additional considerations.

### Use REST instead of Exchange Web Services

The [Office.context.mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method is not supported in Outlook Mobile. Add-ins should prefer to get information from the Office.js API when possible. If add-ins require information not exposed by the Office.js API, then they should use the [Outlook REST APIs](/outlook/rest/) to access the user's mailbox.

Mailbox requirement set 1.5 introduced a new version of [Office.context.mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) that can request an access token compatible with the REST APIs, and a new [Office.context.mailbox.restUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties) property that can be used to find the REST API endpoint for the user.

### Pinch zoom

By default users can use the "pinch zoom" gesture to zoom in on task panes. If this does not make sense for your scenario, be sure to disable pinch zoom in your HTML.

### Close task panes

In Outlook Mobile, task panes take up the entire screen and by default require the user to close them to return to the message. Consider using the [Office.context.ui.closeContainer](/javascript/api/office/office.ui#office-office-ui-closecontainer-member(1)) method to close the task pane when your scenario is complete.

### Compose mode and appointments

Currently, add-ins in Outlook Mobile only support activation when reading messages. Add-ins are not activated when composing messages or when viewing or composing appointments. However, there are two exceptions:

1. Online meeting provider integrated add-ins can be activated in Appointment Organizer mode. For more about this exception (including available APIs), refer to [Create an Outlook mobile add-in for an online-meeting provider](online-meeting.md#available-apis).
1. Add-ins that log appointment notes and other details to customer relationship management (CRM) or note-taking services can be activated in Appointment Attendee mode. For more about this exception (including available APIs), refer to [Log appointment notes to an external application in Outlook mobile add-ins](mobile-log-appointments.md#available-apis).

### Unsupported APIs

APIs introduced in requirement set 1.6 or later are not supported by Outlook Mobile. The following APIs from earlier requirement sets are also not supported.

- [Office.context.officeTheme](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context#officetheme-officetheme)
- [Office.context.mailbox.ewsUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties)
- [Office.context.mailbox.convertToEwsId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.convertToRestId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayMessageForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayNewAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.item.dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
- [Office.context.mailbox.item.displayReplyAllForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.displayReplyForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getEntities](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getEntitiesByType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getFilteredEntitiesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getRegexMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getRegexMatchesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)

## See also

[Requirement sets supported by Exchange servers and Outlook clients](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)