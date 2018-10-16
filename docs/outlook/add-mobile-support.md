---
title: Add mobile support to an Outlook add-in | Microsoft Docs
description: Adding support for Outlook Mobile requires updating the add-in manifest and possibly changing your code for mobile scenarios.
author: jasonjoh
ms.topic: article
ms.technology: office-add-ins
ms.date: 06/13/2017
ms.author: jasonjoh
---

# Add support for add-in commands for Outlook Mobile

Using add-in commands in Outlook Mobile allows your users to access the same functionality (with some [limitations](#code-considerations)) that they already have in Outlook for Windows, Outlook for Mac, and Outlook on the web. Adding support for Outlook Mobile requires updating the add-in manifest and possibly changing your code for mobile scenarios.

## Updating the manifest

The first step to enabling add-in commands in Outlook Mobile is to define them in the add-in manifest. The **VersionOverrides** v1.1 schema defines a new form factor for mobile, [MobileFormFactor](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/mobileformfactor).

This element contains all of the information for loading the add-in in mobile clients. This enables you to define completely different UI elements and JavaScript files for the mobile experience.

The following example shows a single taskpane button in a **MobileFormFactor** element.

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

This is very similar to the elements that appear in a [DesktopFormFactor](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/desktopformfactor) element, with some notable differences.

- The [OfficeTab](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/officetab) element is not used.
- The [ExtensionPoint](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/extensionpoint) element must have only one child element. If the add-in only adds one button, the child element should be a [Control](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/control) element. If the add-in adds more than one button, the child element should be a [Group](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/group) element that contains multiple `Control` elements.
- There is no `Menu` type equivalent for the `Control` element.
- The [Supertip](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/supertip) element is not used.
- The required icon sizes are different. Mobile add-ins minimally must support 25x25, 32x32 and 48x48 pixel icons.

## Code considerations

Designing an add-in for mobile introduces some additional considerations.

### Use REST instead of Exchange Web Services

The [Office.context.mailbox.makeEwsRequestAsync](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox#makeewsrequestasyncdata-callback-usercontext) method is not supported in Outlook Mobile. Add-ins should prefer to get information from the Office.js API when possible. If add-ins require information not exposed by the Office.js API, then they should use the [Outlook REST APIs](https://docs.microsoft.com/outlook/rest/) to access the user's mailbox.

Mailbox requirement set 1.5 introduces a new version of [Office.context.mailbox.getCallbackTokenAsync](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox#getcallbacktokenasyncoptions-callback) that can request an access token compatible with the REST APIs, and a new [Office.context.mailbox.restUrl](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox#resturl-string) property that can be used to find the REST API endpoint for the user.

### Pinch zoom

By default users can use the "pinch zoom" gesture to zoom in on taskpanes. If this does not make sense for your scenario, be sure to disable pinch zoom in your HTML.

### Close taskpanes

In Outlook Mobile, taskpanes take up the entire screen and by default require the user to close them to return to the message. Consider using the [Office.context.ui.closeContainer](https://docs.microsoft.com/javascript/api/office/office.ui#closecontainer--) method to close the taskpane when your scenario is complete.

### Compose mode and appointments

Currently add-ins in Outlook Mobile only support activation when reading messages. Add-ins are not activated when composing messages or when viewing or composing appointments.

### Unsupported APIs

The following APIs are not supported by Outlook Mobile.

  - [Office.context.officeTheme](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context#officetheme-object)
  - [Office.context.mailbox.ewsUrl](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox#ewsurl-string)
  - [Office.context.mailbox.convertToEwsId](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox#converttoewsiditemid-restversion--string)
  - [Office.context.mailbox.convertToRestId](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox#converttorestiditemid-restversion--string)
  - [Office.context.mailbox.displayAppointmentForm](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox#displayappointmentformitemid)
  - [Office.context.mailbox.displayMessageForm](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox#displaymessageformitemid)
  - [Office.context.mailbox.displayNewAppointmentForm](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox#displaynewappointmentformparameters)
  - [Office.context.mailbox.makeEwsRequestAsync](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox#makeewsrequestasyncdata-callback-usercontext)
  - [Office.context.mailbox.item.dateTimeModified](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#datetimemodified-date)
  - [Office.context.mailbox.item.displayReplyAllForm](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#displayreplyallformformdata)
  - [Office.context.mailbox.item.displayReplyForm](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#displayreplyformformdata)
  - [Office.context.mailbox.item.getEntities](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#getentities--entitiesjavascriptapioutlook15officeentities)
  - [Office.context.mailbox.item.getEntitiesByType](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion)
  - [Office.context.mailbox.item.getFilteredEntitiesByName](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion)
  - [Office.context.mailbox.item.getRegexMatches](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#getregexmatches--object)
  - [Office.context.mailbox.item.getRegexMatchesByName](https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/Office.context.mailbox.item#getregexmatchesbynamename--nullable-array-string-)