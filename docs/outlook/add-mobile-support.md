---
title: Add support for add-in commands in Outlook on mobile devices
description: Learn how to add support for Outlook on mobile devices including how to update the add-in manifest and change your code for mobile scenarios, if necessary.
ms.date: 01/31/2025
ms.localizationpriority: medium
---

# Add support for add-in commands in Outlook on mobile devices

Implement add-in commands in Outlook on mobile devices to access the same functionality (with some [limitations](#code-considerations)) that you already have in Outlook on the web, on Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic), and on Mac. Adding support for Outlook mobile requires updating the add-in manifest and possibly changing your code for mobile scenarios.

## Update the manifest

The first step to enabling add-in commands in Outlook mobile is to define them in the add-in manifest.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

1. In the [`"extensions.ribbons.requirements.formFactors"`](/microsoft-365/extensibility/schema/requirements-extension-element#formfactors) array, add `"mobile"` as an item. When you are finished, the array should look like the following.

    ```json
    "formFactors": [
        "mobile",
        <!-- Typically, there'll be other form factors listed. -->
    ]
    ```

1. If your add-in uses Appointment Attendee mode, such as an add-in that integrates a provider of a note-taking or customer relationship management (CRM) application, add `"logEventMeetingDetailsAttendee"` to the [`"extensions.ribbons.contexts"`](/microsoft-365/extensibility/schema/extension-ribbons-array#contexts) array. The following is an example.

    ```json
    "contexts": [
        "meetingDetailsAttendee",
        "logEventMeetingDetailsAttendee"
    ],
    ```

1. If your add-in uses an integrated online meeting provider, add `"onlineMeetingDetailsOrganizer"` to the `"extensions.ribbons.contexts"` array. The following is an example.

    ```json
    "contexts": [
        "meetingDetailsOrganizer",
        "onlineMeetingDetailsOrganizer"
    ],
    ```

1. In the [`"extensions.ribbons.tabs"`](/microsoft-365/extensibility/schema/extension-ribbons-array#tabs) array, find the tab with the `"builtInTabId"` of `"TabDefault"`. Add a child `"customMobileRibbonGroups"` array to it (as a peer of the existing `"groups"` property). Inside this array, create an object and do the following:

   - Set appropriate `"id"` and `"label"` values.
   - Create an object in the `"controls"` array to represent a button and configure it as follows.
      - Set appropriate `"id"` and `"label"` values. To ensure that the button fits correctly in the ribbon, we recommend that you limit the `"label"` to 16 characters.
      - Set `"type"` to `"mobileButton"`.
      - Assign a function to the `"actionId"` property. This should match the `"id"` of the object in the `"extensions.runtimes.actions"` array.
      - Be sure you have all nine required icons.
  
   The following is an example.

    ```json
    "tabs": [
        {
            "builtInTabId": "TabDefault",
            "groups": [
                <-- Non-mobile group objects omitted. -->
            ],
            "customMobileRibbonGroups": [
                {
                    "id": "mobileApptComposeGroup",
                    "label": "Contoso Meeting",
                    "controls": [
                        { 
                            "id": "mobileInsertMeetingButton",
                            "label": "Add meeting",
                            "type": "mobileButton",
                            "actionId": "insertContosoMeeting",
                            "icons": [
                                {
                                    "scale": 1,
                                    "size": 25,
                                    "url": "https://contoso.com/assets/icon-25.png"
                                },
                                {
                                    "scale": 1,
                                    "size": 32,
                                    "url": "https://contoso.com/assets/icon-32.png"
                                },
                                {
                                    "scale": 1,
                                    "size": 48,
                                    "url": "https://contoso.com/assets/icon-48.png"
                                },                                
                                {
                                    "scale": 2,
                                    "size": 25,
                                    "url": "https://contoso.com/assets/icon-25.png"
                                },
                                {
                                    "scale": 2,
                                    "size": 32,
                                    "url": "https://contoso.com/assets/icon-32.png"
                                },
                                {
                                    "scale": 2,
                                    "size": 48,
                                    "url": "https://contoso.com/assets/icon-48.png"
                                },                                
                                {
                                    "scale": 3,
                                    "size": 25,
                                    "url": "https://contoso.com/assets/icon-25.png"
                                },
                                {
                                    "scale": 3,
                                    "size": 32,
                                    "url": "https://contoso.com/assets/icon-32.png"
                                },
                                {
                                    "scale": 3,
                                    "size": 48,
                                    "url": "https://contoso.com/assets/icon-48.png"
                                }
                            ]
                        }
                    ]
                }
            ]
        }
    ]    
    ```

# [Add-in only manifest](#tab/xmlmanifest)

The [VersionOverrides](/javascript/api/manifest/versionoverrides) v1.1 schema defines a new form factor for mobile, [MobileFormFactor](/javascript/api/manifest/mobileformfactor). The **\<MobileFormFactor\>** element contains all of the information for loading the add-in in mobile clients. This way, you can define completely different UI elements and JavaScript files for the mobile experience.

The following example shows a single task pane button in a **\<MobileFormFactor\>** element. This is very similar to the elements that appear in a [DesktopFormFactor](/javascript/api/manifest/desktopformfactor) element, with some notable differences.

- The [OfficeTab](/javascript/api/manifest/officetab) element isn't used.
- The [ExtensionPoint](/javascript/api/manifest/extensionpoint) element must have only one child element. If your add-in implements the [MobileOnlineMeetingCommandSurface](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface) or [MobileLogEventAppointmentAttendee](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee) extension point, you must include a [Control](/javascript/api/manifest/control) child element. If your add-in implements the [MobileMessageReadCommandSurface](/javascript/api/manifest/extensionpoint#mobilemessagereadcommandsurface) extension point, you must include a [Group](/javascript/api/manifest/group) child element that contains multiple **\<Control\>** elements.
- There is no `Menu` type equivalent for the **\<Control\>** element.
- The [Supertip](/javascript/api/manifest/supertip) element isn't used.
- The required icon sizes are different. Mobile add-ins minimally must support 25x25, 32x32 and 48x48 pixel icons. For more information, see [Additional requirements for mobile form factors](/javascript/api/manifest/icon#additional-requirements-for-mobile-form-factors).

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
        ...
        <Hosts>
            <Host xsi:type="MailHost">
                ...
                <MobileFormFactor>
                    <FunctionFile resid="residUILessFunctionFileUrl" />
                    <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
                        <Group id="mobileMsgRead">
                            <Label resid="groupLabel" />
                            <Control xsi:type="MobileButton" id="TaskPaneBtn">
                                <Label resid="residTaskPaneButtonName" />
                                <Icon xsi:type="bt:MobileIconList">
                                    <bt:Image size="25" scale="1" resid="icon_25" />
                                    <bt:Image size="25" scale="2" resid="icon_25" />
                                    <bt:Image size="25" scale="3" resid="icon_25" />
                        
                                    <bt:Image size="32" scale="1" resid="icon_32" />
                                    <bt:Image size="32" scale="2" resid="icon_32" />
                                    <bt:Image size="32" scale="3" resid="icon_32" />
                        
                                    <bt:Image size="48" scale="1" resid="icon_48" />
                                    <bt:Image size="48" scale="2" resid="icon_48" />
                                    <bt:Image size="48" scale="3" resid="icon_48" />
                                </Icon>
                                <Action xsi:type="ShowTaskpane">
                                    <SourceLocation resid="residTaskpaneUrl" />
                                </Action>
                            </Control>
                        </Group>
                    </ExtensionPoint>
                </MobileFormFactor>
            </Host>
        </Hosts>
        ...
    </VersionOverrides>
</VersionOverrides>
```

---

## Code considerations

Designing an add-in for mobile introduces some additional considerations.

### Use Microsoft Graph

Add-ins should prefer to get information from the Office.js API when possible. If your add-in requires information not exposed by the Office.js API, use [Microsoft Graph](/graph/overview) to access the user's mailbox.

### Pinch zoom

By default users can use the "pinch zoom" gesture to zoom in on task panes. If this doesn't make sense for your scenario, be sure to disable pinch zoom in your HTML.

### Close task panes

In Outlook mobile, task panes take up the entire screen and by default require the user to close them to return to the message. Consider using the [Office.context.ui.closeContainer](/javascript/api/office/office.ui#office-office-ui-closecontainer-member(1)) method to close the task pane when your scenario is complete.

### Compose mode and appointments

Currently, add-ins in Outlook mobile only support activation when reading messages. Add-ins aren't activated when composing messages or when viewing or composing appointments. However, there are some exceptions.

1. Online meeting provider integrated add-ins activate in Appointment Organizer mode. For more information about this exception (including available APIs), see [Create an Outlook mobile add-in for an online-meeting provider](online-meeting.md#available-apis).
1. Add-ins that log appointment notes and other details to customer relationship management (CRM) or note-taking services activate in Appointment Attendee mode. For more information about this exception (including available APIs), see [Log appointment notes to an external application in Outlook mobile add-ins](mobile-log-appointments.md#available-apis).
1. Event-based add-ins activate when the `OnNewMessageCompose` event occurs. For more information about this exception (including additional supported APIs), see [Implement event-based activation in Outlook mobile add-ins](mobile-event-based.md).

### Supported APIs

Although Outlook mobile supports up to [Mailbox requirement set 1.5](/javascript/api/outlook?view=outlook-js-1.5&preserve-view=true), you can now implement additional APIs from later requirement sets to further extend the capability of your add-in on Outlook mobile. For guidance on which APIs you can implement in your mobile add-in, see [Outlook JavaScript APIs supported in Outlook on mobile devices](outlook-mobile-apis.md).

## See also

- [Requirement sets supported by Exchange servers and Outlook clients](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)
- [Outlook JavaScript APIs supported in Outlook on mobile devices](outlook-mobile-apis.md)
