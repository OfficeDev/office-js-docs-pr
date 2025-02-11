---
title: Create Outlook add-ins for compose forms
description: Learn about scenarios and capabilities of Outlook add-ins for compose forms.
ms.date: 04/12/2024
ms.topic: overview
ms.localizationpriority: high
---

# Create Outlook add-ins for compose forms

You can create compose add-ins, which are Outlook add-ins activated in compose forms. In contrast with read add-ins (Outlook add-ins that are activated in read mode when a user is viewing a message or appointment), compose add-ins are available in the following user scenarios.

- Composing a new message, meeting request, or appointment in a compose form.

- Viewing or editing an existing appointment, or meeting item in which the user is the organizer.

- Composing an inline response message or replying to a message in a separate compose form.

- Editing a response (**Accept**, **Tentative**, or **Decline**) to a meeting request or meeting item.

- Proposing a new time for a meeting item.

- Forwarding or replying to a meeting request or meeting item.

In each of these scenarios, any add-in command buttons defined by the add-in are shown in compose form.

:::image type="content" source="../images/outlook-compose-form.png" alt-text="Sample add-in command buttons in compose form.":::

## Types of add-ins available in compose mode

Compose add-ins are implemented as [add-in commands](../design/add-in-commands.md). To activate add-ins for composing email or meeting responses, add-ins include a [MessageComposeCommandSurface extension point element](/javascript/api/manifest/extensionpoint#messagecomposecommandsurface) in the manifest. To activate add-ins for composing or editing appointments or meetings where the user is the organizer, add-ins include a [AppointmentOrganizerCommandSurface extension point element](/javascript/api/manifest/extensionpoint#appointmentorganizercommandsurface). For more information on manifests, see [Office Add-in manifests](../develop/add-in-manifests.md).

> [!NOTE]
> Add-ins developed for servers or clients that don't support add-in commands define activation rules using the [Rule](/javascript/api/manifest/rule) element contained in the [OfficeApp](/javascript/api/manifest/officeapp) element. Unless the add-in is being specifically developed for older clients and servers, new add-ins should use add-in commands.
>
> Add-ins that use activation rules aren't supported in an add-in that uses a [Unified manifest for Microsoft 365](../develop/unified-manifest-overview.md).

## API features available to compose add-ins

- [Add and remove attachments to an item in a compose form in Outlook](add-and-remove-attachments-to-an-item-in-a-compose-form.md)
- [Get, set, or add recipients when composing an appointment or message in Outlook](get-set-or-add-recipients.md)
- [Get or set the subject when composing an appointment or message in Outlook](get-or-set-the-subject.md)
- [Insert data in the body when composing an appointment or message in Outlook](insert-data-in-the-body.md)
- [Get or set the location when composing an appointment in Outlook](get-or-set-the-location-of-an-appointment.md)
- [Get or set the time when composing an appointment in Outlook](get-or-set-the-time-of-an-appointment.md)
- [Manage the sensitivity label of your message or appointment in compose mode](sensitivity-label.md)
- [Manage the delivery date and time of a message](delay-delivery.md)

## See also

- [Get Started with Outlook add-ins for Office](../quickstarts/outlook-quickstart-yo.md)
