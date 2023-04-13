---
title: Accessibility guidelines for Office Add-ins
description: Learn how to make your Office Add-in accessible to all users.
ms.date: 09/24/2018
ms.topic: best-practice
ms.localizationpriority: medium
---

# Accessibility guidelines

As you design and develop your Office Add-ins, you'll want to ensure that all potential users and customers are able to use your add-in successfully. Apply the following guidelines to ensure that your solution is accessible to all audiences.

## Design for multiple input methods

- Ensure that users can perform operations by using only the keyboard. Users should be able to move to all actionable elements on the page by using a combination of the Tab and arrow keys.
- On a mobile device, when users operate a control by touch, the device should provide useful audio feedback.
- Provide helpful labels for all interactive controls. 

## Make your add-in easy to use

- Don't rely on a single attribute, such as color, size, shape, location, orientation, or sound, to convey meaning in your UI.
- Avoid unexpected changes of context, such as moving the focus to a different UI element without user action.
- Provide a way to verify, confirm, or reverse all binding actions.
- Provide a way to pause or stop media, such as audio and video.
- Do not impose a time limit for user action.

## Make your add-in easy to see

- Avoid unexpected color changes.
- Provide meaningful and timely information to describe UI elements, titles and headings, inputs, and errors. Ensure that names of controls adequately describe the intent of the control.
- Follow [standard guidelines](https://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html) for color contrast.

## Account for assistive technologies

- Avoid using features that interfere with assistive technologies, including visual, audio, or other interactions.
- Do not provide text in an image format. Screen readers cannot read text within images.
- Provide a way for users to adjust or mute all audio sources.
- Provide a way for users to turn on captions or audio description with audio sources.
- Provide alternatives to sound as a means to alert users, such as visual cues or vibrations.

## See also

- [Web Content Accessibility Guidelines (WCAG) 2.0](https://www.w3.org/TR/wcag2ict/#REF-WCAG20)
- [Guidance on Applying WCAG 2.0 to Non-Web Information and Communications Technologies (WCAG2ICT)](https://www.w3.org/TR/wcag2ict/)
- [European Standard on accessibility requirements for Information and Communication Technologies (ICT)](https://www.etsi.org/deliver/etsi_en/301500_301599/301549/01.00.00_20/en_301549v010000c.pdf) 
