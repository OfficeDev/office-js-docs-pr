# Office Add-in design guidelines

Enhance the user experience in your Office Add-in by developing UI that matches the Office voice, and apply accessibility guidelines to ensure that your add-in is accessible to all users.

If you plan to make your add-in [available in the Office Store](https://dev.office.com/officestore/docs/submit-to-the-office-store), make sure that your language and content complies with the [Validation policies](https://dev.office.com/officestore/docs/validation-policies).

## Voice guidelines 

As you design your Office Add-ins, consider the voice that you use in your UI text and elements. Strive to match the voice and tone of the Office UI, which is conversational, engaging, and accessible to users. 

To align your text with the principles of the Office voice:

- **Use a natural style.** Write the way that you speak. Avoid jargon and overly technical words and phrases. Use terms that are familiar to your users.
- **Use simple, direct language.** Use short words and sentences, and active voice in your text. 
- **Be consistent.** Use the same words for the same concepts throughout.
- **Engage the user.** Address the user as "you". Avoid using third person. Use imperatives for user tasks.
- **Be helpful and empathetic.** Make your text positive, polite, supportive, and encouraging. Emphasize what users can accomplish ―- not what they can't.
- **Know your customers.** Be mindful of cultural considerations and globalization when you use idioms or colloquialisms.

## Accessibility guidelines

As you design and develop your Office Add-ins, you'll want to ensure that all potential users and customers are able to use your add-in successfully. Apply the following guidelines to ensure that your solution is accessible to all audiences.

### Design for multiple input methods

- Ensure that users can perform operations by using only the keyboard. Users should be able to move to all actionable elements on the page by using a combination of the Tab and arrow keys.
- On a mobile device, when users operate a control by touch, the device should provide useful audio feedback.
- Provide helpful labels for all interactive controls. 

### Make your add-in easy to use

- Don't rely on a single attribute, such as color, size, shape, location, orientation, or sound, to convey meaning in your UI.
- Avoid unexpected changes of context, such as moving the focus to a different UI element without user action.
- Provide a way to verify, confirm, or reverse all binding actions.
- Provide a way to pause or stop media, such as audio and video.
- Do not impose a time limit for user action.

### Make your add-in easy to see

- Avoid unexpected color changes.
- Provide meaningful and timely information to describe UI elements, titles and headings, inputs, and errors. Ensure that names of controls adequately describe the intent of the control.
- Follow [standard guidelines](http://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html) for color contrast.

### Account for assistive technologies

- Avoid using features that interfere with assistive technologies, including visual, audio, or other interactions.
- Do not provide text in an image format. Screen readers cannot read text within images.
- Provide a way for users to adjust or mute all audio sources.
- Provide a way for users to turn on captions or audio description with audio sources.
- Provide alternatives to sound as a means to alert users, such as visual cues or vibrations.

### Accessibility resources

- [Web Content Accessibility Guidelines (WCAG) 2.0](http://www.w3.org/TR/wcag2ict/#REF-WCAG20)
- [Guidance on Applying WCAG 2.0 to Non-Web Information and Communications Technologies (WCAG2ICT)](http://www.w3.org/TR/wcag2ict/)
- [European Standard on accessibility requirements for Information and Communication Technologies (ICT)](http://www.etsi.org/deliver/etsi_en/301500_301599/301549/01.00.00_20/en_301549v010000c.pdf) 



