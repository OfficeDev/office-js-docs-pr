# UX design patterns for Office Add-ins. 

When designing Office Add-ins, the UX design of your add-in should provide compelling experiences that extend Office. To create a great add-in, your add-in should provide a first-run experience, a first-class UX experience, and smooth transitions between pages, among other things. Providing a clean, modern UX experience increases user retention and adoption of your add-in. 

The [UX design patterns for Office Add-ins project](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code "UX design patterns for Office Add-ins project") provides developers with resources to create great looking Office Add-ins that address common customer scenarios. Available in the UX design patterns project are HTML, JavaScript and CSS files that save developers the time involved in creating the UX for their add-in. Instead, developers can focus on writing code to meet their business requirements. 

Use the resources in the UX design patterns project to:

* Adhere to Office add-in best practices.
* Use Office Fabric components and styles.
* Build add-ins that look like a natural extension of the default Office UI.  

## How do I get started with the UX design patterns project?

To get started with the UX design patterns project:

1. Review the UX patterns described in this topic and decide which ones are important to your add-in. For example, pick one of the first-run experiences.
2. (Optional) Review the [UX designer specifications](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Source%20Files) and use them as a guide when creating your own UX design. 
3. Clone the [UX design patterns for Office Add-ins project repo](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code "UX design patterns for Office Add-ins project"). 
4. Copy the [assets folder](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/assets), and the code folder for the individual pattern you choose to your add-in project.  
5. Incorporate the individual pattern into your add-in. For example:
	- Edit the source location or add-in command URL in the manifest.
	- Use the UX design pattern as a template for other pages.
	- Link to/from the UX design pattern.
6. Run your add-in.

## Types of UX design patterns
### General landing pages 

General landing pages can be applied to any page in your add-in and generally don't have a special purpose. An example of a special purpose page, would be any of the first-run patterns. The following list describes the general landing pages available:

* **Landing (or generic) page** is a standard add-in page. Users may be redirected to a landing page after a first-run experience or sign-in process. For more information on the landing page, see the [specification](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Helpful%20Templates/AddIn_Template_Standard_Layout.pdf "PDF"), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/landing-page).
* **Brand image in brand bar** is the landing page with an image in the footer that represents your brand. For more information on the Brand image in brand bar, see the [specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Brand_Bar.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/brand-bar).

<table>
 <tr><th>Landing</th><th>Brand Bar</th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/landing-page"><img src="../../images/landing.page.PNG" alt="landing page" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/brand-bar"><img src="../../images/brand.bar.PNG" alt="brand bar" style="width: 264px;"/></A></td></tr>
 </table>

### First-Run

A first-run experience is the experience a user has when opening your add-in for the first time. The following lists the first-run design patterns you can include in your add-in. 

* **Steps to start** provides users with an ordered list of steps to perform to get started using your add-in. For more information on Steps to start, see the [specification](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_StepsToStart.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/instruction-step).
* **Value** communicates your add-in's value proposition. For more information on the Value pattern, see the [specification](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_ValuePlacemat.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/value-placemat).
* **Video** shows users a video before they start using your add-in. For more information on the Video pattern, see the [specification](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_VideoPlacemat.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat).
* **Walkthrough** takes users through a series of features or information before they start using the add-in. For more information on the Walkthrough pattern, see the [specification](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_PagingPanel.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/walkthrough).
* The [Office Store](https://msdn.microsoft.com/en-us/library/office/jj220033.aspx) has a system that manages trial versions of an add-in, but if you want to control the UI of your addin's trial experience, use the following patterns:

	* **Trial** shows users how to get started with a trial version of your add-in. For more information on the Trial pattern, see the [specification](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_TrialVersion.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat).
	* **Trial feature** advises users that the feature they are trying to use is not available in the trial version of the add-in. Alternatively, if your add-in is free but there's a feature in it that requires a subscription, you should consider using this pattern. You might also consider using this pattern to provide a downgraded experience after a trial has ended. For more information on the Trial feature pattern, see the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat-feature).

> Important: If you decide to manage your own trial, and not use the Office Store to manage the trial, ensure you include the **Additional purchase may be required** tag in the testing notes in the seller dashboard.

Consider whether showing users the first-run experience once or many times is important to your scenario. For example, if users use your add-in periodically, they may forget how to use the add-in. Seeing the first-run experience again may be helpful to those users. 

 <table>
 <tr><th>Steps to Start</th><th>Value</th><th>Video</th></tr>
 <tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/instruction-step"><img src="../../images/instruction.step.PNG" alt="instruction steps" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/value-placemat"><img src="../../images/value.placemat.PNG" alt="value placemat" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat"><img src="../../images/video.placemat.PNG" alt="video placemat" style="width: 264px;"/></A></td></tr>
 </table>

 <table>
 <tr><th>Walkthrough first page</th><th>Trial</th><th>Trial feature</th></tr>
 <tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/walkthrough"><img src="../../images/walkthrough1.PNG" alt="walkthrough 1" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat"><img src="../../images/trial.placemat.PNG" alt="trial placemat" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat-feature"><img src="../../images/trial.placemat.feature.PNG" alt="trial placemat feature" style="width: 264px;"/></A></td></tr>
 </table> 

### Notifications

There are a variety of ways that your add-in can notify users of events, such as errors, or of progress. The following lists the notification techniques you can use in your add-in. 

* **Embedded dialog**  shows a dialog inside the task pane that provides information and, optionally, an interactive experience, using buttons or other controls. Consider using one to prompt a user to confirm an action. Use the Embedded dialog pattern when you want to keep the user experience in the task pane. For more information on the Embedded dialog pattern, see the [specification](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Embedded_Dialog.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/embedded-dialog).
* **Inline message** indicates error, success, or information, and it can appear at a specified location in the task pane. For example, if a user enters an improperly formatted email address in a text box, an error message appears just below the text box. For more information about Inline messages, see the [specification](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_Inline_Message.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/inline-message).
* **Message banner** provides information and, optionally, a simple call to action, in a banner that can be collapsed to a single line, expanded to multiple lines, or dismissed. Consider using message banners to report a service update or a helpful tip when the add-in starts. For more information on Message banner, see the [specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Notification_MessageBanner.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/message-banner).
* **Progress bar** indicates the progress of a long-running, synchronous process, such as a configuration task that must complete before the user can take any further action. It is a separate interstitial page that also reinforces the add-in brand. Use a progress bar when the process can send periodic measures of how far along it is back to the add-in. For more information on Progress bar, see the [specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Notification_Progress.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/progress-bar).
* **Spinner** indicates that a long-running, synchronous process is underway, but provides no indication of how far along it is. It is a separate interstitial page that also reinforces the add-in brand. Use a spinner when the add-in cannot know reliably how far along a process is. For more information on Spinner, see the [specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Notification_Progress.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/spinner).
* **Toast** provides a brief message that fades away after a few seconds. Since the user might not see the message, use toast only for inessential information. It is a good choice for notifying users of an event in a remote system, such as the receipt of an email. For more information on Toast, see the [specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Notification_Toast.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/toast).

 <table>
 <tr><th>Embedded dialog</th><th>Inline message</th><th>Message banner</th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/embedded-dialog"><img src="../../images/embedded.dialog.PNG" alt="embedded dialog" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/inline-message"><img src="../../images/inline.message.PNG" alt="inline message" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/message-banner"><img src="../../images/message.banner.PNG" alt="message banner" style="width: 264px;"/></A></td></tr>
 </table>

 <table>
 <tr><th>Progress bar</th><th>Spinner</th><th>Toast</th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/progress-bar"><img src="../../images/progress.bar.PNG" alt="progress bar" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/spinner"><img src="../../images/spinner.PNG" alt="spinner" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/toast"><img src="../../images/toast.PNG" alt="toast" style="width: 264px;"/></A></td></tr>
 </table>

### Navigation

Navigation patterns provide a way to organize different pages and commands in your add-in. The following list of navigations can be included in your add-in.

* **Back button navigation** shows a task pane with a Back button. Use the Back button to navigate back to the previous page. For more information on the Back button pattern, see the [specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Back_Button.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/back-button).
* **Navigation** shows a hamburger menu. For more information on the Navigation pattern, see the [specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Navigation.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation).
* **Navigation with commands** shows a hamburger menu with commands. Use the Navigation with commands to allow users to navigate between different content and have commands. For more information on the Navigation with commands pattern, see the [specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Navigation_%26_Commands.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation-commands).
* **Pivot** shows a Pivot navigation. Use the Pivot navigation to allow users to navigate between different content. For more information on the Pivot navigation pattern, see the [specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Pivot.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/pivot).
* **Tab bar** shows a top-level navigation. Use the Tab bar to allow users to provide tabs with short and descriptive titles that communicate the tab's functionality. For more information on the Tab bar pattern, see the [specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Tab_Bar.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/tab-bar).


<table>
<tr><th>Back button</th><th>Navigation</th><th>Navigation with commands</th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/back-button">
<img src="../../images/back.button.png" alt="back button" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation">
<img src="../../images/navigation.png" alt="navigation" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation-commands">
<img src="../../images/navigation.commands.png" alt="navigation with commands" style="width: 264px;"/></A></td></tr>
</table>

<table>
<tr><th>Pivot</th><th>Tab bar</th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/pivot">
<img src="../../images/pivot.png" alt="pivot navigation" style="width: 300px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/tab-bar">
<img src="../../images/tab.bar.png" alt="tab bar" style="width: 264px;"/></A></td>
</tr></tr> 
</table>

### General components

The following are general components that you can use in your add-ins in a variety of scenarios.  

#### Client dialogs

Client dialogs provide another way for users to work with your add-in that's not limited to a task pane. The following list of dialogs can be included in your add-in.

* **Typeramp dialog** shows a dialog box with textual content. Use the typeramp dialog to display elaborative information to users. For more information on the Typeramp dialog pattern, see the [specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Client_Dialog.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/typeramp).
* **Alert dialog** shows an alert box with important information, like errors or notifications, to users. For more information on the Alert dialog pattern, see the [specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Client_Dialog.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/alert).
* **Navigation dialog** shows a dialog box with navigation. Use the navigation dialog to allow users to navigate between different content. For more information on the Navigation dialog pattern, see the [specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Client_Dialog.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation).

<table>
 <tr><th>Typeramp dialog</th><th>Alert dialog</th></tr>
<tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/typeramp"><img src="../../images/typeramp.dialog.png" alt="typeramp dialog" style="width: 300px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/alert"><img src="../../images/alert.dialog.png" alt="alert dialog" style="width: 264px;"/></A></td>
</tr></tr>
 </table>
 
 <table>
 <tr><th>Navigation dialog</th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation"><img src="../../images/navigation.dialog.png" alt="navigation dialog" style="width: 300px;"/></A></td></tr>
</tr>
 </table>


#### Feedback and Ratings

To improve the visibility and adoption of your add-in, you should provide users with the ability to  rate and review your add-in in the Office Store. This pattern demonstrates how to present feedback and ratings from within the add-in using two techniques:

- User initiated feedback - a user chooses to send feedback using either the navigation menu (for example, using the **Send Feedback** link) or icon on the footer.
- System initiated feedback - after the add-in runs 3 times, a user is prompted to provide feedback using a Message Banner.

Using either technique opens a dialog of the add-in's page in the Office Store. For more information on the Feedback and ratings pattern, see the [specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Notification_Feedback.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/feedback/office-store).

>Important: This pattern currently points to the home page of the Office Store. Ensure you update this URL to your add-in's page in the Office Store.

 <table>
 <tr><th>Feedback and ratings</th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/feedback/office-store"><img src="../../images/feedback.ratings.PNG" alt="Feedback and Ratings" style="width: 264px;"/></A></td></tr>
</tr>
 </table>

#### Settings

You can provide a Settings page and a Privacy page to include common components that may be used in your add-in, and to also notify the user of any privacy policies that your add-in adheres to. 

* **Settings** shows a task pane with common components. Use the Settings page to display provide user options. For more information on the Settings pattern, see the [specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Settings.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings).
* **Privacy Policy** shows task pane with important information about privacy policies. For more information on the Privacy pattern, see the [specification](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Settings.md), and the [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings).

<table>
 <tr><th>Settings</th><th>Privacy Policy</th></tr>
<tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings"><img src="../../images/settings.png" alt="settings" style="width: 300px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings"><img src="../../images/privacy.policy.png" alt="privacy" style="width: 264px;"/></A></td>
</tr></tr>
 </table>

## Additional resources

* [Best practices for developing Office Add-ins](https://dev.office.com/docs/add-ins/overview/add-in-development-best-practices)
* [Office UI Fabric](http://dev.office.com/fabric/)
