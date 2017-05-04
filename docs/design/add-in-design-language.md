# Office Add-in design language

The Office design language is a clean and simple visual system that ensures consistency across experiences. It contains a set of visual elements that define Office interfaces, including: 

- A common typeface
- A standard color palette
- A set of typographic sizes and weights
- Icon guidelines
- Shared icon assets
- Animation definitions
- Common components

Office UI Fabric is the official front-end framework for building with the Office design language. Using Fabric is optional, but it is the fastest way to ensure that your add-ins feel like a natural extension of Office. Take advantage of Fabric to design and build add-ins that complement Office.

Many Office Add-ins are associated with a preexisting brand. You can retain a strong brand and its visual or component language in your add-in. Look for opportunities to retain your own visual language while integrating with Office. Consider ways to swap out Office colors, typography, icons, or other stylistic elements with elements of your own brand. Consider ways to follow common add-in layouts or UX design patterns while inserting controls and components that are familiar to your customers.

Inserting a heavily branded HTML-based UI inside of Office can create dissonance for customers. Find a balance that fits seamlessly in Office but also clearly aligns with your service or parent brand. When an add-in does not fit with Office, it's often because stylistic elements conflict. For example, typography is too large and off grid, colors are contrasting or particularly loud, or animations are superfluous and behave differently than Office. The appearance and behavior of controls or components veer too far from Office standards.

## Typography
Segoe is the standard typeface for Office. Use it in your add-in to align with Office task panes, dialog boxes, and content objects. Office UI Fabric gives you access to Segoe. It provides a full type ramp of Segoe with many variations - across font weight and size - in convenient CSS classes. Not all Office UI Fabric sizes and weights will look great in an Office Add-in. To fit harmoniously or avoid conflicts, consider using a subset of the Fabric type ramp. Here's a list of Fabric's base classes that we recommend for use in Office Add-ins.

|Sample |Class |Size |Weight |Recommended Usage |
|------ |----- |---- |------ |----------------- |
|img |.ms-font-xxl |28 px | Segoe Light |<ul><li>This class is larger than all other typographic elements in Office. Use it sparingly to avoid unseating visual hierarchy.</li><li>Avoid use on long strings in constrained spaces.</li><li>Provide ample whitespace around text using this class.</li><li>Commonly used for first run messages, hero elements, or other calls to action.</li></ul> |
| img |.ms-font-xl |21 px |Segoe Light | <ul><li>This class matches the task pane title of Office applications.</li><li>Use it sparingly to avoid a flat typographic hierarchy.</li><li>Commonly used as the top-level element such as dialog box, page, or content titles.</li><li></ul> |
|img |.ms-font-l |17 px |Segoe Semilight | <ul><li>This class is the first stop below titles.</li><li>Commonly used as a subtitle, navigation element, or group header.</li><ul> |
|img |.ms-font-m |14 px |Segoe Regular |*Commonly used as body text within add-ins. |
|img |.ms-font-xs |11 px | Segoe Regular |*Commonly used for secondary or tertiary text such as timestamps, by lines, captions, or field labels. |
|img |.ms-font-mi |10 px |Segoe Semibold |*The smallest step in the type ramp should be used rarely. It's available for circumstances where legibility is not required. |
> Text color is not included in these base classes. Use Fabric's "neutral primary" for most text on white backgrounds.

## Color
Color is often used to emphasize brand and reinforce visual hierarchy. It helps identify an interface as well as guide customers through an experience. Inside Office, color is used for the same goals but it is applied purposefully and minimally. At no point does it overwhelm customer content. Even when each Office app is branded with its own dominant color, it is used to sparingly.

Office UI Fabric includes a set of default theme colors. When Fabric is applied to an Office Add-in as components or in layouts, the same goals apply. Color should communicate hierarchy, purposefully guiding customers to action without interfering with content. Fabric theme colors can introduce a new accent color to the overall interface. This new accent can conflict with Office app branding and interfere with hierarchy. In other words, Fabric can introduce a new accent color to the overall interface when used inside an add-in. This new accent color can distract and interfere with the overall hierarchy. Consider ways to avoid conflicts and interference. Use neutral accents or overwrite Fabric theme colors to match Office app branding or your own brand colors.

Office applications allow customers to personalize their interfaces by applying an Office UI theme. Customers can choose between four UI themes to vary styling of backgrounds and buttons in Word, PowerPoint, Excel and other apps in the Office suite. To make your add-ins feel like a natural part of Office and respond to personalization, use our Themeing APIs. For example, task pane background colors switch to a dark gray in some themes. Our theming APIs allow you to follow suit and adjust foreground text to ensure [accessibility](add-in-design-guideline.md#accessibility-guidelines).

>  For mail and task pane add-ins, use the [Context.officeTheme](https://dev.office.com/docs/reference/shared/office.context.officetheme.htm) property to match the theme of the Office applications. This API is currently only available in Office 2016.

> For PowerPoint content add-ins, see [Use Office themes in your PowerPoint add-ins](https://dev.office.com/docs/add-ins/powerpoint/use-document-themes-in-your-powerpoint-add-ins.htm).

Apply the following general guidelines for color:

* Use color sparingly to communicate hierarchy and reinforce brand.
* Overuse of a single accent color applied to both interactive and non-interactive elements can lead to confusion. For example, avoid using the same color for selected and unselected items in a navigation menu.
* Avoid unnecessary conflicts with Office branded app colors.
* Use your own brand colors to build association with your service or company.
* Ensure that all text is accessible. Be sure that there is a 4.5:1 constrast ratio between foreground text and background.
* Be aware of color blindness. Use more than just color to indicate interactivity and hierarchy.
* Refer to [icon guidelines](design-icons.md) to learn more about designing add-in command icons with the Office icon color pallet.

## Layout
Each HTML container embedded in Office will have a layout. These layouts are the main screens of your add-in. In them you will create experiences that enable customers to initiate actions, modify settings, view, scroll, or navigate content. Design your add-in with a consistent layouts across screens to guarantee continuity of experience. If you have an existing website that your customers are familiar with using, consider reusing layouts from your existing web pages. Adapt them to fit harmoniously within Office HTML containers.

For guidelines on layout, see [Task pane](task-pane-add-ins.md), [Content](content-add-ins.md), and [Dialog box](dialog-boxes.md). For more information about how to assemble Office UI Fabric components into common layouts and user experience flows, see [UX design patterns templates](ux-design-patterns.md).

Apply the following general guidelines for layouts:

*	Avoid narrow or wide margins on your HTML containers. 20 pixels is a great default. 
*	Align elements intentionally. Extra indents and new points of alignment should aid visual hierarchy.
*	Office interfaces are on a 4px grid. Aim to keep your padding between elements at multiples of 4. 
*	Overcrowding your interface can lead to confusion and inhibit ease of use with touch interactions. 
*	Keep layouts consistent across screens. Unexpected layout changes look like visual bugs that contribute to a lack of confidence and trust with your solution. 
*	Follow common layout patterns. Conventions help users understand how to use an interface.
*	Avoid redundant elements like branding or commands.
*	Consolidate controls and views to avoid requiring too much mouse movement. 
*	Create responsive experiences that adapt to HTML container widths and heights.

## Component language

Screens and layouts are composed of content and components. Components are controls that help your customers interact with elements of your software or service. Buttons, navigation, badges, alerts, and dropdowns are all examples of common components that often have consistent styles and behaviors.

Office UI Fabric renders components that look and behave like a part of Office. Take advantage of Fabric to integrate seamlessly with Office. If your add-in has its own preexisting component language, you don't need to discard it in favor of Fabric. Look for opportunities to retain it while integrating with Office. Consider ways to swap out stylistic elements, remove conflicts, or adopt styles and behaviors that remove user confusion.

Apply the following general guidelines for components:

*	Don’t replicate the Office ribbon inside your add-in
*	Avoid creating menus, buttons, or other components that behave differently from Office components.
*	Use the [Office UI Fabric](office-ui-fabric.md) components the we recommend for add-ins.
*	Use the [UX design patterns templates](ux-design-patterns.md) for common Office UI components. 

## Icons
Icons are the visual representation of a behavior or concept. They are often used to add meaning to controls and commands. Visuals, either realistic or symbolic, enable the user to navigate the UI the same way signs help users navigate their environment. They should be simple and clear, and contain only the necessary details, to enable customers to quickly parse what action will occur when they choose a control.

Office ribbon interfaces have a standard visual style. If you are designing an add-in command for the Office ribbon, follow our [icon guidelines](design-icons). This ensures consistency and familiarity across Office apps. The guidelines will help you design a set of PNG assets for your solution that fit in as a natural part of Office.

Many HTML containers contain controls with iconography. Use Office UI Fabric’s custom font to render Office styled icons in your add-in. Fabric’s icon font contains many glyphs for common Office metaphors that you can scale, color, and style to suit your needs. If you have an existing visual language with your own set of icons, feel free to use it in your HTML canvases. Building continuity with your own brand with a standard set of icons is an important part of any design language. Be careful to avoid creating confusion for customers by conflicting with Office metaphors.

Apply the following general guidelines for icons:

* Don’t repurpose Office UI Fabric glyphs for add-in commands in the Office ribbon or contextual menus. Fabric icons are stylistically different and will not match.
* Use the Office icon language to represent behaviors or concepts.
* Reuse common Office visual metaphors such as paintbrush for format or magnifying glass for find.
* Don’t misuse metaphors for unrelated actions. Using the same visual for a different behavior or concept can cause confusion for users.

## Animation
UI elements, controls, and components often have interactive behaviors that require transitions, motion, or animation. Common characteristics of motion across UI elements define the animation aspects of a design language. Because Office is focused on productivity, the Office animation language supports the goal of helping customers get things done. It strikes a balance between performant response, reliable choreography, and detailed delight.

Office UI fabric includes an animation library to control motion in your HTML containers. Use it to fit seamlessly in Office. It will help you create experiences that are more felt than observed. The animation CSS classes provide directionality, enter/exit, and duration specifics that reinforce Office mental models and provide opportunities for customers to learn how to interact with your add-in. 

If your add-in has its own animation language, use it. Look for opportunities to retain your branded animation while integrating with Office. Be careful not to interfere with or conflict with common motion patterns in Office. Avoid creating experiences that are embellishments that only distract your customers.

Apply the following general guidelines for animations:

* Animations should be felt, and experienced subconsciously, to avoid hindering task completion.
* Avoid anticipations, bounces, rubberband, or other effects that emulate natural world physics.
* Choreograph elements to reinforce hierarchy and mental models.
* Use motion to guide the user and provide compositional focus on key elements for task completion. 
* Consider the origin of your triggering element. Use motion to create a link between the action and the resulting UI.
* Consider tone and purpose of your content when choosing animations. Handle critical messages differently than exploratory navigations.
