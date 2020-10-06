1. Open the Office application on the web; for example, PowerPoint on the web.
1. Open the browser's developer tools by pressing F12.
1. Open the **Elements** tab in the tools.
1. Press Ctrl+Shift+C to turn on the element inspection feature.
1. Move the cursor to the ribbon button that you want to identify (or whose parent group you want to identify) and it will be highlighted. Be sure that you are highlighting the whole button and not just the icon image on it.
1. Left-click to select the control. The HTML in the **Elements** tab of the tools will expand and highlight the `<button>` element. The `id` attribute of the button is the Office control ID that you use in a `<OfficeControl>` element. (If there is no `id` attribute, use the `data-unique-id` attribute.)
1. The group to which the control belongs is the parent or grandparent `<div>`. Use the `id` or `data-unique-id` attribute of the `<div>` in an `<OfficeGroup>` element.

The following screenshot shows an example of this procedure. The Bold button has been highlighted. Note in the HTML that the `data-unique-id` of the button is `Ribbon-Bold` and that the `data-unique-id` of the parent `<div>` is `Font`.

![Screenshot of an Office on the web ribbon bar with the Bold button highlighted and corresponding button element highlighted in the HTML markup to the right.](../images/control-ids.png)
