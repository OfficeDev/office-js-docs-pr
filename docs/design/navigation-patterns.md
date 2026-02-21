---
title: Navigation patterns for Office Add-ins
description: Learn best practices for using command bars, tab bars, and back buttons to design the navigation of an Office Add-in.
ms.date: 02/20/2026
ms.topic: best-practice
ms.localizationpriority: medium
---

# Navigation patterns for Office Add-ins

Office Add-ins present unique navigation challenges due to their constrained task pane width, integration with the Office ribbon, and multiplatform requirements. Effective navigation ensures users can access features easily while maintaining context with their Office document.

This article covers navigation patterns within task panes and how to coordinate navigation between the Office ribbon and your add-in UI.

## Unique navigation challenges for Office Add-ins

When designing navigation for Office Add-ins, consider the following constraints.

- **Narrow task pane width**: Task panes range from 320 px (Outlook on the web) to 350 px (Excel), limiting horizontal navigation space. For detailed task pane dimensions and variants, see [Task panes in Office Add-ins](task-pane-add-ins.md).
- **Personality menu obstruction**: The Office personality menu can block top-right UI elements. For exact dimensions and positioning, see [Task panes in Office Add-ins](task-pane-add-ins.md#personality-menu).
- **Ribbon integration**: Navigation must coordinate between ribbon commands and task pane state. See [Add-in commands](add-in-commands.md) for ribbon integration details.
- **Multiplatform variations**: Navigation UI might need to adapt for Web, Windows, Mac, and mobile platforms.
- **Side-by-side context**: Users focus on their document, not your add-in, so navigation must be immediately clear.
- **Multiple containers**: Add-ins might use task panes, dialogs, and content add-ins with different navigation models. See [Office UI elements](interface-elements.md) for an overview.

## Best practices

| Do | Don't |
| :---- | :---- |
| Design navigation to fit within the narrow task pane width. Use vertical navigation or collapsible patterns for space efficiency. | Don't use horizontal navigation patterns designed for wide screens. Avoid wide tab bars that require scrolling or truncate labels. |
| Ensure the user has a clearly visible navigation option, typically at the top of the task pane below the add-in name. | Don't complicate the navigation process by using nonstandard UI or hiding primary navigation controls. |
| Coordinate navigation state between ribbon commands and task pane UI. Deep link to specific views when users select ribbon buttons. | Don't make it difficult for the user to understand their current place or context within the add-in. |
| Keep navigation persistent and visible. Avoid patterns that hide navigation behind multiple taps or clicks. | Don't use navigation patterns that require excessive mouse movement or make it hard to return to frequently used features. |
| Use back buttons for linear workflows, but provide a way to exit multistep processes and return to the main menu. | Don't trap users in deep navigation hierarchies with no clear way to return to the top level. |
| Test navigation on the narrowest platform your add-in supports (often Outlook on the web at 320 px). | Don't design only for the widest task pane and assume it works everywhere. |
| Leverage contextual tabs on the ribbon for context-specific features rather than overloading the task pane with conditional navigation. | Don't show or hide large portions of task pane navigation based on document state. Use the ribbon for contextual commands instead. |

## Ribbon and task pane coordination

Office Add-ins have two primary UI surfaces: the Office ribbon (via [add-in commands](add-in-commands.md)) and the task pane. Plan how these surfaces work together.

### When to use ribbon commands vs. task pane navigation

Use ribbon commands for:

- **Primary entry points** - Major features users access directly from Office (Insert, Format, Analyze).
- **Context-specific actions** - Commands that act on selected content (Format selection, Insert chart).
- **Quick actions** - Single-click operations that don't require extra input.
- **Feature discovery** - Making key capabilities visible without opening the task pane.

Use task pane navigation for:

- **Multistep workflows** - Processes requiring multiple inputs or decisions.
- **Settings and configuration** - Options that don't fit in a single dialog.
- **Content browsing** - Galleries, lists, or catalogs users explore.
- **Persistent state** - Features users keep open while working in their document.

### Coordinating ribbon and task pane state

When a user clicks a ribbon command that opens the task pane, deep link to the relevant view. Use `Office.addin.showAsTaskpane()` to programmatically show the task pane, and then navigate to the appropriate view based on which ribbon command was clicked.

```javascript
// In your ribbon command handler, set the navigation target before showing task pane.
localStorage.setItem('navigationTarget', 'settings');
Office.addin.showAsTaskpane();

// In your task pane startup, navigate to the target view.
const target = localStorage.getItem('navigationTarget');
if (target) {
  navigateToView(target);
  localStorage.removeItem('navigationTarget');
}
```

If you have multiple ribbon commands, each should open the task pane to a specific, relevant view. Avoid having different commands open the same view or the default home screen.

> [!NOTE]
> For complete details on programmatically showing and hiding task panes, see [Show or hide the task pane of your Office Add-in](../develop/show-hide-add-in.md).

### Using contextual tabs with task pane navigation

For Excel add-ins, [contextual tabs](contextual-tabs.md) provide ribbon commands that appear only when specific conditions are met (for example, when a chart is selected). Use contextual tabs to:

- Reduce ribbon clutter by showing commands only when relevant.
- Provide quick actions on selected content without changing the task pane state.
- Coordinate with the task pane - contextual tab commands can trigger task pane navigation changes.

**Example:** A data visualization add-in might show a contextual tab when a chart is selected, with a "Format Chart" button that opens the task pane to chart formatting options.

## Task pane navigation patterns

The following patterns work within the constrained space of an Office task pane.

### Command bar

The [CommandBar](https://react.fluentui.dev/?path=/docs/components-commandbar--default) is a surface at the top of the task pane that houses commands and navigation controls. It typically appears immediately below the add-in name.

:::image type="content" source="../images/add-in-command-bar.png" alt-text="Illustration showing a command bar within an Office desktop application task pane. This example shows a command bar immediately below the add-in name that includes a hamburger menu and search.":::

Use the command bar when:

- Your add-in has four or more top-level sections.
- You need to save horizontal space in narrow task panes.
- You want to include search or filter functionality.

Adding a hamburger menu to the command bar works well in narrow task panes. It places navigation in a collapsible menu to maximize content space.

Use recognizable icons with tooltips to make icon-only commands that save space. Make 44x44 px touch targets so your add-in is more accessible.

> [!IMPORTANT]
> Position right-aligned commands at least 40 px from the top-right corner to avoid the Office personality menu.

### Tab bar

The tab bar navigates with buttons that use vertically stacked text and icons. Tabs appear at the top of the task pane and switch between major sections.

:::image type="content" source="../images/add-in-tab-bar.png" alt-text="Illustration showing a tab bar within an Office desktop application task pane. This example shows a tab bar immediately below the add-in name with 'Home', 'Settings', 'Favorites', and 'Account' tabs.":::

Use a tab bar when:

- Your add-in has 2-4 primary sections of equal importance.
- Users frequently switch between sections.
- Each section has distinct, non-overlapping content.

Because task panes have a narrow width, design tab bars with the following considerations.

- **Width**: Limit to 3-4 tabs maximum in task panes.
  - 2 tabs: ~160-185px each.
  - 3 tabs: ~106-123px each.
  - 4 tabs: ~80-92px each.
- **Truncation risk** - Use short tab labels (one or two words, max twelve characters) to prevent truncation.
- **Icon + text** - Stack the icon above text to reduce width. Use 20x20px icons.
- **Responsive behavior** - On very narrow views (mobile web), consider switching to a command bar.

To retain the Office look and feel, consider using [Fluent UI React Pivot](https://react.fluentui.dev/?path=/docs/components-pivot--default) components. For accessibility, make sure the active tab is visually distinct and use `aria-current="page"` on the active tab.

### Back Button

The back button allows users to navigate backward through a linear workflow or return from detail views to list views.

:::image type="content" source="../images/add-in-back-button.png" alt-text="Illustration showing a back button within an Office desktop application task pane. This example shows a back button immediately below the add-in name, in the top left.":::

Use a back button when:

- You want drill-down navigation, such as when the user selects an item from a list to see details.
- You have multistep workflows, such as wizard-like flows with sequential steps (such as Setup → Configure → Confirm).
- You have modal detail views, such as temporary views that users exit to return to main content.

Position the back button in the top-left corner, immediately below the add-in name, within the title/header area. Make it at least 32x32px, but 44x44px is recommended for better touch accessibility. In narrow task panes, show the current location as text next to the back button instead of a full breadcrumb trail. For example, `[← Back] Display Options` (minimal) or `[← Back] Settings > Display Options` (if space allows). Consider including a home or close icon to allow users to exit multistep flows entirely.

Don't use a back button for tab-based navigation (tabs should be persistent), undo operations (use explicit Undo commands), or browser-like history (add-ins don't have page-based navigation).

To implement a back button, use browser `history.pushState()`. This enables browser back button functionality. Include keyboard shortcuts (Alt+Left Arrow or Escape) for accessibility. Make the back button visually distinct from other navigation elements.

> [!NOTE]
> Breadcrumbs aren't recommended for most add-ins. They consume valuable vertical space and often truncate in narrow task panes.

### Vertical navigation (nav bar)

A persistent vertical list of navigation links, typically on the left side of the task pane.

Use vertical navigation when:

- Your add-in has five or more top-level sections.
- Users need to see all navigation options at once.
- You want to emphasize a hierarchy (parent/child navigation items).

Reserve 48 to 64 px for icon-only navigation, or 120 to 160 px for icon and label navigation. This leaves 160 to 250 px for content in the 320 to 350 px task pane, so using icon-only navigation with tooltips is recommended to maximize content space. Consider a collapsed (icon-only) state by default with an option to expand. Keep navigation fixed while content scrolls, or include navigation in the scroll container if you have many items. Avoid more than one level of nesting due to space constraints.

Consider using [Fluent UI React Nav](https://react.fluentui.dev/?path=/docs/components-nav--default) components to retain the Office look and feel. Highlight the active or selected navigation item and support keyboard navigation (Tab, Arrow keys) for accessibility.

## Multi-step workflows and wizards

For linear, multistep processes (onboarding, setup, configuration), use these patterns:

### Stepped progress indicator

Show users where they are in a multistep process with a visual progress indicator. This pattern is useful for onboarding, setup wizards, and configuration flows.

Use a simple horizontal progress bar (20-40 px height) to save vertical space. Place **Previous** and **Next** buttons in a fixed footer so they're always accessible.

Limit wizards to 3-5 steps maximum. Allow users to skip optional steps and provide a **Cancel** or **Exit** option to return to the main add-in. Save progress automatically so users can resume if they close the task pane.

### Accordion navigation (collapsible sections)

To conserve space in the task pane, use expandable and collapsible sections that show content in place when clicked.

Use accordion navigation when:

- You have five or more content sections that don't need to be visible simultaneously.
- Content within sections varies in length.
- Users typically work with one section at a time.

In narrow panes, allow only one section to be open at a time to avoid excessive scrolling. Use clear expand and collapse icons (▼/▶ or +/−) and remember which sections were expanded when the task pane reopens.

Don't use accordion navigation for primary navigation (use tabs or vertical navigation instead) or frequently toggled content (which causes excessive clicking).

## Navigation state management

### Preserving state across sessions

Users might close and reopen the task pane multiple times during a work session. Preserve navigation state to reduce friction by saving the current view, selected tab, or navigation context.

Use `Office.context.document.settings` to store navigation state that should persist with the document, or use `localStorage` for user-specific preferences that apply across all documents.

```javascript
// Save navigation state.
function navigateTo(view) {
  // Save the current view in document settings to persist with the document.
  Office.context.document.settings.set('currentView', view);
  Office.context.document.settings.saveAsync();

  // You can also use localStorage for user-specific state.
  localStorage.setItem('lastView', view);
}

// Restore the view when the task pane reopens.
Office.onReady(() => {
  const savedView = Office.context.document.settings.get('currentView')
                    || localStorage.getItem('lastView')
                    || 'home';
  navigateTo(savedView);
});
```

> [!NOTE]
> To learn more about persisting add-in state and settings, including best practices for using `Office.context.document.settings`, `localStorage`, and storage partitioning, see [Persist add-in state and settings](../develop/persisting-add-in-state-and-settings.md).

### Deep linking from ribbon commands

Enable ribbon buttons to open the task pane to specific views. This pattern is especially important when you have multiple ribbon commands that should each open different sections of your task pane.

```javascript
// Call this function from your ribbon command handler (function command).
function openSettings(event) {
  localStorage.setItem('openToView', 'settings');
  Office.addin.showAsTaskpane();
  event.completed();
}

// In your task pane startup, check for the target view and navigate accordingly.
Office.onReady(() => {
  const targetView = localStorage.getItem('openToView');
  if (targetView) {
    navigateTo(targetView);
    localStorage.removeItem('openToView');
  }
});
```

For more information about creating function commands in your manifest, see [Create add-in commands](../develop/create-addin-commands.md).

## Multi-container navigation flows

Office Add-ins can use multiple UI containers. Plan ahead to ensure smooth navigation for the user.

### Task pane to dialog navigation

Dialogs provide focused, modal experiences that overlay the Office document. Use dialogs for:

- Authentication and sign-in flows.
- Critical confirmations.
- Forms that require the user's full attention.
- Content that needs more space than the task pane provides.

When you open a dialog from a task pane, use `Office.context.ui.displayDialogAsync()`. Use message passing to communicate between the dialog and task pane.

Some best practices for task pane to dialog navigation include:

- **Clear triggers**: Use explicit buttons, such as "Sign In", to open dialogs.
- **Return path**: Always provide a way to close the dialog and return to the task pane.
- **State synchronization**: Update task pane UI when the dialog closes.
- **Error handling**: Handle when the dialog closes via X button or Escape key.

For complete details on using dialogs, including code examples and API reference, see [Use the Office dialog API in Office Add-ins](../develop/dialog-api-in-office-add-ins.md).

## Content add-in navigation

[Content add-ins](content-add-ins.md) are embedded directly in Excel, PowerPoint, or Word documents and have different navigation constraints than task panes. All UI is within the embedded content boundary, and users can resize the content add-in. Use simple navigation patterns such as single-page layouts or minimal tabs. For complex content add-ins, consider using a companion task pane for settings.

## Platform-specific navigation considerations

### Office on the web

Office on the web has the narrowest task panes. Test your navigation at 320 px width (Outlook on the web) to ensure it works on the most constrained platform. Users might click the browser back button, so use `history.pushState()` to handle this gracefully. Consider responsive breakpoints for phones and tablets accessing Office on the web.

### Office on Windows and Mac

Office on Windows and Mac supports slightly wider task panes, up to 350 px. Test keyboard navigation by using Tab, arrow keys, and standard shortcuts (such as Ctrl+F for search) to make sure your navigation is accessible. Reserve the top-right 40x40 px space to avoid the personality menu.

### Outlook mobile (iOS/Android)

Outlook mobile has even more constrained space than Office on the web. Use single-level navigation or a simple command bar and prioritize features ruthlessly. Make all tappable elements at least 44x44 px for touch accessibility.

### Platform detection and adaptive navigation

Detect the platform and adapt your navigation accordingly. Use `Office.context.platform` to determine the current platform and adjust your navigation layout.

```javascript
// Detect platform and adapt navigation
if (Office.context.platform === Office.PlatformType.OfficeOnline) {
  // Use more compact navigation for web
  useCompactNavigation();
} else if (Office.context.platform === Office.PlatformType.iOS ||
           Office.context.platform === Office.PlatformType.Android) {
  // Use touch-optimized navigation
  useMobileNavigation();
}
```

## Accessibility in navigation

All navigation must be accessible by using a keyboard with Tab, arrow keys, and Enter. Provide a clear visual indication of which element has focus. Use proper ARIA labels, including `role="navigation"`, `aria-label`, and `aria-current="page"` for screen reader support. Provide a "Skip to content" link for screen reader users and ensure navigation appears early in the tab order, before main content.

The following example shows accessible navigation markup.

```html
<nav role="navigation" aria-label="Main navigation">
  <a href="#main-content" class="skip-link">Skip to content</a>
  <ul>
    <li><a href="#home" aria-current="page">Home</a></li>
    <li><a href="#settings">Settings</a></li>
  </ul>
</nav>
```

## Decision tree: Choosing a navigation pattern

Use this guide to select the appropriate navigation pattern for your add-in.

1. **How many top-level sections does your add-in have?**
   - 1-2 sections: No persistent navigation needed; use ribbon commands and simple content.
   - 2-4 sections: **Tab Bar** (if equal importance) or **Command Bar with menu** (if hierarchical).
   - 4-7 sections: **Command Bar with hamburger menu** or **Vertical Nav** (icon-only).
   - 8+ sections: Reconsider information architecture and consolidate.

1. **Do users frequently switch between sections?**
   - Yes: **Tab Bar** (keep all options visible).
   - No: **Command Bar with menu** (save space).

1. **Does your add-in have multistep workflows?**
   - Yes: Use **Back Button** and **Progress Indicator** for linear flows.
   - No: Use direct navigation (tabs or menu).

1. **Does your add-in have more than one point of entry from the ribbon?**
   - Yes: Plan **deep linking** from ribbon to specific task pane views.
   - No: Task pane can be independent of the ribbon.
   - Consider **Contextual Tabs** (Excel only) for context-specific commands.

## See also

- [Task panes in Office Add-ins](task-pane-add-ins.md)
- [Add-in commands](add-in-commands.md)
- [Office UI elements](interface-elements.md)
- [Create custom contextual tabs](contextual-tabs.md)
- [UX design pattern templates](ux-design-pattern-templates.md)
- [Fluent UI React components](../quickstarts/fluent-react-quickstart.md)
- [Persist add-in state and settings](../develop/persisting-add-in-state-and-settings.md)
- [Use the Office dialog API](../develop/dialog-api-in-office-add-ins.md)
