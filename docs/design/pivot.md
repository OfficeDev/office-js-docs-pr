# Pivot component in Office UI Fabric

In Add-ins, Pivot control, “Tab pattern” are used for quick navigation to frequently accessed, distinct content categories. It allows for navigation between two or more content views and relies on text headers to articulate the different sections of content. Tabs are a visual variant of Pivot that use a combination of icons and text or just icons to articulate section content.

#### Example: Pivot in a task pane

![An image showing the Pivot](../../images/overview_withApp_pivot.png)

## Best Practices

|**Do**|**Don't**|
|:------------|:--------------|
|Navigation labels should be concise, ideally using only one or two words rather than a phrase.|Don’t use full sentences or complex punctuation, such as colons or semicolons.|
|Persist pivot headers on-screen even if another tab is selected.| |
|Limit pivot controls to 3-5 tabs.| |
|Use pivots as navigational elements close to the top of the page. Don't mix pivots into page content.| |
|Use pivots on content-heavy pages that require a significant amount of scrolling.| |

## Variants

|**Variation**|**Description**|**Example**|
|:------------|:--------------|:----------|
|**Basic Example**|Use as the default pivot option.|![Basic Example image](../../images/pivotBasic.png)|
|**Links of Tab Style**|Use when tab style pivot buttons are preferred.|![Links of Tab Style image](../../images/pivotTab.png)|

## Implementation

For details, see [Pivot](https://dev.office.com/fabric#/components/pivot) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## Additional resources

* [UX design patterns](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
* [Office UI Fabric in Office Add-ins](office-ui-fabric.md)
