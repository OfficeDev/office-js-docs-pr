---
title: Pivot component in Office UI Fabric
description: ''
ms.date: 12/04/2017
---


# Pivot component in Office UI Fabric

Pivots provide quick navigation to frequently accessed content. Pivots allow for navigation between two or more content views. Text headers specify which content is in each section of the pivot. Content in each section of the pivot may belong to distinct content categories. In Office Add-ins, use the Pivot control with tab styles. The tabs may use a combination of icons and text to communicate the type of content that the tab contains. 

#### Example: Pivot in a task pane

![An image showing the pivot](../images/overview-with-app-pivot.png)

## Best practices

|**Do**|**Don't**|
|:------------|:--------------|
|Navigation labels should be concise, ideally using only one or two words rather than a phrase.|Don’t use full sentences or complex punctuation, such as colons or semicolons.|
|Persist pivot headers on-screen even if another tab is selected.| |
|Limit pivot controls to 3–5 tabs.| |
|Use pivots as navigational elements close to the top of the page. Don't mix pivots into page content.| |
|Use pivots on content-heavy pages that require a significant amount of scrolling.| |

## Variants

|**Variation**|**Description**|**Example**|
|:------------|:--------------|:----------|
|**Basic example**|Use as the default pivot option.|![Basic Example image](../images/pivot-basic.png)<br/>|
|**Links of tab style**|Use when tab style pivot buttons are preferred.|![Links of Tab Style image](../images/pivot-tab.png)<br/>|

## Implementation

For details, see [Pivot](https://dev.office.com/fabric#/components/pivot) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## Additional resources

- [UX Design Patterns](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Office UI Fabric in Office Add-ins](office-ui-fabric.md)
