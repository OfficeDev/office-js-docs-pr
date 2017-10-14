# ChoiceGroup component in Office UI Fabric

The ChoiceGroup component, also known as a radio button, presents users with two or more mutually exclusive options. Users can select only one ChoiceGroup button in a group. Each option is represented by one ChoiceGroup button. 
  
#### Example: ChoiceGroup in a task pane

 ![An image showing a ChoiceGroup](../../images/overview_withApp_choicegroup.png)

<br/>

## Best practices

|**Do**|**Don't**|
|:------------|:--------------|
|Keep ChoiceGroup options at the same level.<br/><br/>![Do ChoiceGroup example](../../images/choiceDo.png)<br/>|Don't use nested ChoiceGroups or check boxes.<br/><br/>![Don't ChoiceGroup example](../../images/choiceDont.png)<br/>|
|Use ChoiceGroups with 2-7 options, ensuring there is enough screen space to show all options. Otherwise, use a check box or drop-down list.|Don't use when the options are numbers with a fixed step, for example 10, 20, 30, and so on. Instead, use a slider component.|
|If users may not choose any of the options, consider including an option such as **None** or **Does not apply**.|Don’t use two ChoiceGroup buttons for a single binary choice.|
|If possible, align ChoiceGroup buttons vertically instead of horizontally. Horizontal alignment is harder to read and localize.||
|List options in logical order, for example, the most likely option to be selected to the least, the simplest operation to the most complex, or the least risk to the highest risk. |Don't use alphabetical ordering because it is language dependent.|

## Variants

|**Variation**|**Description**|**Example**|
|:------------|:--------------|:----------|
|**ChoiceGroups**|Use when imagery is not necessary for making a selection.|![ChoiceGroup variant image](../../images/radio.png)<br/>|
|**ChoiceGroups using images**|Use when imagery is necessary for making a selection.|![ChoiceGroup variant with image](../../images/radioImage.png)<br/>|

## Implementation

For details, see [ChoiceGroup](https://dev.office.com/fabric#/components/choicegroup) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## Additional resources

- [UX Design Patterns](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

- [Office UI Fabric in Office Add-ins](office-ui-fabric.md)
