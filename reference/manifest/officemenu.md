# OfficeMenu element
Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel and PowerPoint Add-ins.  

## Attributes

| Attribute            | Required | Description                          |
|:---------------------|:--------:|:-------------------------------------|
| [xsi:type](#xsitype) | Yes      | The type of OfficeMenu being defined.|

## Child elements
|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Control](#control)    | Yes |  Collection of one or more Control objects.  |

## xsi:type
Specifies a built-in menu of the Office client application to add this Office add-in.

- `ContextMenuText` -  Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. Applies to Word, Excel and PowerPoint. 
- `ContextMenuCell` -  Displays the item on the context menu when the user right-clicks on a cell on the spreadsheet. Applies to Excel. 

## Control

Each OfficeMenu requires at one or more [menu](./menu.md#menu-control) controls . 


## Example

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="myMenuID">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>   
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>    
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>    
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />    
          </Action>    
        </Item>
      </Items>
    </Control>   
</OfficeMenu>
```
