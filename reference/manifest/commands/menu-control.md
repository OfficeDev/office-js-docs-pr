# Menu (dropdown button) controls

A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported. 

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Label](#label)     | Yes |  The text for the button.         |
|  [Supertip](./supertip.md)  | Yes |  The supertip for this button.    |
|  [Icon](./icon.md)      | Yes |  An image for the button.         |
|  [Items](#items)     | Yes |  A collection of Buttons to display within the menu. |

## Label
Required. The text for the button. The  **resid** attribute must be set to the value of the **id** attribute 
of a **String** element in the [ShortStrings](./resources.md#shortstrings) element in the [Resources](./resources.md) element.

## Supertip
Required. See [Supertip](./supertip.md).
 
## Icon
Required. See [Icon](./icon.md).

## Items
Required. Contains the  **Item** elements for the menu. Each **Item** element contains the same child elements as a [Button controls](./button-control.md).

## Menu control example
```xml
<Control xsi:type="Menu" id="msgReadMenuButton">
  <Label resid="menuReadButtonLabel" />
  <Supertip>
    <Title resid="menuReadSuperTipTitle" />
    <Description resid="menuReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="red-icon-16" />
    <bt:Image size="32" resid="red-icon-32" />
    <bt:Image size="80" resid="red-icon-80" />
  </Icon>
  <Items>
    <Item id="msgReadMenuItem1">
      <Label resid="menuItem1ReadLabel" />
      <Supertip>
        <Title resid="menuItem1ReadLabel" />
        <Description resid="menuItem1ReadTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getItemClass</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```