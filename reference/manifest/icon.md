# Icon element
Defines **Image** elements for [Button](./button.md) and [Menu](./menu-control.md) controls.

## Child elements
|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Image](#image)        | Yes |   resid of an image to use         |

## Image
An image for the button. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element in the **Images** element in the [Resources](./resources.md) element. The **size** attribute indicates the size in pixels of the image. Three image sizes are required (16, 32, and 80 pixels) while five other sizes are supported (20, 24, 40, 48, and 64 pixels).|


```xml
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
```  