
# AllowSnapshot element
Specifies whether a snapshot image of your content add-in is saved with the host document.

 **Add-in type:** Content


## Syntax:


```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```


## Contained in:

[OfficeApp](../../reference/manifest/officeapp.md)


## Remarks


 **Security Note**   **AllowSnapshot** is **true** by default. This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in. However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.

