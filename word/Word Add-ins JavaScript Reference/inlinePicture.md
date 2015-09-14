# InlinePicture

Represents an inline picture anchored to a paragraph.

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|\
|parentContentControl|  [ContentControl](contentControl.md)   |Gets the content control that contains the body. Returns null if there isn't a parent content control.|
|altTextDescription| string  | Gets or sets a string that represents the alternative text associated with the inline image  |
|altTextTitle| string  | Gets or sets a string that contains the title for the inline image. |
|height| number  |  Gets or sets a number that describes the height of an inline image. You cannot set this value if lockAspectRatio is set to true. |
|hyperlink| string  | Gets or sets the hyperlink associated with the inline shape. |
|lockAspectRatio| bool  | Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it. True if you cannot change the height and width of the shape independently of one another when you resize it, otherwise, false. |
|width| number  | Gets or sets a number that describes the width of an inline image. You cannot set this value if lockAspectRatio is set to true. |

## Methods


| Method     | Return Type    |Description|
|:-----------------|:--------|:----------|
|[getBase64ImageSrc()](#getbase64imagesrc)| string | Gets the base64 encoded string representation of the inline image. | 
|[insertContentControl()](#insertcontentcontrol)| [ContentControl](contentcontrol.md)  |Wraps the calling inline picture with a Rich Text content control. |  
|[load(param: option)](#loadparam-option)| void | Fills the inline picture proxy object created in the JavaScript layer with property and object values specified in the parameter.|

## API Specification

### getBase64ImageSrc()
Gets the base64 encoded string representation of the inline image.

#### Syntax
```js
    inlinePicture.getBase64ImageSrc();
```
#### Parameters

None

#### Returns

string


#### Example

```js
    //gets all the images in the body of the document and then gets the base64 for each.
    var ctx = new Word.RequestContext();


    var pics = ctx.document.body.inlinePictures;
    ctx.load(pics);
    ctx.references.add(pics);

    ctx.executeAsync().then(
        function () {
            var results = new Array();

            for (var i = 0; i < pics.items.length; i++) {
                results.push(pics.items[i].getBase64ImageSrc());
            }
            ctx.executeAsync().then(
                function () {
                    for (var i = 0; i < results.length; i++) {
                        console.log("pics[" + i + "].base64 = " + results[i].value);
                    }
                }
            );
        },
        function (result) {
            console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
            console.log(result.traceMessages);
        }
    );

```
[Back](#methods)


### insertContentControl()

Wraps the calling inline picture with a Rich Text content control.

#### Syntax
```js
    inlinePicture.insertContentControl();
```
#### Parameters

None

#### Returns

[ContentControl](contentControl.md).

#### Example

```js
    // grabs the first paragraph in the document and inserts an image at the end of it, then sets a
    // few properties, then wraps it inside a content control.
    var ctx = new Word.RequestContext();
    var paras = ctx.document.body.paragraphs;
    ctx.load(paras);

    var myImage = paras.getItem(0).insertInlinePictureFromBase64("iVBORw0KGgoAAAANSUhEUgAAAIAAAACABAMAAAAxEHz4AAAAJFBMVEX///9GRkZGRkZGRkZGRkZGRkZGRkZGRkYBpO9/ugDyUCL/uQGm4PjWAAAACHRSTlMBCQ0RFRknMx7uViEAAAB3SURBVGje7dcxCYBQGEXhi6izYBHB0RIiiAXkzW5iAMEKFnCwguVscJd/ecM5Ab79SNHK5FqlZXeNql/XIx23awMAAAAAAAAAAAAAAAAAyBwIvzNJxeyapLZ3Naou1ykNn6sDAAAAAAAAAAAAAAAAAMgcCL9ztB/UhshWs1l/WAAAAABJRU5ErkJggg==", Word.InsertLocation.end);


    myImage.width = 100;
    myImage.height = 100;
    myImage.lockAspectRatio = true;
    myImage.hyperlink = "http://dev.office.com";
    var myCC = myImage.insertContentControl();
    myCC.title = "My Image";
    myCC.appearance = "tags";

    ctx.references.add(myImage);

    ctx.executeAsync().then(
        function () {
            console.log("*" + myImage.id);
            console.log("Success");
        },
        function (result) {
            console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
            console.log(result.traceMessages);
        }
    );
```
[Back](#methods)

### load(param: option)
Fills the inlinePicture proxy object created in the JavaScript layer with the property and object values specified in the parameter.

#### Syntax
```js
    document.body.load(param);
```

#### Parameters
| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|param|object| A string, a string with comma separated value, an array of strings, or an object that specifies which properties to load.  |

#### Returns
void

[Back](#methods)
