# Enumerations (JavaScript API for Word)

## Word.ContentControlType
* **RichText**: Identifies a rich text content control.
* **Unknown**
* **RichTextInline**
* **RichTextParagraphs**
* **RichTextTableCell**: contains whole cell
* **RichTextTableRow**: contains whole row
* **RichTextTable**: contains whole table
* **PlainTextInline**
* **PlainTextParagraph**
* **Picture**
* **BuildingBlockGallery**
* **CheckBox**
* **ComboBox**
* **DropDownList**
* **DatePicker**
* **RepeatingSection**
* **PlainText**

## Word.ContentControlAppearance
* **BoundingBox**: Represents a content control shown as a shaded rectangle or bounding box (with optional title).
* **Tags**: Represents a content control shown as start and end markers.
* **Hidden**: Represents a content control that is not shown.

## Word.UnderlineType
* **None**: No underline.
* **Single**: A single underline. This is the default value.
* **Word**: Only underline individual words.
* **Double**: A double underline.
* **Dotted**: A dotted underline.
* **Hidden**: A hidden underline.
* **Thick**: A single thick underline.
* **DashLine**: A single dash underline.
* **DotLine**: A single dot underline.
* **DotDashLine**: An alternating dot-dash underline.
* **TwoDotDashLine**: An alternating dot-dot-dash underline.
* **Wave**: A single wavy underline.

## Word.BreakType
* **Page**: Page break at the insertion point.
* **Column**: Column break at the insertion point.
* **Next**: Section break on next page.
* **SectionContinuous**: New section without a corresponding page break.
* **SectionEven**: Section break with the next section beginning on the next even-numbered page. If the section break falls on an even-numbered page, Word leaves the next odd-numbered page blank.
* **SectionOdd**: Section break with the next section beginning on the next odd-numbered page. If the section break falls on an odd-numbered page, Word leaves the next even-numbered page blank.
* **Line**: Line break.
* **LineClearLeft**: Line break.
* **LineClearRight**: Line break.
* **TextWrapping**: Ends the current line and forces the text to continue below a picture, table, or other item. The text continues on the next blank line that does not contain a table aligned with the left or right margin.

## Word.InsertLocation
* **Before**: Add content before the contents of the calling object.
* **After**: Add content after the contents of the calling object.
* **Start**: Prepend content to the contents of the calling object.
* **End**: Append content to the contents of the calling object.
* **Replace**: Replace the contents of the current object.

## Word.Alignment
* **Unknown**: Unknown alignment.
* **Left**: Alignment to the left.
* **Centered**: Alignment to the center.
* **Right**: Alignment to the right.
* **Justified**: Fully justified alignment.

## Word.HeaderFooterType
* **Primary**: Returns the header or footer on all pages of a section, with the first page or odd pages excluded if they are different.
* **FirstPage**: Returns the header or footer on the first page of a section.
* **EvenPages**: Returns all headers or footers on even-numbered pages of a section.

## Word.ErrorCodes
* **AccessDenied**: The client doesn't have access to the requested resource. Check your permissions.
* **GeneralException**: A general error occurred.
* **InvalidArgument**: One or more of the arguments are not supported on this context. Check the documentation.
* **ItemNotFound**: The resource was not found.
* **NotImplemented**: The operation is not implemented.

## Word.BodyType
* **Unknown**
* **MainDoc**
* **Section**
* **Header**
* **Footer**
* **TableCell**


## Word.SelectionMode
* **Select**
* **Start**
* **End**


## Word.ImageFormat
* **Unsupported**
* **Undefined**
* **Bmp**
* **Jpeg**
* **Gif**
* **Tiff**
* **Png**
* **Icon**
* **Exif**
* **Wmf**
* **Emf**
* **Pict**
* **Pdf**
* **Svg**


## Word.RangeLocation
* **Whole**
* **Start**
* **End**
* **Before**
* **After**
* **Content**

## Word.ListLevelType
* **Bullet**
* **Number**
* **Picture**


## Word.LocationRelation
* **Unrelated**: this instance and the range are in different sub-documents
* **Equal**: this instance and the range represent the same range
* **ContainsStart**: this instance contains the range and that it shares the same start character. The range does not share the same end character as this instance
* **ContainsEnd**: this  instance contains the range and that it shares the same end character. The range does not share the same start character as this instance
* **Contains**: this instance contains the range, with the exception of the start and end character of this instance
* **InsideStart**: this  instance is inside the range and that it shares the same start character. The range does not share the same end character as this instance
* **InsideEnd**: this instance is inside the range and that it shares the same end character. The range does not share the same start character as this instance
* **Inside**: this instance is inside the range. The range does not share the same start and end characters as this instance
* **AdjacentBefore**: this instance occurs before, and is adjacent to, the range
* **OverlapsBefore**: this instance starts before the range and overlaps the range’s first character
* **Before**: this instance occurs before the range
* **AdjacentAfter**: this instance occurs after, and is adjacent to, the range
* **OverlapsAfter**: this instance starts inside the range and overlaps the range’s last character
* **After**: this instance occurs after the range


## Word.BorderLocation
* **Top**
* **Left**
* **Bottom**
* **Right**
* **InsideHorizontal**
* **InsideVertical**
* **Inside**
* **Outside**
* **All**


## Word.BorderType
* **Mixed**
* **None**
* **Single**
* **Thick**
* **Double**
* **Hairline**
* **Dotted**
* **Dashed**
* **DotDashed**
* **Dot2Dashed**
* **Triple**
* **ThinThickSmall**
* **ThickThinSmall**
* **ThinThickThinSmall**
* **ThinThickMed**
* **ThickThinMed**
* **ThinThickThinMed**
* **ThinThickLarge**
* **ThickThinLarge**
* **ThinThickThinLarge**
* **Wave**
* **DoubleWave**
* **DashedSmall**
* **DashDotStroked**
* **ThreeDEmboss**
* **ThreeDEngrave**


## Word.VerticalAlignment
* **Mixed**
* **Top**
* **Center**
* **Bottom**


## Word.ListBullet
* **Custom **
* **Solid**
* **Hollow**
* **Square**
* **Diamonds**
* **Arrow**
* **Checkmark**


## Word.ListNumbering
* **None**
* **Arabic**
* **UpperRoman**
* **LowerRoman**
* **UpperLetter**
* **LowerLetter**


## Word.Style
* **Other**
* **Normal**
* **Heading1**
* **Heading2**
* **Heading3**
* **Heading4**
* **Heading5**
* **Heading6**
* **Heading7**
* **Heading8**
* **Heading9**
* **Toc1**
* **more..**

## Word.DocumentPropertyType
* **String**
* **Number**
* **Date**
* **Boolean**
