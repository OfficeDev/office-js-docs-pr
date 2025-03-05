---
title: Use fields in your Word add-in
description: Learn to use fields in your Word add-in.
ms.date: 03/12/2025
ms.localizationpriority: medium
---

# Use fields in your Word add-in

A [field](https://support.microsoft.com/office/c429bbb0-8669-48a7-bd24-bab6ba6b06bb) in a Word document is a placeholder. It allows you to provide instructions for the content instead of the content itself. You can use fields to create and format a Word template. Word documents support a number of [field types](https://support.microsoft.com/office/1ad6d91a-55a7-4a8d-b535-cf7888659a51), many with associated parameters for configuring the field. However, Word on the web doesn't support adding or editing field through the UI. For more information, see [Field codes in Word for the web](https://support.microsoft.com/office/d8f46094-13c3-4966-98c3-259748f3caf1).

Word JavaScript APIs - TODO: more about this

## Addin field

The Addin field is meant to store add-in data that's hidden from the Word user interface, regardless of whether fields in the document are set to show or hide its content. The Addin field isn't available in the Word UI's **Field** dialog box. Use the API to insert the Addin field type and set the field's data. Unlike other fields, Word on the web does allow you to update the Addin field.

TODO: Include code example and link to Script Lab snippet

## Date field

TODO: Include code example and link to Script Lab snippet

[Field codes: Date field](https://support.microsoft.com/office/d0c7e1f1-a66a-4b02-a3f4-1a1c56891306)

## Hyperlink field

TODO: Include code example and link to Script Lab snippet

[Field codes: Hyperlink field](https://support.microsoft.com/office/864f8577-eb2a-4e55-8c90-40631748ef53)

## TOC (Table of Contents) field

TODO: Include code example and link to Script Lab snippet

[Field codes: TOC (Table of Contents) field](https://support.microsoft.com/office/1f538bc4-60e6-4854-9f64-67754d78d05c)

## See also

- [Field codes in Word for the web](https://support.microsoft.com/office/d8f46094-13c3-4966-98c3-259748f3caf1)
- [Insert, edit, and view fields in Word](https://support.microsoft.com/office/c429bbb0-8669-48a7-bd24-bab6ba6b06bb)
