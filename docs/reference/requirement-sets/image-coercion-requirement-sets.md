---
title: Image Coercion requirement sets
description: 'Support for Image Coercion requirement sets with Office Add-ins across Excel, PowerPoint, and Word.'
ms.date: 02/19/2021
ms.prod: non-product-specific
localization_priority: Normal
---

# Image Coercion requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

## ImageCoercion 1.1

ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method. The following applications are supported.

- Excel 2013 and later on Windows
- Excel 2016 and later on Mac
- Excel on iPad
- OneNote on the web
- PowerPoint 2013 and later on Windows
- PowerPoint 2016 and later on Mac
- PowerPoint on the web
- PowerPoint on iPad
- Word 2013 and later on Windows
- Word 2016 and later on Mac
- Word on the web
- Word on iPad

## ImageCoercion 1.2

ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method. The following applications are supported.

- Excel on Windows (connected to a Microsoft 365 subscription)
- Excel on Mac (connected to a Microsoft 365 subscription)
- PowerPoint on Windows (connected to a Microsoft 365 subscription)
- PowerPoint on Mac (connected to a Microsoft 365 subscription)
- PowerPoint on the web
- Word on Windows (connected to a Microsoft 365 subscription)
- Word on Mac (connected to a Microsoft 365 subscription)

## Office Common API requirement sets

For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).

## See also

- [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md)
- [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../develop/add-in-manifests.md)
