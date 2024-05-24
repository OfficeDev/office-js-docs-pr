---
title: Publish your Office Add-in to Microsoft AppSource
description: Learn how to publish your Office Add-in to Microsoft AppSource 
ms.topic: concept-article
ms.date: 05/24/2024
CustomerIntent: As a developer, I want to publish my Office Add-in to Microsoft AppSource so that customers can deploy and use my new add-in.
---

# Publish your Office Add-in to Microsoft AppSource

Publish your Office Add-in to Microsoft AppSource to make it widely available to customers and businesses. Microsoft AppSource is an online store that contains thousands of business applications and services built by industry-leading software providers. When you publish your add-in to Microsoft AppSource, you also make it available in the in-product experience within Office.

Microsoft AppSource provides a convenient location for you to upload new Office and SharePoint Add-ins, Microsoft Teams apps, and Power BI visuals that provide solutions for both consumers and businesses. When you add your app solution to Microsoft AppSource, you also make it available in the in-product experience within Office. To include your solution in Microsoft AppSource and within Office, you submit it to Partner Center. You need to create a company account and, if applicable, add payout information. For details, see the following articles:

## The publishing process

To include your solution in Microsoft AppSource and within Office, you submit it to Partner Center. Then it goes through an approval and certification process. For complete details, see [Make your solutions available in Microsoft AppSource and within Office](/partner-center/marketplace/submit-to-appsource-via-partner-center).

## Provide an install link

To help customers discover and install your add-in, create an install link that provides a "click and run" experience. Put the link on your website, social media, or anywhere you think helps your customers discover your add-in.

The link opens a new Word, Excel, or PowerPoint document in the browser for the signed-in user. Your add-in is automatically loaded in the new document so you can guide users to try your add-in without the need to search for it in Microsoft AppSource and install it manually.

To create the link, use the following URL template as a reference.

`https://go.microsoft.com/fwlink/?linkid={{linkId}}&templateid={{addInId}}&templatetitle={{addInName}}`

Change the three parameters in the previous URL to support your add-in as follows.

- **linkId**: Specifies which web endpoint to use when opening the new document.

  - For Word on the web: `2261098`
  - For Excel on the web: `2261819`
  - For PowerPoint on the web: `2261820`

  **Note:** Outlook is not supported at this time.

- **templateid**:  The ID of your add-in as listed in Microsoft AppSource.
- **templatetitle**:  The full title of your add-in. This must be HTML encoded.

For example, if you want to provide an install link for [Script Lab](https://appsource.microsoft.com/product/office/wa104380862), use the following link.

[https://go.microsoft.com/fwlink/?linkid=2261819&templateid=WA104380862&templatetitle=Script%20Lab,%20a%20Microsoft%20Garage%20project](https://go.microsoft.com/fwlink/?linkid=2261819&templateid=WA104380862&templatetitle=Script%20Lab,%20a%20Microsoft%20Garage%20project)

The following parameter values are used for the Script Lab install link.

- **linkid:**  The value `2261819` specifies the Excel endpoint. Script Lab supports Word, Excel and PowerPoint, so this value can be changed to support different endpoints.
- **templateid:** The value `WA104380862` is the Microsoft AppSource id for Script Lab.
- **templatetitle:** The value `Script%20Lab,%20a%20Microsoft%20Garage%20project` which is the HTML encoded value of the title.

## Related content

- [Make your solutions available in Microsoft AppSource and within Office](/partner-center/marketplace/submit-to-appsource-via-partner-center)
- [What is Microsoft AppSource?](/marketplace/appsource-overview)
