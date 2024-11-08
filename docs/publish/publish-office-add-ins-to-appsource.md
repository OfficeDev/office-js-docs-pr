---
title: Publish your Office Add-in to Microsoft AppSource
description: Learn how to publish your Office Add-in to Microsoft AppSource and install the add-in with a Windows app or COM/VSTO add-in.
ms.topic: concept-article
ms.date: 10/16/2024
CustomerIntent: As a developer, I want to publish my Office Add-in to Microsoft AppSource so that customers can deploy and use my new add-in.
---

# Publish your Office Add-in to Microsoft AppSource

Publish your Office Add-in to Microsoft AppSource to make it widely available to customers and businesses. Microsoft AppSource is an online store that contains thousands of business applications and services built by industry-leading software providers. When you publish your add-in to Microsoft AppSource, you also make it available in the in-product experience within Office.

## The publishing process

Before you proceed:

- Have a [Partner Center account](/partner-center/marketplace-offers/open-a-developer-account).
- Ensure that your add-in adheres to the applicable [AppSource validation policies](/legal/marketplace/certification-policies).
- Confirm that you're [ready to publish](/partner-center/marketplace-offers/checklist).

When you're ready to include your solution in Microsoft AppSource and within Office, submit it to Partner Center. Then, it goes through an approval and certification process. For complete details, see [Make your solutions available in Microsoft AppSource and within Office](/partner-center/marketplace/submit-to-appsource-via-partner-center).

When your add-in is available in AppSource, there are two further steps you can take to make it more widely installed. 

- [Provide an installation link](#provide-an-installation-link)
- [Include the add-in in the installation of a Windows app or a COM or VSTO add-in](#include-the-add-in-in-the-installation-of-a-windows-app-or-comvsto-add-in)

### Provide an installation link

After you publish to Microsoft AppSource, you can create an installation link to help customers discover and install your add-in. The installation link provides a "click and run" experience. Put the link on your website, social media, or anywhere you think helps your customers discover your add-in.

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

For example, if you want to provide an installation link for [Script Lab](https://appsource.microsoft.com/product/office/wa104380862), use the following link.

[https://go.microsoft.com/fwlink/?linkid=2261819&templateid=WA104380862&templatetitle=Script%20Lab,%20a%20Microsoft%20Garage%20project](https://go.microsoft.com/fwlink/?linkid=2261819&templateid=WA104380862&templatetitle=Script%20Lab,%20a%20Microsoft%20Garage%20project)

The following parameter values are used for the Script Lab installation link.

- **linkid:**  The value `2261819` specifies the Excel endpoint. Script Lab supports Word, Excel, and PowerPoint, so this value can be changed to support different endpoints.
- **templateid:** The value `WA104380862` is the Microsoft AppSource ID for Script Lab.
- **templatetitle:** The value `Script%20Lab,%20a%20Microsoft%20Garage%20project` which is the HTML encoded value of the title.

### Include the add-in in the installation of a Windows app or COM/VSTO add-in

When you have a Windows app or a COM or VSTO add-in whose functions overlap with your Office Web Add-in, consider including the web add-in in the installation (or an upgrade) of the Windows app or COM/VSTO add-in. (This installation option is supported only for Excel, PowerPoint, and Word add-ins.) To do this, include in the installation program a function to add an entry like the following example to the Windows Registry. (The exact code will depend on your installation framework.)

```
[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Wef\AutoInstallAddins\{{OfficeApplication}}\{{add-inName}}] 
"AssetIds"="{{assetId}}"
```

Replace the placeholders as follows:

- `{{OfficeApplication}}` with the name of the Office application that the add-in should be installed in. Only `Word`, `Excel`, and `PowerPoint` are supported.

   > [!NOTE]
   > If the add-in's manifest is configured to support more than one Office application, replace `{{OfficeApplication}}` with any *one* of the supported applications. Don't create separate registry entries for each supported application. The add-in will be installed for all the Office applications that it supports. 

- `{{add-inName}}` with the name of the add-in; for example `ContosoAdd-in`.
- `{{assetId}}` with the AppSource asset ID of your add-in, such as `WA999999999`.

The following is an example.

```
[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Wef\AutoInstallAddins\Word\ContosoAdd-in] 
"AssetIds"="WA999999999"
```

When an end user runs your installation executable, their experience with the web add-in installation will depend on two factors.

- Whether you're a [certified Microsoft 365 developer](/microsoft-365-app-certification/docs/certification). For more information, see [Microsoft 365 App Compliance Program](https://developer.microsoft.com/microsoft-365/app-compliance-program).
- The security settings made by the user's Microsoft 365 administrator.

If you're certified and the administrator has enabled automatic approval for all apps from certified developers, then the web add-in is installed without the need for any special action by the user after the installation executable is started. If you're not certified or the administrator hasn't granted automatic approval for all apps from certified developers, then the user will be prompted to approve inclusion of the web add-in as part of the overall installation. After installation, the web add-in is available to the user in Office on the web as well as Office on Windows.

If you're combining the installation of a web add-in with a COM/VSTO add-in, you need to think about the relationship between the two. For more information, see [Make your Office Add-in compatible with an existing COM add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).

## Related content

- [Make your solutions available in Microsoft AppSource and within Office](/partner-center/marketplace/submit-to-appsource-via-partner-center)
- [What is Microsoft AppSource?](/marketplace/appsource-overview)
