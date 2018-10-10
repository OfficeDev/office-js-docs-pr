# Identity API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Office Add-ins run across multiple versions of Office. The following table lists the Identity API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.

|  Requirement set  | Office 2013 for Windows | Office 365 for Windows   |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com & Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | N/A | Preview**&#42;** | Coming soon | Preview**&#42;**| Preview | Preview| Coming soon | Coming soon |

> **&#42;** During the preview phase, the Identity API is supported on Windows 2016 and Mac only for users in the Insiders program using the Fast option. To join the Insiders program, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1). To switch to the Fast track, see [Insider Fast](https://answers.microsoft.com/en-us/msoffice/forum/msoffice_officeinsider-mso_win10-msoinsider_reg/its-here-office-insider-fast-for-office-2016-on/dbe8e7bb-9523-44a4-948b-9436fedfd961).

To find out more about versions, build numbers, and Office Online Server, see:

- [Version and build numbers of update channel releases for Office 365 clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [What version of Office am I using?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Where you can find the version and build number for an Office 365 client application](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server overview](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## Office common API requirement sets

For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).

## IdentityAPI 1.1 

The Single Sign On IdentityAPI 1.1 is the first version of the API. For details about this API, see the [SSO API reference](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) section of [Enable SSO in an add-in](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins).

## See also

- [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Specify Office hosts and API requirements](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office Add-ins XML manifest](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
