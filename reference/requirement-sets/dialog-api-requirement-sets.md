
# Dialog API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Specify Office hosts and API requirements](../docs/overview/specify-office-hosts-and-api-requirements.md).

Office Add-ins run across multiple versions of Office. The following table lists the Dialog API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.

|  Requirement set  |  Office 2016 for Windows   |  Office 2016 for iPad  |  Office 2016 for Mac  | Office Online  | 
|:-----|-----|:-----|:-----|:-----|
| DialogApi 1.1  |  Version 1602 (Build 6741.0000) or later | 1.22 or later | 15.20 or later| We're working on it. |

> **Note**: The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version does not contain the DialogApi 1.1 requirement set.

To find out more about versions and build numbers, see:

- [Version and build numbers of update channel releases for Office 365 clients](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [What version of Office am I using?](https://support.office.com/en-us/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-US&rs=en-US&ad=US&fromAR=1)
- [Where you can find the version and build number for an Office 365 client application](https://technet.microsoft.com/en-us/library/mt592918.aspx#Anchor_1)

## Office common API requirement sets
For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).

## Dialog API 1.1 
The Dialog API 1.1 is the first version of the API. For details about the API, see the [Dialog API](../shared/officeui.md) reference topics.

## Additional resources

- [Specify Office hosts and API requirements](../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](https://dev.office.com/docs/add-ins/overview/add-in-manifests)
