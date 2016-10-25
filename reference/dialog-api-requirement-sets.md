# Dialog API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Specify Office hosts and API requirements](../docs/overview/specify-office-hosts-and-api-requirements.md).

Office Add-ins run across multiple versions of Office including Office 2016 for Windows, Office for the iPad and Office for the Mac. The following table lists the Dialog API requirement sets, the Office host applications that support that requirement set, and the build versions.

|  Requirement set  |  Office 2016 for Windows   |  Office 2016 for iPad  |  Office 2016 for Mac  | Office Online  | 
|:-----|-----|:-----|:-----|:-----|
| Dialog API 1.1  |  Version 1602 or later | 1.22 or later | 15.20 or later| Currently not supported |

> **Note**: The build number for Office 2016 install via MSI is 16.0.4266.1001.  

To find out more about versions and build numbers, see [Version and build numbers of update channel releases for Office 365 clients](https://technet.microsoft.com/en-us/library/mt592918.aspx)cecenter/default.aspx)

## Office common API requirement sets
For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).

## Dialog API 1.1 
The following are the Dialog APIs in requirement set 1.1. 

## Additional resources

- [Specify Office hosts and API requirements](../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](https://dev.office.com/docs/add-ins/overview/add-in-manifests)
