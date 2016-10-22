# Dialog API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Specify Office hosts and API requirements](../docs/overview/specify-office-hosts-and-api-requirements.md).

Office Add-ins run across multiple versions of Office including Office Online, Office 365 ProPlus, Office for the iPad and Office for the Mac. The following table lists the Dialog API requirement sets, the Office host applications that support that requirement set, and the build versions.

|  Requirement set  |  Office Online  | Office 365 ProPlus  |  Office 2016 for iPad  |  Office 2016 for Mac  |
|:-----|-----|:-----|:-----|:-----|
| Dialog API 1.1  | Currently not supported| 16.0.6741.0000 | 1.22 or later | 15.20 or later|

> **Note**: The build number for Office 2016 install via MSI is 16.0.4266.1001.  

For more information about Office 365 ProPlus, Office 2016 install via MSI, and Office Online Server, see the following topics:

- [Overview of update channels for Office 365 ProPlus](https://technet.microsoft.com/en-us/library/mt455210.aspx)
- [Version and build numbers of update channel releases for Office 365 clients](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [Office Online Servers](https://technet.microsoft.com/en-us/library/jj219437(v=office.16).aspx)
- [Volume Licensing Service Center](https://www.microsoft.com/Licensing/servicecenter/default.aspx)

## Dialog API 1.1 
The following are the Dialog APIs in requirement set 1.1. 

## Office common API requirement sets
For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).

## Additional resources

- [Specify Office hosts and API requirements](../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifeste](https://dev.office.com/docs/add-ins/overview/add-in-manifests)
