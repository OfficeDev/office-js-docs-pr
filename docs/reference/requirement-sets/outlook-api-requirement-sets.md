# Outlook JavaScript API requirement sets

Outlook add-ins declare what API versions they require by using the [Requirements](/javascript/office/manifest/requirements) element in their [manifest](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests). Outlook add-ins always include a [Set](/javascript/office/manifest/set) element with a `Name` attribute set to `Mailbox` and a `MinVersion` attribute set to the minimum API requirement set that supports the add-in's scenarios.

For example, the following manifest snippet indicates a minimum requirement set of 1.1:

```xml
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

All Outlook APIs belong to the `Mailbox` [requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements). The `Mailbox` requirement set has versions, and each new set of APIs that we release belongs to a higher version of the set. Not all Outlook clients support the newest set of APIs, but if an Outlook client declares support for a requirement set, it supports all of the APIs in that requirement set.

Setting a minimum requirement set version in the manifest controls which Outlook client the add-in will appear in. If a client does not support the minimum requirement set, it does not load the add-in. For example, if requirement set version 1.3 is specified, this means the add-in will not show up in any Outlook client that doesn't support at least 1.3.

## Using APIs from later requirement sets

Setting a requirement set does not limit the available APIs that the add-in can use. For example, if the add-in specifies requirement set 1.1, but it is running in an Outlook client which support 1.3, the add-in can use APIs from requirement set 1.3.

To use newer APIs, developers can just check for their existence by using standard JavaScript technique:

```js
if (item.somePropertyOrFunction !== undefined) {
  item.somePropertyOrFunction ...
}
```

No such checks are necessary for any APIs which are present in the requirement set version specified in in the manifest.

## Choosing a minimum requirement set

Developers should use the earliest requirement set that contains the critical set of APIs for their scenario, without which the add-in won't work.

## Clients

The following clients support Outlook add-ins.

| Client | Supported API requirement sets |
| --- | --- |
| Outlook 2019 for Windows | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2019 for Mac | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook 2016 (Click-to-Run) for Windows | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2016 (MSI) for Windows | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Outlook 2016 for Mac | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook 2013 for Windows | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Outlook for iPhone | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Outlook for Android | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Outlook on the web (Office 365 and Outlook.com) | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook Web App (Exchange 2013 on-premises) | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1) |
| Outlook Web App (Exchange 2016 on-premises) | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3) |

> [!NOTE]
> Support for 1.3 in Outlook 2013 was added as part of the [December 8, 2015, update for Outlook 2013 (KB3114349)](https://support.microsoft.com/kb/3114349). Support for 1.4 in Outlook 2013 was added as part of the [September 13, 2016, update for Outlook 2013 (KB3118280)](https://support.microsoft.com/help/3118280).
