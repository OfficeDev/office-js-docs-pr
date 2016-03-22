 

# Understanding API requirement sets

Outlook Add-ins declare what API versions they require by using the [Requirements](https://msdn.microsoft.com/EN-US/library/office/dn592036.aspx) element in their [manifest](https://msdn.microsoft.com/en-us/library/office/fp123693.aspx). Outlook add-ins always include a [Set](https://msdn.microsoft.com/EN-US/library/office/dn592049.aspx) element with a `Name` attribute set to `Mailbox` and a `MinVersion` attribute set to the minimum API requirement set that supports the add-in's scenarios.

For example, the following manifest snippet indicates a minimum requirement set of 1.1:

```
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

All Outlook APIs belong to the `Mailbox` [requirement set](https://msdn.microsoft.com/EN-US/library/office/dn535871.aspx#SpecifyRequirementSets_intro). The `Mailbox` requirement set has versions, and each new set of APIs that we release belongs to a higher version of the set. Not all Outlook clients support the newest set of APIs, but if an Outlook client declares support for a requirement set, it supports all of the APIs in that requirement set.

Setting a minimum requirement set version in the manifest controls which Outlook client the add-in will appear in. If a client does not support the minimum requirement set, it does not load the add-in. For example, if requirement set version 1.3 is specified, this means the add-in will not show up in any Outlook client that doesn't support at least 1.3.

## Using APIs from later requirement sets

Setting a requirement set does not limit the available APIs that the add-in can use. For example, if the add-in specifies requirement set 1.1, but it is running in an Outlook client which support 1.3, the add-in can use APIs from requirement set 1.3\.

To use newer APIs, developers can just check for their existence by using standard JavaScript technique

```
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
| Outlook 2016 | 1.1, 1.2, 1.3 |
| Mac Outlook 2016 | 1.1 |
| Outlook 2013 | 1.1, 1.2 |
| Outlook on the web (Office 365 and Outlook.com) | 1.1, 1.2 |
| Outlook Web App (Exchange 2013 On-Premise) | 1.1, 1.2 |