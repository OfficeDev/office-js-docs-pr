---
title: Outlook JavaScript API requirement sets
description: 'Learn more about the Outlook JavaScript API requirement sets.'
ms.date: 07/10/2020
ms.prod: outlook
localization_priority: Priority
---

# Outlook JavaScript API requirement sets

Outlook add-ins declare what API versions they require by using the [Requirements](../manifest/requirements.md) element in their [manifest](../../develop/add-in-manifests.md). Outlook add-ins always include a [Set](../manifest/set.md) element with a `Name` attribute set to `Mailbox` and a `MinVersion` attribute set to the minimum API requirement set that supports the add-in's scenarios.

For example, the following manifest snippet indicates a minimum requirement set of 1.1.

```xml
<Requirements>
  <Sets>
    <Set Name="Mailbox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

All Outlook APIs belong to the `Mailbox` [requirement set](../../develop/specify-office-hosts-and-api-requirements.md). The `Mailbox` requirement set has versions, and each new set of APIs that we release belongs to a higher version of the set. Not all Outlook clients support the newest set of APIs, but if an Outlook client declares support for a requirement set, generally it supports all of the APIs in that requirement set (check the documentation on a specific API or feature for any exceptions).

Setting a minimum requirement set version in the manifest controls which Outlook client the add-in will appear in. If a client does not support the minimum requirement set, it does not load the add-in. For example, if requirement set version 1.3 is specified, this means the add-in will not show up in any Outlook client that doesn't support at least 1.3.

> [!NOTE]
> To use APIs in any of the numbered requirement sets, you should reference the **production** library on the CDN (https://appsforoffice.microsoft.com/lib/1/hosted/office.js).
>
> For information about using preview APIs, see the [Using preview APIs](#using-preview-apis) section later in this article.

## Using APIs from later requirement sets

Setting a requirement set does not limit the available APIs that the add-in can use. For example, if the add-in specifies requirement set "Mailbox 1.1", but it is running in an Outlook client which supports "Mailbox 1.3", the add-in can use APIs from requirement set "Mailbox 1.3".

To use a newer API, developers can check if a particular application supports the requirement set by doing the following.

```js
if (Office.context.requirements.isSetSupported('Mailbox', '1.3')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

Alternatively, developers can check for the existence of a newer API by using standard JavaScript technique.

```js
if (item.somePropertyOrFunction !== undefined) {
  // Use item.somePropertyOrFunction.
  item.somePropertyOrFunction;
}
```

No such checks are necessary for any APIs which are present in the requirement set version specified in the manifest.

## Choosing a minimum requirement set

Developers should use the earliest requirement set that contains the critical set of APIs for their scenario, without which the add-in won't work.

## Requirement sets supported by Exchange servers and Outlook clients

In this section, we note the range of requirement sets supported by Exchange server and Outlook clients. For details about server and client requirements for running Outlook add-ins, see [Outlook add-ins requirements](../../outlook/add-in-requirements.md).

> [!IMPORTANT]
> If your target Exchange server and Outlook client support different requirement sets, then you're restricted to the lower requirement set range. For example, if an add-in is running in Outlook 2016 on Mac (highest requirement set: 1.6) against Exchange 2013 (highest requirement set: 1.1), your add-in is limited to requirement set 1.1.

### Exchange server support

The following servers support Outlook add-ins.

| Product | Major Exchange version | Supported API requirement sets |
|---|---|---|
| Exchange Online | Latest build | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md) |
| Exchange on-premises | 2019 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
|| 2016 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
|| 2013 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md) |

### Outlook client support

Add-ins are supported in Outlook on the following platforms.

| Platform | Major Office/Outlook version | Supported API requirement sets |
|---|---|---|
| Windows | Microsoft 365 subscription | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<sup>1</sup> |
|| 2019 one-time purchase (retail) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<sup>1</sup> |
|| 2019 one-time purchase (volume-licensed) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md) |
|| 2016 one-time purchase | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)<sup>2</sup> |
|| 2013 one-time purchase | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)<sup>2</sup>, [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)<sup>2</sup> |
| Mac | current UI<br>(connected to a Microsoft 365 subscription) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md) |
|| new UI (preview)<sup>3</sup><br>(connected to a Microsoft 365 subscription) | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |
|| 2019 one-time purchase | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |
|| 2016 one-time purchase | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |
| iOS | Microsoft 365 subscription | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)<sup>4</sup> |
| Android | Microsoft 365 subscription | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)<sup>4</sup> |
| Web browser | modern Outlook UI when connected to<br>Exchange Online: Microsoft 365 subscription, Outlook.com | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md), [1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md), [1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md) |
|| classic Outlook UI when connected to<br>Exchange on-premises | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md), [1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md), [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md), [1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md), [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), [1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |

> [!NOTE]
> <sup>1</sup> Support for 1.8 in Outlook on Windows with a Microsoft 365 subscription or a retail one-time purchase is available from version 1910 (build 12130.20272). For more details according to your version, see the update history page for [Office 2019](/officeupdates/update-history-office-2019) or [Microsoft 365](/officeupdates/update-history-office365-proplus-by-date) and how to [find your Office client version and update channel](https://support.office.com/article/about-office-what-version-of-office-am-i-using-932788b8-a3ce-44bf-bb09-e334518b8b19).
>
> <sup>2</sup> Support for 1.3 in Outlook 2013 was added as part of the [December 8, 2015, update for Outlook 2013 (KB3114349)](https://support.microsoft.com/kb/3114349). Support for 1.4 in Outlook 2013 was added as part of the [September 13, 2016, update for Outlook 2013 (KB3118280)](https://support.microsoft.com/help/3118280). Support for 1.4 in Outlook 2016 (one-time purchase) was added as part of the [July 3, 2018, update for Office 2016 (KB4022223)](https://support.microsoft.com/help/4022223).
>
> <sup>3</sup> Preview support for the new Outlook on Mac is available from version 16.38.506. For more information, see the [New Outlook on Mac](../../outlook/compare-outlook-add-in-support-in-outlook-for-mac.md#new-outlook-on-mac-preview) section.
>
> <sup>4</sup> Currently, there are additional considerations when designing and implementing add-ins for mobile clients. For example, the only supported mode is Message Read. For more details, see [code considerations when adding support for add-in commands for Outlook Mobile](../../outlook/add-mobile-support.md#code-considerations).

> [!TIP]
> You can distinguish between classic and modern Outlook in a web browser by checking your mailbox toolbar.
>
> **modern**
>
> ![partial screenshot of the modern Outlook toolbar](../../images/outlook-on-the-web-new-toolbar.png)
>
> **classic**
>
> ![partial screenshot of the classic Outlook toolbar](../../images/outlook-on-the-web-classic-toolbar.png)

## Using preview APIs

New Outlook JavaScript APIs are first introduced in "preview" and later become part of a specific, numbered requirement set after sufficient testing occurs and user feedback is acquired. To provide feedback about a preview API, please use the feedback mechanism at the end of the web page where the API is documented.

> [!NOTE]
> Preview APIs are subject to change and are not intended for use in a production environment.

For more details about the preview APIs, see [Outlook API Preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md).
