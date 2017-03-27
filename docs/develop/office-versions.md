# Understanding Office versions

|||
|:--|:--|:--|
|![An image of the Building Office Add-ins using Office.js book cover](../../images/book-cover.png)|**Note:** This article is an excerpt from the book [Building Office Add-ins using Office.js](https://leanpub.com/buildingofficeaddins) by Michael Zlatkovsky, available for purchase as an e-book on LeanPub.com.<br/><br/>Copyright © 2016-2017 by Michael Zlatkovsky, all rights reserved.|

To develop and distribute add-ins that use the new Office 2016 API model, you need either Office 2016 or Office 365 (the subscription-based superset that includes all Office 2016 features).  This seems reasonably straightforward, but the devil is in the details.


**The golden path**

The simplest possible case is when both you (the developer) and your end-users have the latest-and-greatest versions of Office 365. Certainly, for development / prototyping needs, having Office 365 with the latest updates will be simplest.  But if you are an ISV (Independent Software Vendor) and hence have no control over what version your customers are running; or if you work inside of an enterprise that might not be on the bleeding edge, that's where understanding Office versions becomes important.

**Why you should care**

Different versions (and categories of versions) offer different API surface areas.  For example, Office 2016 RTM only offered the first batch of the new wave of Excel and Word APIs; those APIs have been greatly expanded since then.  Likewise, other functionality -- most notably, add-in commands (ribbon extensibility) and the ability to launch dialog boxess -- were not present in the original RTM version.

In the next few pages, I will describe the different installation possibilities.  It may help to bear the following image in mind:

![An image that shows the Office 2016 MSI release, and the Office 365 subscription. The subscription has two versions - consumer and enterprise. The consumer version has current, insider slow, and insider fast releases. The enterprise version has deferred, first release for deferred channel, current, and first release for current channel releases.](../../images/office-versions.png)


## Office 2016 vs. Office 365

The first place where the API surface-area forks off is at the split between the MSI-based installation of Office 2016, and the subscription-based (sometimes called "click-to-run") installation of Office 365.

Let's pause for a second to talk about Office 365, as I've seen some confusion about the term.  

Office 365 is a subscription service that provides the most up-to-date tools from Microsoft. There are Office 365 plans for home and personal use, as well as for small and midsized businesses, large enterprises, schools, and nonprofits. All Office 365 plans for home and personal use include Office 2016 with the fully installed Office applications like Word, PowerPoint, and Excel, as well as online storage, and more. Office 365 for business users provides email and social networking services through Exchange Server, Skype for Business Server, Office Online, and Yammer integration, in addition to Office software.

So, for those coming from the SharePoint world: yes, SharePoint Online is part of an Office 365 subscription, as are the Office Online in-browser editors that come with it. But, it is not the *only* part of the subscription.  Getting access to the *same desktop/mac Office programs that you know and love* is also part of that same subscription (as is getting iOS and Android versions of those Word, Excel, PowerPoint, etc. programs).


Now, back to the APIs:  if you have Office 2016 (non-subscription), you will *only* have the initial set of the new wave of Excel and Word APIs (`ExcelApi 1.1` and `WordApi 1.1`).  Or to put it even clearer:  you will only have the *initial set of extensibility functionality* -- period.  So in addition to missing improvements to the Excel and Word APIs, you will also lack other add-in functionality like the ability to customize the Ribbon or launch dialogs.

It's also worth noting that the original RTM offering of the APIs does have some bugs.  In my personal opinion, I would treat RTM as more of *the start of a journey* into rich host-specific APIs, rather than a destination of its own.

So again: Office 2016, from an API / extensibility standpoint, is frozen in time... frozen to the functionality that was there when it shipped in September 2015.  

Meanwhile, Office 365 means "subscription".  This translates to being on the latest-and-greatest stable build (where "stable" for enterprise might be a build that's a few months old; more on that below).

If you want access to the latest-and-greatest API functionality -- which as a developer, you absolutely do -- you *must* be on a subscription-based installation of Office, rather than the frozen-in-time Office 2016 MSI installation.  Moreover, for most of the new functionality, you probably want your customers to be on a subscription-based installation, as well.


## Office 365 flavors for the Consumer

The Consumer (non-business) versions of Office 365 include **Office 365 Personal** and **Office 365 Home** (with the only difference between the two being the number of active devices on the subscription -- 1 PC or Mac, 1 Tablet, and 1 Phone, as opposed to 5 of each).  There is also **Office 365 University**, which is the same as Personal, but offers activating on *two* devices rather than one.  For all three, the difference is merely the cost of the plan and the number of devices supported; they are all 100% equivalent from an API & functionality standpoint.

The Consumer versions of Office 356 are updated each month, with the updates installed silently and automatically. Thus, consumer versions of Office 365, provided the computer is connected to the internet, will always have access to the latest-and-greatest functionality.  The default is the "Current" channel (i.e., what is publicly available worldwide), but the adventurous user (developer) can also opt in to be on one of the *Insider* tracks.  The Insider tracks come in two flavors, Insider Fast and Insider Slow, with the *Fast* being really-bleeding-edge, and *Slow* being a few weeks behind, anchored around more stable builds.  In both cases, they let you preview the forthcoming functionality a month or two ahead of the general public.  For developers, this can be particularly useful for trying out the latest APIs ahead of your customers, allowing you to deliver new functionality as soon as it's publicly-available on your customers' machines.  Combined with using the Beta CDN for Office.js, it can also let you provide real-time API feedback back to the team, before the APIs get cemented and go live!  To become an Insider, see <https://products.office.com/en-us/office-insider>.


## Office 365 flavors for Enterprise

For users of the enterprise / business flavor of Office 365, there are also a number of options (typically handled by the IT administrator).  Like with the Consumer versions, there is a "Current" channel (latest-and-greatest stable build, updated monthly) -- and similarly, there is a "First Release for Current channel", which essentially is the same as the "Insider" builds in the consumer version.

However, risk-averse enterprises may also choose to be on a Deferred channel, which updates once every four months instead of once every month.  Moreover, these enterprises can also stay on the Deferred channel for four or even eight months, before jumping ahead to a newer build.  Thus, a business on the Deferred channel may still be a fair bit behind the developer in terms of what API functionality is available (though less behind than someone on the RTM build of Office 2016).


## Office on other platforms (Mac, iOS, Online)

For non-PC platforms, there is also a span of time before different functionality lights up.  This is sometimes dependent not just on the *delay* between something being code-complete and getting in front of customers' hands (i.e., the difference between Insider and Current and Deferred), but also on the order in which functionality gets implemented on these platforms.  For the Excel APIs to date, I have seen them light up on most platforms at roughly the same time; for Word, Office for desktop has generally been ahead of Office Online.  For the non-API functionality (i.e., dialog boxes, ribbon extensibility), these have also generally come to the desktop first, followed by Office Online and Mac.  The different speeds of implementation is why it's important to keep in mind not just Office host versions, but also API versions and Requirement Sets.
 
