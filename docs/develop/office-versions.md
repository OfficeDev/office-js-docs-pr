{pagebreak}

## Office versions:  Office 2016 vs. Office 365 (MSI vs. Click-to-Run); Deferred vs. Current channels; Insider tracks {#office-versions}

To develop and distribute Add-ins that use the new Office 2016 API model, you need either Office 2016 or Office 365 (the subscription-based superset that includes all 2016 features).  This seems reasonably straightforward, but the devil is in the details.


A> ### The golden path
A>
A> The simplest possible case is when both you (the developer) and your end-users have the latest-and-greatest versions of Office 365. Certainly, for development / prototyping needs, having Office 365 with the latest updates will be simplest.  But if you are an ISV (Independent Software Vendor) and hence have no control over what version your customers are running; or if you work inside of an enterprise that might not be on the bleeding edge, that's where understanding Office versions becomes important.

A> ### Why you should care
A>
A> Different versions (and categories of versions) offer different API surface areas.  For example, Office 2016 RTM only offered the first batch of the new 2016+ wave of Excel and Word APIs; those APIs have been greatly expanded since then (see the discussion on [API versioning](#api-versioning) in the next section) .  Likewise, other functionality -- most notably, Add-in Commands (ribbon extensibility) and the ability to launch dialogs -- were not present in the original RTM version.


{pagebreak}

In the next few pages, I will describe the different installation possibilities.  It may help to bear the following image in mind:

![](images/Office-Versions.jpg)


{pagrebreak}

**Office 2016 vs. Office 365**

The first place where the API surface-area forks off is at the split between the MSI-based installation of Office 2016, and the subscription-based (sometimes called "click-to-run") installation of Office 365.

Let's pause for a second to talk about Office 365, as I've seen some confusion about the term.  [Wikipedia](https://en.wikipedia.org/wiki/Office_365) explains it nicely:

> Office 365 is the brand name Microsoft uses for a group of software and services subscriptions, which together provide productivity software and related services to subscribers. For consumers, the service allows the use of Microsoft Office apps on Windows and macOS, provides storage space on Microsoft's cloud storage service OneDrive, and grants 60 Skype minutes per month. For business users, Office 365 offers service plans providing e-mail and social networking services through hosted versions of Exchange Server, Skype for Business Server, SharePoint and Office Online, integration with Yammer, as well as access to the Microsoft Office software.

So, for those coming from the SharePoint world: yes, SharePoint Online is part of an Office 365 subscription, as are the Office Online in-browser editors that come with it. But, it is not the *only* part of the subscription.  Getting access to the *same desktop/mac Office programs that you know and love* is also part of that same subscription (as is getting iOS and Android versions of those Word, Excel, PPT, etc. programs).


Now, back to the APIs:  if you have Office 2016 (non-subscription), you will *only* have the initial set of the new wave of Excel and Word APIs (`ExcelApi 1.1` and `WordApi 1.1`).  Or to put it even clearer:  you will only have the *initial set of extensibility functionality* -- period.  So in addition to missing improvements to the Excel and Word APIs, you will also lack other add-in functionality like the ability to customize the Ribbon or launch dialogs.

It's also worth noting that the original RTM offering of the APIs did have some bugs.  Some were more innocent than others[^bugs], but -- in my personal opinion -- I would treat RTM as more of *the start of a journey* into rich host-specific APIs, rather than a destination of its own.

So again: Office 2016, from an API / extensibility standpoint, is frozen in time... frozen to the functionality that was there when it shipped in September 2015.  (For those unclear on how a "2016"-branded product could have shipped in 2015, I guess it's a bit like buying next year's car models. Or, perhaps, `2015 + 3/4-of-a-year` rounds to `2016`...)

Meanwhile, Office 365 means "subscription".  This translates to being on the latest-and-greatest stable build (where "stable" for enterprise might be a build that's a few months old; more on that below).

If you want access to the latest-and-greatest API functionality -- which as a developer, you absolutely do -- you *must* be on a subscription-based installation of Office, rather than the frozen-in-time Office 2016 MSI installation.  Moreover, for most of the new functionality, you probably want your customers to be on a subscription-based installation, as well.


**Office 365 flavors for the Consumer**

The Consumer (non-business) versions of Office 365 include **Office 365 Personal** and **Office 365 Home** (with the only difference between the two being the number of active devices on the subscription -- 1 PC or Mac, 1 Tablet, and 1 Phone, as opposed to 5 of each).  There is also **Office 365 University**, which is the same as Personal, but offers activating on *two* devices rather than one.  For all three, the difference is merely the cost of the plan and the number of devices supported; they are all 100% equivalent from an API & functionality standpoint.

The Consumer versions of Office 356 are updated each month, with the updates installed silently and automatically. Thus, consumer versions of Office 365, provided the computer is connected to the internet, will always have access to the latest-and-greatest functionality.  The default is the "Current" channel (i.e., what is publicly available worldwide), but the adventurous user (developer) can also opt in to be on one of the *Insider* tracks.  The Insider tracks come in two flavors, Insider Fast and Insider Slow, with the *Fast* being really-bleeding-edge, and *Slow* being a few weeks behind, anchored around more stable builds.  In both cases, they let you preview the forthcoming functionality a month or two ahead of the general public.  For developers, this can be particularly useful for trying out the latest APIs ahead of your customers, allowing you to deliver new functionality as soon as it's publicly-available on your customers' machines.  Combined with using the Beta CDN for Office.js, it can also let you provide real-time API feedback back to the team, before the APIs get cemented and go live!  To become an Insider, see <https://products.office.com/en-us/office-insider>.


**Office 365 flavors for Enterprise**

For users of the enterprise / business flavor of Office 365, there are also a number of options (typically handled by the IT administrator).  Like with the Consumer versions, there is a "Current" channel (latest-and-greatest stable build, updated monthly) -- and similarly, there is a "First Release for Current channel", which essentially is the same as the "Insider" builds in the consumer version.

However, risk-averse enterprises may also choose to be on a Deferred channel, which updates once every four months instead of once every month.  Moreover, these enterprises can also stay on the Deferred channel for four or even eight months, before jumping ahead to a newer build.  Thus, a business on the Deferred channel may still be a fair bit behind the developer in terms of what API functionality is available (though less behind than someone on the RTM build of Office 2016).


**Office on other platforms (Mac, iOS, Online)**

For non-PC platforms, there is also a span of time before different functionality lights up.  This is sometimes dependent not just on the *delay* between something being code-complete and getting in front of customers' hands (i.e., the difference between Insider and Current and Deferred), but also on the order in which functionality gets implemented on these platforms.  For the Excel APIs to date, I have seen them light up on most platforms at roughly the same time; for Word, the Desktop has generally been ahead of Online.  For the non-API functionality (i.e., dialogs, ribbon extensibility), these have also generally come to the Desktop first, followed by Online and Mac.  The different speeds of implementation is why it's important to keep in mind not just Office host versions, but also API versions and Requirement Sets -- the subject of the very next section.



[^bugs]: There are a number of bugs that come to mind.  On the Excel side, for example, reading back values (even without manipulating the document) would blow the undo stack -- not a catastrophic issue, but irksome all the same.  On the Word side, there were some issues with document-identity when accessing items in a collection (e.g., paragraphs), whereby the proxy object remembered *its index in the collection*, but not the actual document entity that it belonged to; and so, if the document was manipulated, the proxy object (paragraph) would now point at the wrong item.