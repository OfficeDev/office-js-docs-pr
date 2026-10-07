> [!IMPORTANT]
> This article is aimed at developers troubleshooting existing contextual Outlook add-ins. We recommend that you not create new contextual Outlook add-ins.
>
> - Entity-based contextual Outlook add-ins are now retired.  
> - Scenarios that previously required regular expression rules in contextual add-ins can now be implemented with [Event-based Outlook add-ins](../develop/event-based-activation.md), which are activated automatically in response to events such as creating a new message or meeting. JavaScript has built-in support for regular expressions, and you can also reference JavaScript libraries that provide advanced regular expression functionality. Unlike contextual add-ins, Event-based add-ins can work in compose contexts as well as read mode.
