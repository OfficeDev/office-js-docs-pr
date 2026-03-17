---
title: Office Add-ins glossary of terms
description: A glossary of terms commonly used throughout the Office Add-ins documentation.
ms.date: 03/16/2026
ms.topic: glossary
ms.localizationpriority: medium
---

# Office Add-ins glossary

This glossary lists terms commonly used throughout the Office Add-ins documentation.

## add-in

Office Add-ins are web applications that extend Office applications. These web applications add new functionality to the Office application, such as bringing in external data, automating processes, or embedding interactive objects in Office documents.

Office Add-ins differ from VBA, COM, and VSTO add-ins because they offer cross-platform support (usually web, Windows, Mac, and iPad) and are based on standard web technologies (HTML, CSS, and JavaScript). The primary programming language of an Office Add-in is JavaScript or TypeScript.

## add-in commands

**Add-in commands** are UI elements, such as buttons and menus, that extend the Office UI for your add-in. When users select an add-in command element, they initiate actions such as running JavaScript code or displaying the add-in in a task pane. Add-in commands let your add-in look and feel like a part of Office, which gives users more confidence in your add-in. To learn more, see [Add-in commands](../design/add-in-commands.md).

See also: [ribbon, ribbon button](#ribbon-ribbon-button).

## add-in only manifest

An **add-in only manifest** is an XML file that defines how an Office Add-in integrates with Office applications. It specifies the add-in's settings, capabilities, and UI entry points. The add-in only manifest is the original manifest format for Office Add-ins and doesn't support combining with other Microsoft 365 extensions. To learn more, see [Office Add-ins manifest](../develop/add-in-manifests.md) and [Office Add-ins with the add-in only manifest](../develop/xml-manifest-overview.md).

See also: [unified manifest for Microsoft 365](#unified-manifest-for-microsoft-365).

## append-on-send, prepend-on-send

**Append-on-send** and **prepend-on-send** are Outlook add-in features that automatically add content to the end or beginning of a message body when the user sends it. By using these features, add-ins can insert disclaimers, signatures, or other content without requiring the user to take any action. To learn more, see [Prepend or append content to a message or appointment body on send](../outlook/append-on-send.md).

## application

**Application** refers to an Office application. The Office applications that support Office Add-ins are Excel, OneNote, Outlook, PowerPoint, Project, and Word.

See also: [client](#client), [host](#host), [Office application, Office client](#office-application-office-client).

## application-specific API

Application-specific APIs provide strongly-typed objects that interact with objects that are native to a specific Office application. For example, you call the Excel JavaScript APIs for access to worksheets, ranges, tables, charts, and more. Application-specific APIs are currently available for Excel, OneNote, PowerPoint, Visio, and Word. To learn more, see [Application-specific API model](../develop/application-specific-api-model.md).

See also: [Common API](#common-api).

## AppSource

**AppSource** is the former name for **Microsoft Marketplace**, the online store where users and administrators can discover, try, and deploy Office Add-ins, Teams apps, and other Microsoft 365 extensions. Add-in developers publish to Microsoft Marketplace through [Partner Center](#partner-center).

See also: [Microsoft Marketplace](#microsoft-marketplace), [Partner Center](#partner-center).

## Centralized Deployment

**Centralized Deployment** is used by Microsoft 365 administrators to deploy Office Add-ins that use the [add-in only manifest](#add-in-only-manifest) to users and groups within their organization. Through the Microsoft 365 admin center, administrators can assign add-ins to specific users, groups, or the entire organization without requiring each user to install individually. To learn more, see [Deploy and publish Office Add-ins](../publish/publish.md).

See also: [integrated apps portal](#integrated-apps-portal), [sideloading](#sideloading).

## client

**Client** typically refers to an Office application. The Office applications, or clients, that support Office Add-ins are Excel, OneNote, Outlook, PowerPoint, Project, and Word.

See also: [application](#application), [host](#host), [Office application, Office client](#office-application-office-client), [Office desktop application, Office desktop client, desktop client, desktop](#office-desktop-application-office-desktop-client-desktop-client-desktop).

## COM add-in

A **COM add-in** is a legacy Office extensibility model that uses the Component Object Model (COM). COM add-ins run only on Windows and require installation on each user's machine. Office Add-ins (web add-ins) are the modern cross-platform replacement for COM add-ins.

See also: [add-in](#add-in), [VSTO add-in](#vsto-add-in), [VBA](#vba).

## Common API

Use Common APIs to access features such as UI, dialogs, and client settings that are common across multiple Office applications. This API model uses [callbacks](https://developer.mozilla.org/docs/Glossary/Callback_function), which you use to specify only one operation in each request sent to the Office application.

Common APIs were introduced in Office.js with Office 2013. Some Common APIs are legacy APIs from the early 2010s. Excel, PowerPoint, and Word all have Common API functionality, but most of this functionality is replaced or superseded by the application-specific API model. Use the application-specific APIs when possible.

Other Common APIs, such as the Common APIs related to Outlook, UI, and authentication, are the modern and preferred APIs for these purposes. For details about the Common API object model, see [Common JavaScript API object model](../develop/office-javascript-api-object-model.md).

See also: [application-specific API](#application-specific-api).

## compose mode, read mode

**Compose mode** is the state in Outlook when a user is creating or editing a message or appointment. **Read mode** is the state when a user is viewing a received message or appointment. Outlook add-ins can specify different behavior depending on the current mode. To learn more, see the "Extension points" section of [Outlook add-ins overview](../outlook/outlook-add-ins-overview.md#extension-points).

## content add-in

**Content add-ins** are webviews, or web browser views, that are embedded directly into Excel, OneNote, or PowerPoint documents. Content add-ins give users access to interface controls that run code to modify documents or display data from a data source. Use content add-ins when you want to embed functionality directly into the document. To learn more, see [Content Office Add-ins](../design/content-add-ins.md).

See also: [webview](#webview).

## content delivery network (CDN)

A **content delivery network** or **CDN** is a distributed network of servers and data centers. It typically provides higher resource availability and performance when compared to a single server or data center.

## contextual add-in

A **contextual add-in** is an Outlook add-in that activates automatically when certain text patterns are detected in a message or appointment. Contextual add-ins use regular expression rules to match content such as addresses, phone numbers, or tracking numbers. They only activate in read mode. To learn more, see [Contextual Outlook add-ins](../outlook/contextual-outlook-add-ins.md).

See also: [compose mode, read mode](#compose-mode-read-mode).

## Contoso

**Contoso** Ltd. (also known as Contoso and Contoso University) is a fictional company that Microsoft uses as an example company and domain.

## custom function

A **custom function** is a user-defined function that's packaged with an Excel add-in. By defining functions in JavaScript as part of an add-in, developers can add new functions beyond the typical Excel features. Users in Excel can access custom functions just as they would any native function in Excel. To learn more, see [Create custom functions in Excel](../excel/custom-functions-overview.md).

[!include[Excel custom functions definition](../includes/excel-custom-functions-definition.md)]

## custom functions runtime

A **custom functions runtime** is a [JavaScript-only runtime](../testing/runtimes.md#javascript-only-runtime) that runs custom functions on some combinations of Office host and platform. It has no UI and can't interact with Office.js APIs. If your add-in only has custom functions, this runtime is a good lightweight option. If your custom functions need to interact with the task pane or Office.js APIs, configure a [shared runtime](../testing/runtimes.md#shared-runtime). To learn more, see [Configure your Office Add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

See also: [runtime](#runtime), [shared runtime](#shared-runtime).

## custom functions-only add-in

An add-in that contains a custom function but no UI such as a task pane. The custom functions in this kind of add-in run in a [JavaScript-only runtime](../testing/runtimes.md#javascript-only-runtime). A custom function that includes a UI can use either a shared runtime or a combination of a JavaScript-only runtime and an HTML-supporting runtime. Use a shared runtime if your add-in has a UI.

See also: [custom function](#custom-function), [custom functions runtime](#custom-functions-runtime).

## delegate access, shared mailbox

**Delegate access** is an Outlook scenario where one user manages another user's mailbox or calendar. A **shared mailbox** is a mailbox that multiple users can access for sending and receiving email from a common address (for example, info@contoso.com). Outlook add-ins require specific API support to work correctly in these shared scenarios. To learn more, see [Implement shared folders and shared mailbox scenarios in an Outlook add-in](../outlook/delegate-access.md).

## Entra ID

**Entra ID** (formerly Azure Active Directory or Azure AD) is Microsoft's cloud-based identity and access management service. Office Add-ins use Entra ID for authenticating users and obtaining tokens for accessing Microsoft Graph and other protected resources.

See also: [nested app authentication (NAA)](#nested-app-authentication-naa), [single sign-on (SSO)](#single-sign-on-sso).

## event-based activation

**Event-based activation** is a feature that runs add-in code automatically in response to specific events without requiring the user to explicitly launch the add-in. To learn more, see [Event-based activation](../develop/event-based-activation.md).

See also: [Smart Alerts](#smart-alerts).

## extension point

An **extension point** is a configuration element in the add-in manifest that defines where and how the add-in integrates with the Office UI. Common extension points include ribbon command surfaces, context menus, and event-based launch triggers. The available extension points vary by Office application.

See also: [add-in commands](#add-in-commands), [add-in only manifest](#add-in-only-manifest), [unified manifest for Microsoft 365](#unified-manifest-for-microsoft-365).

## function command

Function commands are buttons or menu items that run JavaScript functions. Unlike task pane commands, function commands don't display any user interface other than the command button or menu item itself.

See also: [add-in commands](#add-in-commands).

## host

`<Host>` typically refers to an Office application. The Office applications, or hosts, that support Office Add-ins are Excel, OneNote, Outlook, PowerPoint, Project, and Word.

See also: [application](#application), [client](#client), [Office application, Office client](#office-application-office-client).

## Information Rights Management (IRM)

**Information Rights Management** (**IRM**) is an Office feature that protects documents and email messages from unauthorized access, forwarding, printing, or copying. IRM-protected content might restrict the functionality available to Office Add-ins. To learn more, see [Mail items protected by IRM](../outlook/outlook-add-ins-overview.md#mail-items-protected-by-irm).

## Integrated apps portal

The **integrated apps portal** in the Microsoft 365 admin center is the primary method for Microsoft 365 administrators to deploy Office Add-ins to users and groups within their organization. Through the integrated apps portal, administrators can assign add-ins to specific users, groups, or the entire organization without requiring each user to install individually. For more information, see [Integrated apps portal in the Microsoft 365 admin center](../publish/publish.md#integrated-apps-portal-in-the-microsoft-365-admin-center).

See also: [Microsoft 365 admin center](#microsoft-365-admin-center), [sideloading](#sideloading).
## JavaScript-only runtime

A **JavaScript-only runtime** is a lightweight runtime environment that includes a JavaScript engine but no HTML rendering engine, DOM, localStorage, or cookies. Use it to run event-based tasks, custom functions in Excel, and integrated spam reporting in Outlook. To learn more, see [Runtimes in Office Add-ins](../testing/runtimes.md#javascript-only-runtime).

See also: [custom functions runtime](#custom-functions-runtime), [runtime](#runtime), [shared runtime](#shared-runtime).

## load and sync pattern

The **load and sync pattern** is the core programming paradigm for application-specific APIs in Office.js. Because Office represents objects as [proxy objects](#proxy-object), your code first queues property reads by using `load()`, and then sends the batch to Office by using `context.sync()`. This batching approach minimizes round trips between the add-in and the Office application. To learn more, see [Application-specific API model](../develop/application-specific-api-model.md).

See also: [proxy object](#proxy-object), [RequestContext](#requestcontext).

## Long-Term Service Channel (LTSC)

**LTSC** refers to the perpetual version of Office that's available through a volume-licensing agreement between Microsoft and your company.

See also: [perpetual](#perpetual), [volume-licensed, volume-licensed perpetual, volume licensing](#volume-licensed-volume-licensed-perpetual-volume-licensing).

## Microsoft 365 admin center

Tenant administrators use the **Microsoft 365 admin center** web portal to manage users, licenses, and organizational settings. In the context of Office Add-ins, administrators use the admin center to deploy add-ins to users and groups through the [integrated apps portal](#integrated-apps-portal) or [Centralized Deployment](#centralized-deployment). For more information, see [Publish Office Add-ins](../publish/publish.md).

## Microsoft 365 Agents Toolkit

The **Microsoft 365 Agents Toolkit** (formerly Teams Toolkit) is a VS Code extension for creating and managing Microsoft 365 extensions, including Office Add-ins. It supports the unified manifest format. To learn more, see [Create Office Add-in projects with Microsoft 365 Agents Toolkit](../develop/agents-toolkit-overview.md).

See also: [unified manifest for Microsoft 365](#unified-manifest-for-microsoft-365), [Yeoman generator, Yo Office](#yeoman-generator-yo-office).

## Microsoft Marketplace

**Microsoft Marketplace** (formerly AppSource or Microsoft Commercial Marketplace) is Microsoft's online store where users and administrators can discover, try, and deploy Office Add-ins, Teams apps, and other Microsoft 365 extensions. Add-in developers publish to Microsoft Marketplace through [Partner Center](#partner-center).

See also: [Partner Center](#partner-center).

## Microsoft Graph

**Microsoft Graph** is a unified REST API that provides access to data and intelligence across Microsoft 365 services, including Outlook mail, OneDrive files, Teams conversations, and more. Office Add-ins commonly call Microsoft Graph to access user data beyond what the Office application exposes directly. To learn more, see [Authorize to Microsoft Graph from an Office Add-in](../develop/authorize-to-microsoft-graph-without-sso.md).

## nested app authentication (NAA)

**Nested app authentication** (**NAA**) is the recommended authentication pattern for Office Add-ins. NAA uses the MSAL.js library with a nested app protocol to obtain access tokens, enabling single sign-on across platforms. NAA replaces the legacy `getAccessToken` SSO approach. To learn more, see [Enable single sign-on in an Office Add-in with nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md).

See also: [Entra ID](#entra-id), [single sign-on (SSO)](#single-sign-on-sso).

## new Outlook on Windows, classic Outlook on Windows

**New Outlook on Windows** is the modern Outlook client for Windows based on web technologies, offering closer parity with Outlook on the web. **Classic Outlook on Windows** is the legacy Outlook application on Windows. Depending on the scenario, the two clients may have different add-in feature support, so the documentation distinguishes between them when noting platform availability. To learn more, see [Develop Outlook add-ins for the new Outlook on Windows](../outlook/one-outlook.md).

See also: [Office desktop application, Office desktop client, desktop client, desktop](#office-desktop-application-office-desktop-client-desktop-client-desktop).

## Office application, Office client

**Office client** refers to an Office application. The Office applications, or clients, that support Office Add-ins are Excel, OneNote, Outlook, PowerPoint, Project, and Word.

See also: [application](#application), [client](#client), [host](#host), [Office desktop application, Office desktop client, desktop client, desktop](#office-desktop-application-office-desktop-client-desktop-client-desktop).

## Office cache

The **Office cache** stores resources and data used by Office Add-ins. This cache prevents an add-in from repeatedly downloading the resources it needs, thereby improving its performance.

See also: [web cache](#web-cache), [Wef cache](#wef-cache).

## Office desktop application, Office desktop client, desktop client, desktop

**Office desktop client** refers to an Office application that runs natively on Windows or on Mac. The Office desktop clients that support Office Add-ins are Excel on Windows and on Mac, Outlook on Windows ([new](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627) and classic) and on Mac, PowerPoint on Windows and on Mac, Project on Windows, and Word on Windows and on Mac.

See also: [application](#application), [client](#client), [Office application, Office client](#office-application-office-client).

## Office.js

**Office.js** is the JavaScript library that provides the APIs for building Office Add-ins. Add-ins reference Office.js from the Microsoft CDN (`https://appsforoffice.microsoft.com/lib/1/hosted/office.js`), and it includes both the Common API and application-specific APIs for interacting with Office documents, email, presentations, and more. To learn more, see [Referencing the Office JavaScript API library](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).

See also: [application-specific API](#application-specific-api), [Common API](#common-api).

## Office on the web

**Office on the web** refers to the browser-based versions of Office applications (Excel, Word, PowerPoint, Outlook, and OneNote) that you access through a web browser. Office on the web is one of the primary platforms that support Office Add-ins, alongside Windows, Mac, mobile, and iPad.

See also: [platform](#platform).

## on-behalf-of flow (OBO)

The **on-behalf-of flow** (**OBO**) is a legacy OAuth 2.0 authentication pattern where a server-side component exchanges a user's token for a new token with different permissions or scopes. To learn more, see [Authorize to Microsoft Graph with legacy Office SSO](../develop/authorize-to-microsoft-graph.md). For a modern authentication experience, use the Microsoft Authentication Library (MSAL) with [nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md) (NAA).

See also: [nested app authentication (NAA)](#nested-app-authentication-naa), [single sign-on (SSO)](#single-sign-on-sso).

## Partner Center

**Partner Center** is the Microsoft developer portal where add-in publishers create seller accounts, submit add-in packages, and manage listings on [Microsoft Marketplace](#microsoft-marketplace). Publishing to end users through Microsoft Marketplace requires validation and approval through Partner Center.

See also: [Microsoft Marketplace](#microsoft-marketplace).

## perpetual

**Perpetual** refers to versions of Office that you get through a volume-licensing agreement or retail channels.

Other Microsoft content might use the term **non-subscription** to represent this concept.

See also: [retail, retail perpetual](#retail-retail-perpetual), [volume-licensed, volume-licensed perpetual, volume licensing](#volume-licensed-volume-licensed-perpetual-volume-licensing).

## pinnable task pane

A **pinnable task pane** is an Outlook task pane that users can pin so it stays visible as they switch between different mail items. When users pin the task pane, it stays open and updates its content context as the selected item changes. To learn more, see [Implement a pinnable task pane in Outlook](../outlook/pinnable-taskpane.md).

See also: [task pane](#task-pane).

## platform

A **platform** usually refers to the operating system running the Office application. Platforms that support Office Add-ins include Windows, Mac, iPad, and web browsers.

## proxy object

A **proxy object** is a local JavaScript object that represents an object in the Office document, such as a range, table, or worksheet. Operations on proxy objects queue locally and don't send to Office until `context.sync()` is called. This batching design minimizes round trips and improves performance. To learn more, see [Application-specific API model](../develop/application-specific-api-model.md).

See also: [load and sync pattern](#load-and-sync-pattern), [RequestContext](#requestcontext).

## quick start

A **quick start** is a high-level description of key skills and knowledge required for the basic operation of a particular program. In the Office Add-ins documentation, a quick start is an introduction to developing an add-in for a particular application, such as Outlook. A quick start contains a series of steps that an add-in developer can complete in about five minutes, resulting in a functioning add-in and functional development environment.

See also: [tutorial](#tutorial).

## RequestContext

A **RequestContext** is the object required for interacting with Office through the application-specific APIs. Created by using `Excel.run()`, `Word.run()`, or similar methods, the `RequestContext` maintains the connection to the Office application. Use it to create proxy objects, queue operations, and synchronize state by using `context.sync()`. To learn more, see [Application-specific API model](../develop/application-specific-api-model.md).

See also: [load and sync pattern](#load-and-sync-pattern), [proxy object](#proxy-object).

## requirement set

[!include[Requirement set note](../includes/office-js-requirement-sets.md)]

## retail, retail perpetual

**Retail** refers to perpetual versions of Office available through retail channels. These versions don't include versions provided by a Microsoft 365 subscription or volume-licensing agreement.

Other Microsoft content might use the terms **one-time purchase** or **consumer** to represent this concept.

See also: [perpetual](#perpetual).

## ribbon, ribbon button

A **ribbon** is a command bar that organizes an application's features into a series of tabs or buttons at the top of a window. A **ribbon button** is one of the buttons within this series. For more information, see [Show or hide the ribbon in Office](https://support.microsoft.com/office/d946b26e-0c8c-402d-a0f7-c6efa296b527#ID0EBBD=Newer_Versions).

## runtime

A **runtime** is the host environment (including a JavaScript engine and usually also an HTML rendering engine) that the add-in runs in. In Office on Windows and Office on Mac, the runtime is an embedded browser control (or webview) such as Edge WebView2 or Safari WKWebView. Different parts of an add-in run in separate runtimes. For example, add-in commands, custom functions, and task pane code typically use separate runtimes unless you configure a [shared runtime](../testing/runtimes.md#shared-runtime). For more information, see [Runtimes in Office Add-ins](../testing/runtimes.md) and [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

See also: [custom functions runtime](#custom-functions-runtime), [shared runtime](#shared-runtime), [webview](#webview).

## Script Lab

**Script Lab** is a free Office Add-in available for Excel, Outlook, Word, and PowerPoint that you can use to write and run Office.js code snippets directly within the Office application. It's a useful tool for learning the Office JavaScript APIs and prototyping add-in functionality. To learn more, see [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md).

## sensitivity label

A **sensitivity label** is a classification marking applied to Office documents or Outlook email messages indicating their confidentiality level, such as Public, Internal, or Confidential. Add-ins can read sensitivity labels and enforce policies based on the applied label. To learn how to manage sensitivity labels in Outlook, see [Manage the sensitivity label of your message or appointment in compose mode](../outlook/sensitivity-label.md).

## setless API

An API in the Office JavaScript Library that isn't included in any requirement set.

See also: [requirement set](#requirement-set).

## shared runtime

A **shared runtime** enables code in your task pane, function commands, and custom functions to run in the same runtime and continue running even when the task pane is closed. Code in dialogs generally runs in a separate runtime even when you configure the add-in to use a shared runtime. To learn more, see [shared runtime](../testing/runtimes.md#shared-runtime) and [Tips for using the shared runtime in your Office Add-in](https://devblogs.microsoft.com/microsoft365dev/tips-for-using-the-shared-javascript-runtime-in-your-office-add-in%e2%80%af/).

See also: [custom functions runtime](#custom-functions-runtime), [runtime](#runtime).

## SharePoint app catalog

A **SharePoint app catalog** is a special SharePoint site collection that you use to distribute Office Add-ins within your organization. Administrators upload add-in only manifest files to the catalog, making the add-ins available to users in the organization without publishing to Microsoft Marketplace. To learn more, see [SharePoint app catalog deployment](../publish/publish.md#sharepoint-app-catalog-deployment).

See also: [AppSource](#appsource), [Centralized Deployment](#centralized-deployment).

## sideloading

**Sideloading** is the process of installing an Office Add-in directly for testing purposes without going through Microsoft Marketplace or the integrated apps portal. Developers sideload add-ins during development to test functionality in Office applications. The sideloading method varies by platform. To learn more, see [Sideload an Office Add-in for testing](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).

See also: [Centralized Deployment](#centralized-deployment), [integrated apps portal](#integrated-apps-portal).

## single sign-on (SSO)

**Single sign-on** (**SSO**) allows an Office Add-in to authenticate the user by using their existing Office login credentials, eliminating the need for a separate sign-in step. SSO in Office Add-ins uses Microsoft Entra ID tokens to authorize access to Microsoft Graph and other services. The recommended implementation uses [nested app authentication (NAA)](#nested-app-authentication-naa). To learn more, see [Enable single sign-on in an Office Add-in with nested app authentication](../develop/enable-nested-app-authentication-in-your-add-in.md).

See also: [Entra ID](#entra-id), [nested app authentication (NAA)](#nested-app-authentication-naa).

## Smart Alerts

**Smart Alerts** is an Outlook feature that uses event-based activation to run add-in code when a user sends a message or appointment. The add-in can validate message properties, such as required fields, categories, or attachments, and optionally block the send if conditions aren't met. To learn more, see [Handle OnMessageSend and OnAppointmentSend events in your Outlook add-in with Smart Alerts](../outlook/onmessagesend-onappointmentsend-events.md).

See also: [event-based activation](#event-based-activation).

## subscription

**Subscription** refers to versions of Office available with a Microsoft 365 subscription.

## task pane

Task panes are interface surfaces, or webviews, that typically appear on the right side of the window within Excel, Outlook, PowerPoint, and Word. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to or can't embed functionality directly into the document. To learn more, see [Task panes in Office Add-ins](../design/task-pane-add-ins.md).

See also: [webview](#webview).

## tutorial

A **tutorial** is a teaching aid designed to help people learn to use a product or procedure. In the Office Add-ins context, a tutorial guides an add-in developer through the complete add-in development process for a particular application, such as Excel. This process involves following 20 or more steps and is a greater time investment than a [quick start](#quick-start).

See also: [quick start](#quick-start).

## unified manifest for Microsoft 365

The **unified manifest for Microsoft 365** is a JSON-based manifest format that developers use to define Office Add-ins, Teams apps, and other Microsoft 365 extensions in a single manifest file. To learn more, see [Office Add-ins manifest](../develop/add-in-manifests.md) and [Unified manifest for Microsoft 365](../develop/unified-manifest-overview.md).

See also: [add-in only manifest](#add-in-only-manifest), [Microsoft 365 Agents Toolkit](#microsoft-365-agents-toolkit).

## VBA

**VBA** (Visual Basic for Applications) is a legacy programming language that's built into Office desktop applications for creating macros and automating tasks. VBA macros run only on desktop platforms and aren't part of the modern Office Add-ins platform. Office Add-ins are the cross-platform alternative to VBA-based automation.

See also: [add-in](#add-in), [COM add-in](#com-add-in), [VSTO add-in](#vsto-add-in).

## volume-licensed, volume-licensed perpetual, volume licensing

**Volume-licensed** refers to a perpetual version of Office that you get through a volume-licensing agreement between Microsoft and your company.

Other Microsoft content might use the term **commercial** to represent this concept.

See also: [Long-Term Service Channel (LTSC)](#long-term-service-channel-ltsc), [perpetual](#perpetual).

## VSTO add-in

A **VSTO add-in** (Visual Studio Tools for Office add-in) is a legacy Office extensibility model that uses the .NET Framework. VSTO add-ins run only on Windows and are built by using Visual Studio. Office Add-ins (web add-ins) are the modern cross-platform replacement for VSTO add-ins.

See also: [add-in](#add-in), [COM add-in](#com-add-in), [VBA](#vba).

## web add-in

**Web add-in** is a legacy term for an Office Add-in. The Microsoft 365 documentation might use this term when it needs to distinguish modern Office Add-ins from other types of add-ins like VBA, COM, or VSTO.

See also: [add-in](#add-in).

## web cache

The **web cache** temporarily stores web-based resources and data used by an individual Office Add-in.

See also: [Office cache](#office-cache), [Wef cache](#wef-cache).

## webview

A **webview** is an element or view that displays web content inside an application. Content add-ins and task panes both contain embedded web browsers and are examples of webviews in Office Add-ins.

See also: [content add-in](#content-add-in), [task pane](#task-pane).

## Wef cache

The **Wef cache** locally stores resources and data for all installed Office Add-ins.

See also: [Office cache](#office-cache), [web cache](#web-cache).

## XLL

An **XLL** add-in is an Excel add-in file that provides user-defined functions and has the file extension **.xll**. An XLL file is a type of dynamic link library (DLL) file that only Excel can open. You must write XLL add-in files in C or C++. Custom functions are the modern equivalent of XLL user-defined functions. Custom functions offer support across platforms and are backwards compatible with XLL files. To learn more, see [Extend custom functions with XLL add-ins](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf).

See also: [custom function](#custom-function).

## Yeoman generator, Yo Office

The [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md) uses the open source [Yeoman](https://github.com/yeoman/yo) tool to generate an Office Add-in through the command line. The `yo office` command runs the Yeoman generator for Office Add-ins. The Office Add-ins quick starts and tutorials use the Yeoman generator.

## See also

- [Office Add-ins additional resources](resources-links-help.md)
