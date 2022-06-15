---
title: Support Internet Explorer 11
description: Learn how to support Internet Explorer 11 and ES5 Javascript in your add-in.
ms.date: 05/01/2022
ms.localizationpriority: medium
---

# Support Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer still used in Office Add-ins**
>
> Some combinations of platforms and Office versions, including one-time-purchase versions through Office 2019, still use the webview control that comes with Internet Explorer 11 to host add-ins, as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). We recommend (but don't require) that you continue to support these combinations, at least in a minimal way, by providing users of your add-in a graceful failure message when your add-in is launched in the Internet Explorer webview. Keep these additional points in mind:
>
> - Office on the web no longer opens in Internet Explorer. Consequently, [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) no longer tests add-ins in Office on the web using Internet Explorer as the browser.
> - AppSource still tests for combinations of platform and Office *desktop* versions that use Internet Explorer, however it only issues a warning when the add-in does not support Internet Explorer; the add-in is not rejected by AppSource.
> - The [Script Lab tool](../overview/explore-with-script-lab.md) no longer supports Internet Explorer.

Office Add-ins are web applications that are displayed inside IFrames when running on Office on the web. Office Add-ins are displayed using embedded browser controls when running in Office on Windows or Office on the Mac. The embedded browser controls are supplied by the operating system or by a browser installed on the user's computer.

If you plan to support older versions of Windows and Office, your add-in must work in the embeddable browser control that is based on Internet Explorer 11 (IE11). For information about which combinations of Windows and Office use the IE11-based browser control, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> Internet Explorer 11 doesn't support some HTML5 features such as media, recording, and location. If your add-in must support Internet Explorer 11, then you must either design the add-in to avoid these unsupported features or the add-in must detect when Internet Explorer is being used and provide an alternate experience that doesn't use the unsupported features. For more information, see [Determine at runtime if the add-in is running in Internet Explorer](#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer).

## Support for recent versions of JavaScript

Internet Explorer 11 doesn't support JavaScript versions later than ES5. If you want to use the syntax and features of ECMAScript 2015 or later, or TypeScript, you have two options as described in this article. You can also combine these two techniques.

### Use a transpiler

You can write your code in either TypeScript or modern JavaScript and then transpile it at build-time into ES5 JavaScript. The resulting ES5 files are what you upload to your add-in's web application.

There are two popular transpilers. Both of them can work with source files that are TypeScript or post-ES5 JavaScript. They also work with React files (.jsx and .tsx).

- [babel](https://babeljs.io/)
- [tsc](https://www.typescriptlang.org/index.html)

See the documentation for either of them for information about installing and configuring the transpiler in your add-in project. We recommend that you use a task runner, such as [Grunt](https://gruntjs.com/) or [WebPack](https://webpack.js.org/) to automate the transpilation. For a sample add-in that uses tsc, see [Office Add-in Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React). For a sample that uses babel, see [Offline Storage Add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin).

> [!NOTE]
> If you are using Visual Studio (not Visual Studio Code), tsc is probably easiest to use. You can install support for it with a nuget package. For more information, see [JavaScript and TypeScript in Visual Studio 2019](/visualstudio/javascript/javascript-in-vs-2019). To use babel with Visual Studio, create a build script or use the Task Runner Explorer in Visual Studio with tools like the [WebPack Task Runner](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner) or [NPM Task Runner](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner).

### Use a polyfill

A [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) is earlier-version JavaScript that duplicates functionality from more recent versions of JavaScript. The polyfill works with in browsers that don't support the later JavaScript versions. For example, the string method `startsWith` wasn't part of the ES5 version of JavaScript, and so it won't run in Internet Explorer 11. There are polyfill libraries, written in ES5, that define and implement a `startsWith` method. We recommend the [core-js](https://github.com/zloirock/core-js) polyfill library.

To use a polyfill library, load it like any other JavaScript file or module. For example, you can use a `<script>` tag in the add-in's home page HTML file (for example `<script src="/js/core-js.js"></script>`), or you can use an `import` statement in a JavaScript file (for example, `import 'core-js';`). When the JavaScript engine sees a method like `startsWith`, it will first look to see if there is a method of that name built into the language. If there is, it will call the native method. If, and only if, the method isn't built-in, the engine will look in all loaded files for it. So, the polyfilled version isn't used in browsers that support the native version.

Importing the entire core-js library will import all core-js features. You can also import only the polyfills that your Office Add-in requires. For instructions about how to do this, see [CommonJS APIs](https://github.com/zloirock/core-js#commonjs-api). The core-js library has most of the polyfills that you need. There are a few exceptions detailed in the [Missing Polyfills](https://github.com/zloirock/core-js#missing-polyfills) section of the core-js documentation. For example, it doesn't support `fetch`, but you can use the [fetch](https://github.com/github/fetch) polyfill.

For a sample add-in that uses core.js, see [Word Add-in Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).

## Determine at runtime if the add-in is running in Internet Explorer

Your add-in can discover if it is running in Internet Explorer by reading the [window.navigator.userAgent](https://developer.mozilla.org/docs/Web/API/Navigator/userAgent) property. This enables the add-in to either provide an alternate experience or gracefully fail. The following is an example. Note that Internet Explorer sends a string beginning with "Trident" as the value of userAgent.

```javascript
if (navigator.userAgent.indexOf("Trident") === -1) {

    // IE is not the browser. Provide a full-featured version of the add-in here.

} else {

    // IE is the browser. So here, do one of the following: 
    //  1. Provide an alternate experience that does not use any of the HTML5
    //     features that are not supported in IE.
    //  2. Enable the add-in to gracefully fail by putting a message in the UI that
    //     says something similar to: 
    //      "This add-in won't run in your version of Office. Please upgrade to 
    //      either one-time purchase Office 2021 or to a Microsoft 365 account."          

}
```

> [!IMPORTANT]
> It is not usually a good practice to read the `userAgent` property. Be sure you are familiar with the article, [Browser detection using the user agent](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent), including the recommendations and alternatives to reading `userAgent`. In particular, if you are taking option 1 in the `else` clause above, consider using feature detection instead of testing for the user agent.
>
> As of September 30th, 2021, the text in the section [Which part of the user agent contains the information you are looking for?](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent#which_part_of_the_user_agent_contains_the_information_you_are_looking_for) dates from before Internet Explorer 11 was released. It is still generally accurate and the *tables* in the section of the English version of the article are up-to-date. Similarly, the text, and in most cases the tables, in the non-English versions of the article are out-of-date.

## Test an add-in on Internet Explorer

See [Internet Explorer 11 testing](../testing/ie-11-testing.md).

## Additional resources

- [ECMAScript 6 compatibility table](https://kangax.github.io/compat-table/es6/)
- [Can I use... Support tables for HTML5, CSS3, etc](https://caniuse.com/)
