---
title: Support older Microsoft browsers and Office versions
description: Learn how to support support older Microsoft browsers and Office versions in your add-in.
ms.date: 12/12/2022
ms.localizationpriority: medium
---

# Support older Microsoft browsers and Office versions

> [!IMPORTANT]
> **Internet Explorer and Microsoft Edge Legacy are still used in Office Add-ins**
>
> Some combinations of platforms and Office versions, including perpetual versions through Office 2019, still use the webview controls that come with Internet Explorer 11 and Microsoft Edge Legacy (EdgeHTML-based) to host add-ins, as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). We recommend (but don't require) that you support these combinations, at least in a minimal way, by providing users of your add-in a graceful failure message when your add-in is launched in these webviews. Keep these additional points in mind:
>
> - Office on the web no longer opens in Internet Explorer or Microsoft Edge Legacy. Consequently, [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) doesn't test add-ins in Office on the web on these browsers.
> - AppSource still tests for combinations of platform and Office *desktop* versions that use Internet Explorer and Microsoft Edge Legacy. However, it only issues a warning when the add-in doesn't support these browsers. The add-in isn't rejected by AppSource.
> - The [Script Lab tool](../overview/explore-with-script-lab.md) no longer supports Internet Explorer.

Office Add-ins are web applications that are displayed inside IFrames when running on Office on the web. Office Add-ins are displayed using embedded browser controls when running in Office on Windows or Office on the Mac. The embedded browser controls are supplied by the operating system or by a browser installed on the user's computer.

If you plan to support older versions of Windows and Office, your add-in must work in the embeddable browser controls used by these versions. For example, browser controls based on Internet Explorer 11 (IE11) or Microsoft Edge Legacy (EdgeHTML-based). For information about which combinations of Windows and Office use these browser controls, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

## Determine the browser the add-in is running in at runtime

Your add-in can discover the browser it's running in by reading the [window.navigator.userAgent](https://developer.mozilla.org/docs/Web/API/Navigator/userAgent) property. This enables the add-in to either provide an alternate experience or gracefully fail. The following is an example that determines whether the add-in is running in IE11 or Microsoft Edge Legacy.

```javascript
if (navigator.userAgent.indexOf("Trident") !== -1) {
    /*
       IE11 is the browser in use. Do one of the following:
        1. Provide an alternate add-in experience that doesn't use any of the HTML5
           features that aren't supported in IE11.
        2. Enable the add-in to gracefully fail by adding a message to the UI that
           says something similar to:
           "This add-in won't run in your version of Office. Please upgrade either to
           perpetual Office 2021 or to a Microsoft 365 account."
    */
} else if (navigator.userAgent.indexOf("Edge") !== -1) {
    /*
       Microsoft Edge Legacy is the browser in use. Do one of the following:
        1. Provide an alternate add-in experience that's supported in Microsoft Edge Legacy.
        2. Enable the add-in to gracefully fail by adding a message to the UI that
           says something similar to:
           "This add-in won't run in your version of Office. Please upgrade either to
           perpetual Office 2021 or to a Microsoft 365 account."
    */
} else {
    /* 
       Another browser, other than IE11 or Microsoft Edge Legacy, is in use.
       Provide a full-featured version of the add-in here.
    */
}
```

> [!IMPORTANT]
> It's not usually a good practice to read the `userAgent` property. Be sure you're familiar with the article, [Browser detection using the user agent](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent), including the recommendations and alternatives to reading `userAgent`. In particular, if you're providing an alternate add-in experience to support the use of Internet Explorer 11, consider using feature detection instead of testing for the user agent.
>
> As of September 30th, 2021, the text in the section [Which part of the user agent contains the information you are looking for?](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent#which_part_of_the_user_agent_contains_the_information_you_are_looking_for) dates from before Internet Explorer 11 was released. It's still generally accurate and the *tables* in the section of the English version of the article are up-to-date. Similarly, the text, and in most cases the tables, in the non-English versions of the article are out-of-date.

## Review browser and Office version support information

For more information on how to support specific browsers and Office versions, select the applicable tab.

# [Internet Explorer](#tab/ie)

> [!IMPORTANT]
> Internet Explorer 11 doesn't support some HTML5 features such as media, recording, and location. If your add-in must support Internet Explorer 11, then you must either design the add-in to avoid these unsupported features or the add-in must detect when Internet Explorer is being used and provide an alternate experience that doesn't use the unsupported features. For more information, see [Determine at runtime if the add-in is running in IE11 or Microsoft Edge Legacy](#determine-at-runtime-if-the-add-in-is-running-in-ie11-or-microsoft-edge-legacy).

## Support for recent versions of JavaScript

Internet Explorer 11 doesn't support JavaScript versions later than ES5. If you want to use the syntax and features of ECMAScript 2015 or later, or TypeScript, you have two options as described in this article. You can also combine these two techniques.

### Use a transpiler

You can write your code in either TypeScript or modern JavaScript and then transpile it at build-time into ES5 JavaScript. The resulting ES5 files are what you upload to your add-in's web application.

There are two popular transpilers. Both of them can work with source files that are TypeScript or post-ES5 JavaScript. They also work with React files (.jsx and .tsx).

- [babel](https://babeljs.io/)
- [tsc](https://www.typescriptlang.org/index.html)

See the documentation for either of them for information about installing and configuring the transpiler in your add-in project. We recommend that you use a task runner, such as [Grunt](https://gruntjs.com/) or [WebPack](https://webpack.js.org/) to automate the transpilation. For a sample add-in that uses tsc, see [Office Add-in Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React). For a sample that uses babel, see [Offline Storage Add-in](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin).

> [!NOTE]
> If you're using Visual Studio (not Visual Studio Code), tsc is probably easiest to use. You can install support for it with a nuget package. For more information, see [JavaScript and TypeScript in Visual Studio 2019](/visualstudio/javascript/javascript-in-vs-2019). To use babel with Visual Studio, create a build script or use the Task Runner Explorer in Visual Studio with tools like the [WebPack Task Runner](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner) or [NPM Task Runner](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner).

### Use a polyfill

A [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) is earlier-version JavaScript that duplicates functionality from more recent versions of JavaScript. The polyfill works in browsers that don't support the later JavaScript versions. For example, the string method `startsWith` wasn't part of the ES5 version of JavaScript, and so it won't run in Internet Explorer 11. There are polyfill libraries, written in ES5, that define and implement a `startsWith` method. We recommend the [core-js](https://github.com/zloirock/core-js) polyfill library.

To use a polyfill library, load it like any other JavaScript file or module. For example, you can use a `<script>` tag in the add-in's home page HTML file (for example `<script src="/js/core-js.js"></script>`), or you can use an `import` statement in a JavaScript file (for example, `import 'core-js';`). When the JavaScript engine sees a method like `startsWith`, it will first look to see if there's a method of that name built into the language. If there is, it will call the native method. If, and only if, the method isn't built-in, the engine will look in all loaded files for it. So, the polyfilled version isn't used in browsers that support the native version.

Importing the entire core-js library will import all core-js features. You can also import only the polyfills that your Office Add-in requires. For instructions about how to do this, see [CommonJS APIs](https://github.com/zloirock/core-js#commonjs-api). The core-js library has most of the polyfills that you need. There are a few exceptions detailed in the [Missing Polyfills](https://github.com/zloirock/core-js#missing-polyfills) section of the core-js documentation. For example, it doesn't support `fetch`, but you can use the [fetch](https://github.com/github/fetch) polyfill.

For a sample add-in that uses core.js, see [Word Add-in Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).

## Test an add-in on Internet Explorer

See [Internet Explorer 11 testing](../testing/ie-11-testing.md).

# [Microsoft Edge Legacy](#tab/edge)

## Troubleshoot Microsoft Edge Legacy issues

If you encounter issues as you develop your add-in to support Microsoft Edge Legacy, see the "Troubleshoot Microsoft Edge issues" section of [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md#troubleshoot-microsoft-edge-issues) for guidance.

## Debug an add-in that supports Microsoft Edge Legacy

To debug your add-in that supports Microsoft Edge Legacy, use one of the following options.

- [Debug add-ins using developer tools in Microsoft Edge Legacy](../testing/debug-add-ins-using-devtools-edge-legacy.md)
- [Debug add-ins using the Microsoft Office Add-in Debugger Extension for Visual Studio Code](../testing/debug-with-vs-extension.md)

---

## See also

- [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md)
- [ECMAScript 6 compatibility table](https://kangax.github.io/compat-table/es6/)
- [Can I use... Support tables for HTML5, CSS3, etc](https://caniuse.com/)
