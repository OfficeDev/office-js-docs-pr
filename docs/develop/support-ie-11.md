---
title: Support Internet Explorer 11
description: 'Learn how to support Internet Explorer 11 and ES5 Javascript in your add-in.'
ms.date: 06/18/2021
localization_priority: Normal
---

# Support Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer still used in Office Add-ins**
>
> Microsoft is ending support for Internet Explorer, but this doesn't significantly affect Office Add-ins. Some combinations of platforms and Office versions, including all one-time-purchase versions through Office 2019, will continue to use the webview control that comes with Internet Explorer 11 to host add-ins, as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Moreover, support for these combinations, and hence for Internet Explorer, is still required for add-ins submitted to [AppSource](/office/dev/store/submit-to-appsource-via-partner-center). Two things *are* changing:
>
> - AppSource no longer tests add-ins in Office on the web using Internet Explorer as the browser. But AppSource still tests for combinations of platform and Office *desktop* versions that use Internet Explorer.
> - The [Script Lab tool](../overview/explore-with-script-lab.md) will stop working in Internet Explorer sometime in 2021.

Office Add-ins are web applications that are displayed inside IFrames when running on Office on the web. Office Add-ins are displayed using embedded browser controls when running in Office on Windows or Office on the Mac. The embedded browser controls are supplied by the operating system or by a browser installed on the user's computer.

If you plan to market your add-in through AppSource or you plan to support older versions of Windows and Office, your add-in must work in the embeddable browser control that is based on Internet Explorer 11 (IE11). For information about which combinations of Windows and Office use the IE11-based browser control, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> Internet Explorer 11 doesn't support some HTML5 features such as media, recording, and location. If your add-in must support Internet Explorer 11, then you can't use these features.

Internet Explorer 11 doesn't support JavaScript versions later than ES5. If you want to use the syntax and features of ECMAScript 2015 or later, or TypeScript, you have two options as described in this article. You can also combine these two techniques.

## Use a transpiler

You can write your code in either TypeScript or modern JavaScript and then transpile it at build-time into ES5 JavaScript. The resulting ES5 files are what you upload to your add-in's web application.

There are two popular transpilers. Both of them can work with source files that are TypeScript or post-ES5 JavaScript. They also work with React files (.jsx and .tsx).

- [babel](https://babeljs.io/)
- [tsc](https://www.typescriptlang.org/index.html)

See the documentation for either of them for information about installing and configuring the transpiler in your add-in project. We recommend that you use a task runner, such as [Grunt](https://gruntjs.com/) or [WebPack](https://webpack.js.org/) to automate the transpilation. For a sample add-in that uses tsc, see [Office Add-in Microsoft Graph React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/3ce0e1b74152dbbe8306a091696bc4455c04c0a1/Samples/auth/Office-Add-in-Microsoft-Graph-React). For a sample that uses babel, see [Offline Storage Add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/3ce0e1b74152dbbe8306a091696bc4455c04c0a1/Samples/Excel.OfflineStorageAddin).

> [!NOTE]
> If you are using Visual Studio (not Visual Studio Code), tsc is probably easiest to use. You can install support for it with a nuget package. For more information, see [JavaScript and TypeScript in Visual Studio 2019](/visualstudio/javascript/javascript-in-vs-2019). To use babel with Visual Studio, create a build script or use the Task Runner Explorer in Visual Studio with tools like the [WebPack Task Runner](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner) or [NPM Task Runner](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner).

## Use a polyfill

A [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) is earlier-version JavaScript that duplicates functionality from more recent versions of JavaScript. The polyfill works with in browsers that don't support the later JavaScript versions. For example, the string method `startsWith` wasn't part of the ES5 version of JavaScript, and so it won't run in Internet Explorer 11. There are polyfill libraries, written in ES5, that define and implement a `startsWith` method. We recommend the [core-js](https://github.com/zloirock/core-js) polyfill library.

To use a polyfill library, load it like any other JavaScript file or module. For example, you can use a `<script>` tag in the add-in's home page HTML file (for example `<script src="/js/core-js.js"></script>`), or you can use an `import` statement in a JavaScript file (for example, `import 'core-js';`). When the JavaScript engine sees a method like `startsWith`, it will first look to see if there is a method of that name built into the language. If there is, it will call the native method. If, and only if, the method isn't built-in, the engine will look in all loaded files for it. So, the polyfilled version isn't used in browsers that support the native version.

Importing the entire core-js library will import all core-js features. You can also import only the polyfills that your Office Add-in requires. For instructions about how to do this, see [CommonJS APIs](https://github.com/zloirock/core-js#commonjs-api). The core-js library has most of the polyfills that you need. There are a few exceptions detailed in the [Missing Polyfills](https://github.com/zloirock/core-js#missing-polyfills) section of the core-js documentation. For example, it doesn't support `fetch`, but you can use the [fetch](https://github.com/github/fetch) polyfill.

For a sample add-in that uses core.js, see [Word Add-in Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).

## Testing an add-in on Internet Explorer

See [Internet Explorer 11 testing](../testing/ie-11-testing.md).

## Additional resources

- [ECMAScript 6 compatibility table](https://kangax.github.io/compat-table/es6/)
- [Can I use... Support tables for HTML5, CSS3, etc](https://caniuse.com/)
