---
title: Use polyfills to support Internet Explorer 11 (IE11) in your Office Add-in
description: Expand the reach of your Office Add-in to support the large number of customers still on IE11 by using polyfills. 
ms.date: 05/13/2021
localization_priority: Priority
---
# Use polyfills to support Internet Explorer 11 (IE11) in your Office Add-in

There are still a large number of customers who use the Internet Explorer 11 (IE11) browser, and to reach them, you'll want your Office Add-in to successfully run on it. However, IE11 does not implement many JavaScript features that are available on modern browsers like Microsoft Edge and Google Chrome. This blog shows how to use Polyfills in your Office Add-in to work around missing IE11 functionality.

## How a polyfill works

A polyfill is code that implements a feature on web browsers that do not support the feature. For example, in IE11, the `Number.isNan` function does not exist. If someone had code that needed to call `Number.isNan` they could write a polyfill to ensure their code is compatible with IE11. The following code sample shows how to implement a `Number.isNan` polyfill:

```javascript
if(!Number.isNan) { // If the function is not defined
	// Write a definition for Number.isNan
	Number.isNan = function(x) {
		return typeof x === 'number' && x !== x;
	}
}
```

In the previous example, the code checks if `Number.isNan` exists. If it doesn’t, then we define an implementation for `Number.isNaN` that is used automatically whenever `Number.isNan` is invoked. If the programmer did not include this polyfill, their code would fail to run and throw an exception on certain outdated JavaScript engines.

## Add a polyfill to your Office Add-ins

Let’s look at a feature that works in current browsers but fails in IE11. The `startsWith()` method determines whether a string begins with the characters of a specified string. This method returns `true` if the string begins with the characters; otherwise, it returns `false`.
For example, `'planner'.startsWith('plan')` will return `true`.
The `startsWith()` method is not supported by IE11. So, if you use this method in your script, customers still on IE11 will run into issues. The following code shows how to create a polyfill for this method so your add-in can support customers on IE11.

```javascript
if (!String.prototype.startsWith) {
    Object.defineProperty(String.prototype, 'startsWith', {
        value: function(search, rawPos) {
            var pos = rawPos > 0 ? rawPos|0 : 0;
            return this.substring(pos, pos + search.length) === search;
        }
    });
}
```
> [!NOTE]
> For the previous sample, any copyright is dedicated to the Public Domain. [http://creativecommons.org/publicdomain/zero/1.0/](http://creativecommons.org/publicdomain/zero/1.0/).

If your customer is using a current browser that supports the `startsWith()` method, the previous code will detect this, and nothing will change. However, if they’re using IE11 or another legacy browser, the previous code will ensure compatibility and a good customer experience.

## Use an open source library for polyfills

It would be tedious if you had to write individual definitions for every method that might not be supported by IE11. Fortunately, there are open source libraries available that implement the vast majority of necessary polyfills for IE11 and other legacy browsers. The standard at this time is the **core-js** library, as recommended by Microsoft’s documentation for Office Add-ins: [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Documentation for **core-js** can be found in the [GitHub zloirock/core-js standard library repo](https://github.com/zloirock/core-js).
After you add the **core-js** library to your Office Add-in project, you can add polyfills for all **core-js** features by importing the library at the top of your entry point. The following example shows how to add the **core-js** library to your script. The `run` function shows an example of calling the `startsWith()` method. The **core-js** library contains numerous polyfills, including the `startsWith()` method.

```javascript
import "core-js";
// import "core-js/features/string/starts-with"

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
      console.log('planner'.startsWith('plan'));
    } catch (error) {
    console.error(error);
  }
}
```
The previous example is based on the [Build an Excel task pane add-in](../quickstarts/excel-quickstart-jquery.md) quick start.
In the previous example, the `startsWith()` method works for both current browsers and for IE11. Current browsers don’t need the polyfill provided by **core-js**, while IE11 will use the imported polyfill.
Importing the entire **core-js** library will import all **core-js** features. If you’re concerned about space, then import only the polyfills that your Office Add-in requires. The [core-js documentation](https://github.com/zloirock/core-js#commonjs-api) contains instructions for this. As an example, the previous code shows how to include the required imports `core-js/features/string/starts-with` instead of importing `core-js`.
The **core-js** library typically takes care of most polyfills that you need. There are a few exceptions detailed in the [Missing Polyfills](https://github.com/zloirock/core-js#missing-polyfills) section of the [core-js documentation](https://github.com/zloirock/core-js#commonjs-api). Notably, it does not support `fetch`, but you can use the [GitHub window.fetch JavaScript polyfill](https://github.com/github/fetch).
If you need to use any methods that don’t work on IE11 and are not supported by these or other libraries, you will need to write your own custom polyfill, as detailed previously.

## Test if you need a polyfill

There is really no need to try to determine which browser your customer is using and create different cases for different browsers. If your polyfill is written correctly, you’ll automatically handle compatibility issues for whatever browser they may be using.
However, if you wish to determine if your code is doing something not supported in IE11, there are a couple of solutions. One simple approach is to test your add-in on an environment running IE11 and manually monitor the console.
The second, and more methodical solution, is to install JSHint which is available via Node Package Manager (npm), and then configure it - [JSHint Options Reference](https://jshint.com/docs/options/) - with "esversion" set to 5. This will cause JSHint to produce warnings about your code if it does not conform to ES5 standards, which are 99% compatible with IE11. This won’t find all polyfill issues, but it’ll likely catch many potential issues that need to be polyfilled.

## More details about Office Add-ins and IE11

Office Add-ins are web applications that are displayed inside iframes when running on Office on the web. Office Add-ins are displayed using embedded browser controls when running in Office on Windows, or on the Mac. The embedded browser controls are supplied by the operating system or by a browser installed on the user's computer. For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).
On Windows, add-ins will run in an embedded browser control supplied by either Internet Explorer 11, Microsoft Edge Legacy, or Microsoft Edge (Chromium). Which embedded control is used depends on the installed OS and Office versions.
However, you can encounter cross-browser compatibility issues with JavaScript. In addition, different browsers such as Chrome, Firefox, Microsoft Edge, and Internet Explorer 11 have different proprietary features. Certain features that work in one browser may not be compatible in another browser. To reach and support the greatest number of users, you’ll want to ensure that your Office Add-in works across all of these browser environments.

## Summary

When developing Office Add-ins, take care to ensure your add-ins are compatible with older browsers such as IE11. To do this, be sure to add polyfills to fill in the gaps where IE11 falls short. Oftentimes, all you need to do is install **core-js** and import it in your script. For any method unsupported by IE11 and unsupported by **core-js** or other polyfill libraries, you’ll need to write your own polyfill.

## Additional Resources

- [ECMAScript 6 compatibility table](https://kangax.github.io/compat-table/es6/)
- [Can I use... Support tables for HTML5, CSS3, etc](https://caniuse.com/)
- [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md)
- [GitHub - zloirock/core-js: Standard Library](https://github.com/zloirock/core-js)
- [GitHub - github/fetch: A window.fetch JavaScript polyfill](https://github.com/github/fetch)