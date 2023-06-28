For [some versions of Office and Windows](../concepts/browsers-used-by-office-web-add-ins.md), the JavaScript engine in which add-ins run is linked to the Trident webview that's provided by Internet Explorer. This engine doesn't support versions of JavaScript later than ES5. This means that without special handling, the JavaScript files that your add-in serves cannot use syntax, types, or methods that were added to the language after ES5. This doesn't mean that you must *write* in ES5 syntax. You have two other options:

- Write your code in [ECMAScript 2015](https://www.w3schools.com/Js/js_es6.asp) (also called ES6) or later JavaScript, or in TypeScript, and then compile your code to ES5 JavaScript using a compiler such as [babel](https://babeljs.io/) or [tsc](https://www.typescriptlang.org/index.html).
- Write in ECMAScript 2015 or later JavaScript, but also load a [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) library such as [core-js](https://github.com/zloirock/core-js) that enables IE to run your code.

For more information about these options, see [Support older Microsoft webviews and Office versions](../develop/support-ie-11.md).
