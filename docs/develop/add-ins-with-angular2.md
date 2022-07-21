---
title: Develop Office Add-ins with Angular
description: Use Angular to create an Office Add-in as a single page application.
ms.date: 07/08/2021
ms.localizationpriority: medium
---

# Develop Office Add-ins with Angular

This article provides guidance for using Angular 2+ to create an Office Add-in as a single page application.

> [!NOTE]
> Do you have something to contribute based on your experience using Angular to create Office Add-ins? You can contribute to [this article in GitHub](https://github.com/OfficeDev/office-js-docs-pr/blob/master/docs/develop/add-ins-with-angular2.md) or provide your feedback by submitting an [issue](https://github.com/OfficeDev/office-js-docs-pr/issues) in the repo.

For an Office Add-ins sample that's built using the Angular framework, see [Word Style Checking Add-in Built on Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).

## Install the TypeScript type definitions

Open a Node.js window and enter the following at the command line.

```command&nbsp;line
npm install --save-dev @types/office-js
```

## Bootstrapping must be inside Office.initialize

On any page that calls the Office, Word, or Excel JavaScript APIs, your code must first assign a function to `Office.initialize`. (If you have no initialization code, the function body can just be empty "`{}`" symbols, but you must not leave the `Office.initialize` function undefined. For details, see [Initialize your Office Add-in](initialize-add-in.md).) Office calls this function immediately after it has initialized the Office JavaScript libraries.

**Your Angular bootstrapping code must be called inside the function that you assign to `Office.initialize`** to ensure that the Office JavaScript libraries have initialized first. The following is a simple example that shows how to do this. This code should be in the main.ts file of the project.

```js
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { AppModule } from './app.module';

Office.initialize = function () {
  const platform = platformBrowserDynamic();
  platform.bootstrapModule(AppModule);
};
```

## Use the hash location strategy in the Angular application

Navigating between routes in the application might not work if you don't specify the hash location strategy. You can do this in one of two ways. First, you can specify a provider for the location strategy in your app module, as shown in the following example. It goes into the app.module.ts file.

```js
import { LocationStrategy, HashLocationStrategy } from '@angular/common';
// Other imports suppressed for brevity

@NgModule({
  providers: [
    { provide: LocationStrategy, useClass: HashLocationStrategy },
    // Other providers suppressed
  ],
  // Other module properties suppressed
})
export class AppModule { }
```

If you define your routes in a separate routing module, there is an alternative way to specify the hash location strategy. In your routing module's .ts file, pass a configuration object to the `forRoot` function that specifies the strategy. The following code is an example.

```js
import { RouterModule, Routes } from '@angular/router';
// Other imports suppressed for brevity

const routes: Routes = // route definitions go here

@NgModule({
  imports: [RouterModule.forRoot(routes, { useHash: true })],
  exports: [RouterModule]
})
export class AppRoutingModule { }
```

## Use the Office dialog API with Angular

The Office Add-in dialog API enables your add-in to open a page in a nonmodal dialog box that can exchange information with the main page, which is typically in a task pane.

The [displayDialogAsync](/javascript/api/office/office.ui) method takes a parameter that specifies the URL of the page that should open in the dialog box. Your add-in can have a separate HTML page (different from the base page) to pass to this parameter, or you can pass the URL of a route in your Angular application.

It is important to remember, if you pass a route, that the dialog box creates a new window with its own execution context. Your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog box. So this technique launches a second instance of your single page application in the dialog box. Code that changes variables in the dialog box does not change the task pane version of the same variables. Similarly, the dialog box has its own session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property), which is not accessible from code in the task pane.  

## Trigger the UI update

In an Angular app, the UI sometimes does not update. This is because that part of the code runs out of the Angular zone. The solution is to put the code in the zone, as shown in the following example.

```js
import { NgZone } from '@angular/core';

export class MyComponent {
  constructor(private zone: NgZone) { }

  myFunction() {
    this.zone.run(() => {
      // the codes that need update the UI
    });
  }
}
```

## Use Observable

Angular uses RxJS (Reactive Extensions for JavaScript), and RxJS introduces `Observable` and `Observer` objects to implement asynchronous processing. This section provides a brief introduction to using `Observables`; for more detailed information, see the official [RxJS](https://rxjs-dev.firebaseapp.com/) documentation.

An `Observable` is like a `Promise` object in some ways - it is returned immediately from an asynchronous call, but it might not resolve until some time later. However, while a `Promise` is a single value (which can be an array object), an `Observable` is an array of objects (possibly with only a single member). This enables code to call [array methods](https://www.w3schools.com/jsref/jsref_obj_array.asp), such as `concat`, `map`, and `filter`, on `Observable` objects.

### Push instead of pull

Your code "pulls" `Promise` objects by assigning them to variables, but `Observable` objects "push" their values to objects that *subscribe* to the `Observable`. The subscribers are `Observer` objects. The benefit of the push architecture is that new members can be added to the `Observable` array over time. When a new member is added, all the `Observer` objects that subscribe to the `Observable` receive a notification.

The `Observer` is configured to process each new object (called the "next" object) with a function. (It is also configured to respond to an error and a completion notification. See the next section for an example.) For this reason, `Observable` objects can be used in a wider range of scenarios than `Promise` objects. For example, in addition to returning an `Observable` from an AJAX call, the way you can return a `Promise`, an `Observable` can be returned from an event handler, such as the "changed" event handler for a text box. Each time a user enters text in the box, all the subscribed `Observer` objects react immediately using the latest text and/or the current state of the application as input.

### Wait until all asynchronous calls have completed

When you want to ensure that a callback only runs when every member of a set of `Promise` objects has resolved, use the `Promise.all()` method.

```js
myPromise.all([x, y, z]).then(
  // TODO: Callback logic goes here
)
```

To do the same thing with an `Observable` object, you use the [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md) method.  

```js
const source = Observable.forkJoin([x, y, z]);

const subscription = source.subscribe(
  x => {
    // TODO: Callback logic goes here
  },
  err => console.log('Error: ' + err),
  () => console.log('Completed')
);
```

## Compile the Angular application using the Ahead-of-Time (AOT) compiler

Application performance is one of the most important aspects of user experience. An Angular application can be optimized by using the Angular Ahead-of-Time (AOT) compiler to compile the app at build time. It converts all source code (HTML templates and TypeScript) into efficient JavaScript code. If you compile your app with the AOT compiler, no additional compilation will occur at runtime, which results in faster rendering and faster asynchronous requests for HTML templates. Additionally, the overall application size will be reduced, because the Angular compiler won't need to be included in the application distributable.

To use the AOT compiler, add `--aot` to the `ng build` or `ng serve` command:

```command&nbsp;line
ng build --aot
ng serve --aot
```

> [!NOTE]
> To learn more about the Angular Ahead-of-Time (AOT) compiler, see the [official guide](https://angular.io/guide/aot-compiler).

## Support Internet Explorer if you're dynamically loading Office.js

Based on the Windows version and the Office desktop client where your add-in is running, your add-in may be using Internet Explorer 11. (For more details, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).) Angular depends on a few `window.history` APIs but these APIs don't work in the IE runtime embedded in Windows desktop clients. When these APIs don't work, your add-in may not work properly, for example, it may load a blank task pane. To mitigate this, Office.js nullifies those APIs. However, if you're dynamically loading Office.js, AngularJS may load before Office.js. In that case, you should disable the `window.history` APIs by adding the following code to your add-in's **index.html** page.

```js
<script type="text/javascript">window.history.replaceState=null;window.history.pushState=null;</script>
```
