# Tips for creating Office Add-ins with Angular 2.0 

Get some guidance for creating an Office Add-in as a single page application made with Angular 2.0.

>**Note:** We expect this topic to grow and we would love to have the benefit of your experiences creating Office Add-ins with Angular 2.0. If there's something we should add to this topic, please click the **Edit in GitHub** link above and either make a pull request or raise an issue in the repo. 

##See a sample Angular 2.0 Office Add-in

[Word-Add-in-Angular2-StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)

## Bootstrapping must be inside Office.initialize

On any page that calls the Office.js (or Word.js, Excel.js, etc.) APIs, your code must first assign a method to the `Office.initialize` property. (If you have no initialization code, the method body can be just empty "`{}`" symbols, but you must not leave the `Office.initialize` property undefined. For details see [Initializing your add-in](http://dev.office.com/docs/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in).) Office calls this method immediately after it has initialized the Office JavaScript libraries.

**Your Angular 2.0 bootstrapping code must be called inside the method that you assign to `Office.initialize`** to ensure that the Office JavaScript libraries have initialized first. The following is a simple example of how to do this. This code should be in the main.ts file of the project.

```js
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
    import { AppModule } from './app.module';
	Office.initialize = function () {
    	const platform = platformBrowserDynamic();
        platform.bootstrapModule(AppModule);
  };
```

##Use the Hash Location Strategy in the Angular 2.0 application

Navigating between routes in the application may not work properly if you do not specify the Hash Location Strategy. There are two ways to do this. First, you can specify a provider for the location strategy in your app module as the following code does. It goes into the app.module.ts file.

```js
import { LocationStrategy, HashLocationStrategy } from '@angular/common';
// Other imports suppressed for brevity
    @NgModule({
        providers: [
            {provide: LocationStrategy, useClass: HashLocationStrategy},
            // Other providers suppressed
        ],
        // Other module properties suppressed
  })
  export class AppModule {}
``` 

If you define your routes in a separate routing module, then there is an alternative way to specify the hash location strategy. In your routing module's *.ts file, pass a configuration object to the `forRoot` function that specifies the strategy. The following code is an example. 

```js
import { RouterModule, Routes } from '@angular/router';
// Other imports suppressed for brevity
    const routes: Routes = // route definitions go here
    @NgModule({
      imports: [ RouterModule.forRoot(routes, {useHash: true}) ],
      exports: [ RouterModule ]
    })
    export class AppRoutingModule {}
```   


##Consider wrapping Fabric components with Angular 2.0 components

We recommend using [Office Fabric](http://dev.office.com/fabric#/fabric-js) styling in your add-in. Fabric includes components that come in several versions including a version [based on TypeScript](https://github.com/OfficeDev/office-ui-fabric-js). Consider using Fabric components in your add-in by wrapping them in Angular 2.0 components. See [Word-Add-in-Angular2-StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker) for an example of an add-in that does this. Note, for example, how the Angular component defined in [fabric.textfield.wrapper](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/shared/office-fabric-component-wrappers/fabric.textfield.wrapper.component.ts) imports the Fabric file TextField.ts, where the Fabric component is defined. 


## Using the Office Dialog API with Angular 2.0

The Office add-in Dialog API enables your add-in to open a page in a semi-modal dialog which can exchange information with the main page, which is typically in a task pane. 

The Dialog API's [displayDialogAsync](http://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) method takes a parameter that specifies the URL of the page that should open in the dialog. Your add-in can have complete and separate HTML page to pass to this parameter, or you can pass the URL of a route in your Angular 2.0 appication. 

It is important to remember, if you pass a route, that the dialog creates an entirely new window with it's own execution context. Your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog window. So this technique launches a second instance of your single page application in the dialog window. Code that changes variables in the dialog window does not change the task pane version of the same variables. Similarly, the dialog window has its own session storage, which is not accessible from code in the task pane.  


## Forcing an update of the DOM

In any Angular 2.0 application, whether or not it is an Office Add-in, occasionally notifications to update the DOM do not fire. The framework provides a `tick()` method on the `ApplicationRef` object that will force an update. The following code is an example.
```js
import { ApplicationRef } from '@angular/core';
    export class MyComponent {
        constructor(private appRef: ApplicationRef) {}
        myMethod() {
            // Code that changes the DOM is here
            appRef.tick();
        }
}
``` 

## Using Observables

Angular 2.0 uses RxJS (Reactive Extensions for JavaScript), and RxJS introduces `Observable` and `Observer` objects to implement asynchronous processing. We can't provide a complete discussion, but this section provides a brief introduction before you turn to the official documentation at [RxJS](http://reactivex.io/rxjs/).

An `Observable` is like a `Promise` object in some ways: it is returned immediately from an asynchronous call, but it may not resolve until some time later. However, while a `Promise` is a single value (which may be an array object), an `Observable` is an array of objects (possibly with only a single member). This enables code to call [array methods](http://www.w3schools.com/jsref/jsref_obj_array.asp), such as `concat`, `map`, and `filter` on `Observable` objects. 

### Pushing instead of pulling

Another difference is that your code "pulls" `Promise` objects by assigning them to variables, but `Observable` objects "push" their values to objects that *subscribe* to the `Observable`. The subscribers are `Observer` objects. The benefit of the push architecture is that new members can be added to the `Observable` array over time. Whenever a new member is added, all the `Observer` objects that subscribe to the `Observable` receive a notification. 

The `Observer` is configured to process each new object (called the "next" object) with a function. (It is also configured to respond to an error and a completion notification. See below for an example.) For this reason, `Observable` objects can be easily used in a wider range of scenarios that `Promise` objects. For example, in addition to returning an `Observable` from an AJAX call, as a `Promise` can be, an `Observable` can be returned from an event handler, such as the "changed" event handler for a text box. Each time a user enters text in the box, all the subscribed `Observer` objects react immediately using the latest text and/or the current state of the application as input. 


### Waiting until all asynchronous calls have completed

When you want to ensure that a callback only runs when every member of a set of `Promise` objects has resolved, you would use the `Promise.all()` method:

```js
myPromise.all([x, y, z]).then(// ToDo: callback logic goes here)
``` 

To do the same thing with an `Observable` object, you use the [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md) method:  

```js
var source = Rx.Observable.forkJoin([x, y, z]);

var subscription = source.subscribe(
  function (x) {
    // ToDo: callback logic goes here
  },
  function (err) {
    console.log('Error: ' + err);
  },
  function () {
    console.log('Completed');
  });
``` 

