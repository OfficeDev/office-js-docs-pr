---
title: Connect Office.js to any JavaScript framework
description: Learn how to integrate Office.js with any JavaScript framework including React, Angular, Vue, Svelte, and others.
ms.date: 01/30/2026
ms.topic: best-practice
ms.localizationpriority: medium
---

# Connect Office.js to any JavaScript framework

Office.js is framework-agnostic and works seamlessly with any client-side JavaScript framework or library. Whether you're building with React, Angular, Vue, Svelte, or any other framework, the integration pattern is the same: ensure Office.js initializes before your application renders.

> [!NOTE]
> You can also use server-side frameworks such as ASP.NET, PHP, and Java to build Office Add-ins, but this article doesn't cover them. This article focuses specifically on client-side JavaScript frameworks that run in the browser.

This article explains the universal patterns for integrating Office.js with client-side JavaScript frameworks, important considerations, and provides examples across multiple frameworks.

> [!TIP]
> This article is designed for developers creating Office Add-ins from scratch using their preferred JavaScript framework, or integrating Office.js into an existing framework project. If you're using the [Yeoman generator for Office Add-ins](yeoman-generator-overview.md) or [Microsoft 365 Agents Toolkit](../develop/teams-toolkit-overview.md), these tools already provide the correct Office.js configuration.

## Prerequisites

- Familiarity with your chosen JavaScript framework.
- Basic understanding of Office Add-ins. See [Office Add-ins platform overview](../overview/office-add-ins.md).
- Node.js installed for package management.

## Quick start: The universal pattern

Regardless of which framework you choose, use the following pattern.

1. Reference Office.js from the CDN in your HTML `<head>`.
1. Call `Office.onReady()` and wait for it to complete.
1. Initialize your framework after Office.js is ready.

```typescript
// Universal pattern - works with any framework.
Office.onReady((info) => {
  // Office.js is now ready.
  // Initialize your framework here.
  initializeYourFramework();
});
```

## Load Office.js from the CDN

You must reference the Office JavaScript API from the content delivery network (CDN) in your HTML file. Add the following `<script>` tag in the `<head>` section of your HTML page, before any other script tags or framework bundle references.

```html
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>My Office Add-in</title>

  <!-- Office.js must be loaded from CDN, not bundled -->
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>

  <!-- Your framework bundle loads after Office.js -->
</head>
```

> [!IMPORTANT]
>
> - Load Office.js from the CDN and reference it in your HTML file. Don't import it in your JavaScript or TypeScript code.
> - The Office.js reference must appear in the `<head>` section to ensure the API is fully initialized before any body elements load.
> - Don't bundle Office.js with your application code. Always reference it from the CDN.

For more information about referencing Office.js, including preview APIs and alternative CDN endpoints, see [Referencing the Office JavaScript API library](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## Initialize your framework after Office.onReady

The key to integrating Office.js with any framework is to initialize your application inside the `Office.onReady()` callback. This approach ensures Office.js is fully initialized before your framework starts rendering. This initialization is important because Office.js needs to:

- Download and cache API library files from the CDN.
- Initialize the Office runtime environment.
- Establish communication with the Office application.

If your framework renders before Office.js is ready, calls to Office APIs fail. By initializing your application inside `Office.onReady()`, you guarantee Office.js is ready when your application code runs.

### Examples

The following examples show the same integration pattern across different frameworks. The pattern is identical - only the framework's initialization method changes.

#### React

```typescript
// src/index.tsx
Office.onReady(() => {
  const root = ReactDOM.createRoot(document.getElementById('root'));
  root.render(<App />);
});
```

#### Angular

```typescript
// src/main.ts
Office.onReady(() => {
  platformBrowserDynamic()
    .bootstrapModule(AppModule)
    .catch(err => console.error(err));
});
```

#### Vue

```typescript
// src/main.ts
Office.onReady(() => {
  createApp(App).mount('#app');
});
```

#### Svelte

```typescript
// src/main.ts
Office.onReady(() => {
  new App({ target: document.getElementById('app') });
});
```

#### Simple JavaScript with no framework

```javascript
// src/app.js
Office.onReady((info) => {
  document.getElementById('run-button').onclick = run;

  if (info.host === Office.HostType.Excel) {
    console.log('Running in Excel');
  }
});
```

## Use Office.js APIs in your application

After Office.js initializes (when `Office.onReady()` finishes), you can call Office APIs anywhere in your add-in. Use your framework's lifecycle hooks or event handlers to call Office APIs when needed.

```typescript
// React example: Call an Office JS API in the useEffect lifecycle hook.
import { useEffect, useState } from 'react';

function MyComponent() {
  const [data, setData] = useState('');

  useEffect(() => {
    loadData();
  }, []);

  async function loadData() {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load('values');
      await context.sync();

      // Update component state with the data from Excel.
      const value = range.values[0][0];
      setData(value);
    });
  }

  return <div>Selected cell: {data}</div>;
}

// Similar patterns for other frameworks:
// Angular: ngOnInit() { this.loadData(); }
// Vue: onMounted(() => { loadData(); })
// Svelte: onMount(() => { loadData(); })
```

## TypeScript support

To enable IntelliSense and type checking for Office.js in TypeScript projects, install the type definitions from DefinitelyTyped.

```bash
npm install --save-dev @types/office-js
```

TypeScript automatically recognizes the types. You don't need an import statement in your code because Office.js is loaded globally from the CDN.

```typescript
// TypeScript automatically recognizes Office types.
Office.onReady((info: Office.OfficeInfo) => {
  if (info.host === Office.HostType.Excel) {
    // TypeScript provides IntelliSense for Excel APIs.
  }
});
```

For more information, see [Referencing the Office JavaScript API library](referencing-the-javascript-api-for-office-library-from-its-cdn.md#enabling-intellisense-for-a-typescript-project).

## Other considerations

### Loading indicators

If you want to show a loading indicator while Office.js initializes, display it before calling `Office.onReady()` and hide it inside the callback.

```typescript
// Show loading indicator.
document.getElementById('loading')!.style.display = 'block';

Office.onReady((info) => {
  // Hide loading indicator.
  document.getElementById('loading')!.style.display = 'none';

  // Initialize framework.
  initializeYourFramework();
});
```

For a better user experience with frameworks that have their own loading states, use a simple HTML/CSS loader that displays immediately. Then, let your framework take over once it's mounted.

### Dialog API and component lifecycle

The [Office Dialog API](dialog-api-in-office-add-ins.md) opens pages in separate browser windows. This behavior has important implications for framework applications:

- Each dialog creates a **new execution context** with a separate framework instance.
- The dialog runs its own copy of your application code.
- You must call `Office.onReady()` in the dialog page.
- The main page and dialog windows **don't share state**.
- Session storage **isn't shared** between contexts.

If you use a framework router to navigate to a dialog route, remember that the dialog window creates a completely new instance of your application. It doesn't reuse the existing instance.

```typescript
// Main page - opens a dialog.
Office.context.ui.displayDialogAsync(
  'https://localhost:3000/dialog-route',
  { height: 50, width: 50 },
  (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        // Handle message from dialog.
      });
    } else {
      // Handle error opening the dialog.
      console.error(result.error);
    }
  }
);

// Dialog page - must also call Office.onReady.
Office.onReady(() => {
  // This is a separate framework instance.
  initializeYourFramework();
});
```

### History API workaround

Office.js replaces the default [Window.history](https://developer.mozilla.org/docs/Web/API/History) methods `replaceState` and `pushState` with `null`. If your framework or router depends on these methods (common in React Router, Vue Router, Angular Router, and others), you need to cache and restore them.

Add this code to your HTML file, wrapping the Office.js script tag:

```html
<head>
  <!-- Cache history methods before Office.js loads -->
  <script type="text/javascript">
    window._historyCache = {
      replaceState: window.history.replaceState,
      pushState: window.history.pushState
    };
  </script>

  <!-- Load Office.js -->
  <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>

  <!-- Restore history methods after Office.js loads -->
  <script type="text/javascript">
    window.history.replaceState = window._historyCache.replaceState;
    window.history.pushState = window._historyCache.pushState;
  </script>
</head>
```

> [!NOTE]
> This workaround is only necessary if your application uses client-side routing (React Router, Vue Router, Angular Router, and others). Static applications without routing don't need this workaround.

### Testing outside Office applications

You can develop and test your add-in's UI by using browser developer tools without sideloading into Office. This approach enables faster iteration during development and makes it easier to debug your UI components.

When you open your add-in in a regular browser (outside of an Office application), `Office.onReady()` still executes, but it resolves with `null` for both the host and platform properties.

```typescript
Office.onReady((info) => {
  if (info?.host) {
    console.log(`Running in ${info.host} on ${info.platform}`);
  } else {
    console.log('Running outside of Office (development mode)');
  }

  // Initialize your framework, regardless of whether the add-in is running inside or outside of Office.
  initializeYourFramework();
});
```

### Build tools and bundlers

Modern JavaScript frameworks typically use build tools like Webpack, Vite, Rollup, or esbuild. When configuring your build:

- **Don't import or bundle Office.js** in your JavaScript or TypeScript code.
- Load Office.js from the CDN by using a `<script>` tag in your HTML.
- Configure your bundler to treat `Office` as a global variable.

#### Example: TypeScript configuration with Vite

If you use Vite with TypeScript, you typically don't need special Vite configuration for Office.js. The `@types/office-js` package provides the necessary type definitions. However, if you need to ensure the Office.js types are available, verify your `tsconfig.json`:

```json
// tsconfig.json
{
  "compilerOptions": {
    "types": ["office-js"]
    // ... your other compiler options ...
  }
}
```

#### Example: Webpack configuration

```javascript
// webpack.config.js
module.exports = {
  externals: {
    'office': 'Office'
  }
};
```

Add-in projects generated by the [Yeoman generator for Office Add-ins](yeoman-generator-overview.md) include the correct build configuration by default.

### Network blocking and firewalls

If network filters, firewalls, or browser extensions block the Office.js CDN, `Office.onReady()` never resolves. Consider implementing a timeout for enterprise scenarios where network policies might block the CDN.

```typescript
let officeInitialized = false;

// Set a timeout.
setTimeout(() => {
  if (!officeInitialized) {
    console.error('Office.js failed to initialize. Network may be blocking CDN.');
    // Show error message to user.
  }
}, 10000); // 10 second timeout

Office.onReady((info) => {
  officeInitialized = true;
  initializeYourFramework();
});
```

For more information about CDN considerations, see [Referencing the Office JavaScript API library](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

### Framework-specific zone or reactivity problems

Some frameworks use zones or reactivity systems to track state changes. In rare cases, Office API calls don't trigger UI updates because they run outside the framework's change detection zone.

**Angular:** If the UI doesn't update after Office API calls, wrap your code in `NgZone.run()`:

```typescript
import { NgZone } from '@angular/core';

constructor(private zone: NgZone) {}

async loadDataFromExcel() {
  let cellValue: string;

  // Make Office API call
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load('values');
    await context.sync();
    cellValue = range.values[0][0];
  });

  // Update Angular component state inside zone
  this.zone.run(() => {
    this.myData = cellValue;
  });
}
```

## See also

- [Initialize your Office Add-in](initialize-add-in.md)
- [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md)
- [Referencing the Office JavaScript API library](referencing-the-javascript-api-for-office-library-from-its-cdn.md)
- [Office Dialog API](dialog-api-in-office-add-ins.md)
- [Debug the initialize and onReady functions](../testing/debug-initialize-onready.md)
- [Understanding the Office JavaScript API](understanding-the-javascript-api-for-office.md)
