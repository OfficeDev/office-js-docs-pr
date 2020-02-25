The Office JavaScript library can be accessed via the Office JS content delivery network (CDN) at: `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`. To use Office JavaScript APIs within any of your add-in's web pages, you must reference the CDN in a `<script>` tag in the `<head>` tag of the page.

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
</head>
```

> [!NOTE]
> To use preview APIs, reference the preview version of the Office JavaScript library on the CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.

For more information about accessing the Office JavaScript library, including how to get IntelliSense, see [Referencing the JavaScript API for Office library from its content delivery network (CDN)](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).