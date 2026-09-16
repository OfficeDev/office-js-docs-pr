The Office JavaScript API library can be accessed via the Office JS content delivery network (CDN) at: `https://officeapis.public.onecdn.static.microsoft/1/office.js`. To use Office JavaScript APIs within any of your add-in's web pages, you must reference the CDN in a `<script>` tag in the `<head>` tag of the page.

```html
<head>
    ...
    <script src="https://officeapis.public.onecdn.static.microsoft/1/office.js" type="text/javascript"></script>
</head>
```

> [!NOTE]
> To use preview APIs, reference the preview version of the Office JavaScript API library on the CDN: `https://officeapis.public.onecdn.static.microsoft/beta/office.js`.

For more information about accessing the Office JavaScript API library, including how to get IntelliSense, see [Referencing the Office JavaScript API library from its content delivery network (CDN)](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).
