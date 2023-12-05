## Troubleshooting

Ensure your environment is ready for Office development by following the instructions in [Set up your development environment](../overview/set-up-your-dev-environment.md).

The automatic `npm install` step yo office performs may fail. If you see errors when trying to run `npm start`, navigate to the newly created project folder in a command prompt and manually run `npm install`. For more information about yo office, see [Create Office Add-in projects using the Yeoman Generator](../develop/yeoman-generator-overview.md).

Some of the sample code uses ES6 JavaScript. This isn't compatible with [older versions of Office that use the Trident (Internet Explorer 11) browser engine](/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins). For information on how to support those platforms in your add-in, see [Support older Microsoft webviews and Office versions](/office/dev/add-ins/develop/support-ie-11). [Join the Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program) to get a free, 90-day renewable Microsoft 365 subscription, with the latest Office applications, to use during development.
