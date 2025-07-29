[!include[The common troubleshooting section for all quick starts](../includes/quickstart-troubleshooting-common.md)]

- The automatic `npm install` step Yo Office performs may fail. If you see errors when trying to run `npm start`, navigate to the newly created project folder in a command prompt and manually run `npm install`. For more information about Yo Office, see [Create Office Add-in projects using the Yeoman Generator](../develop/yeoman-generator-overview.md).
- You may see warnings generated when running `npm install` for either Yeoman generator or the project. In most cases, you can safely ignore these warnings. Sometimes, dependencies become deprecated and their replacements aren't supported by other packages on which the project depends. If you would like to resolve these warnings, use the `npm-check-updates` tool.
  - In the command prompt while in the root project directory, run `npm i -g npm-check-updates`. This installs the tool globally.
  - Run `ncu -u`. This provides a report of all packages and to what versions they will be updated.
  - Run `npm install` to update all the packages.
  
   For more information about warnings when running `npm install`, see [Warnings and dependencies in the Node.js and npm world](../overview/npm-warnings-advice.md).