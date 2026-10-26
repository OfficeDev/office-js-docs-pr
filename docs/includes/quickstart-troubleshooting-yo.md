[!include[The common troubleshooting section for all quick starts](../includes/quickstart-troubleshooting-common.md)]

- The automatic `npm install` step Yo Office performs may fail. If you see errors when trying to run `npm start`, navigate to the newly created project folder in a command prompt and manually run `npm install`. For more information about Yo Office, see [Create Office Add-in projects using the Yeoman Generator](../develop/yeoman-generator-overview.md).
- You may see warnings generated when running `npm install` for either Yeoman generator or the project. In most cases, you can safely ignore these warnings. Security or audit warnings require separate assessment; address vulnerabilities in runtime dependencies, or dependencies whose origin you can't determine, before deploying to production. Sometimes, dependencies become deprecated and their replacements aren't supported by other packages that the project depends on. To review available dependency updates, use the `npm-check-updates` tool.
  - In the command prompt while in the root project directory, run `npm i -g npm-check-updates`. This installs the tool globally.
  - Run `ncu` to report available updates without modifying **package.json**.
  - Review the proposed versions for compatibility and update only the appropriate dependency ranges in **package.json**, then run `npm install`. Note that `ncu -u` rewrites dependency ranges and can include breaking major-version upgrades.

   For more information about warnings when running `npm install`, see [Warnings and dependencies in the Node.js and npm world](../overview/npm-warnings-advice.md).

- If you don't see the add-in button on the ribbon (e.g., **Show Task Pane**), select the **Add-ins** button from the **Home** ribbon tab, then choose the test add-in.