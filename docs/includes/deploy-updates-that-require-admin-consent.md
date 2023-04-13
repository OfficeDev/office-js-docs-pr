When you add features or fix bugs in your add-in, you'll need to deploy the updates. If your add-in is deployed by one or more admins to their organizations, some manifest changes will require the admin to consent to the updates. Users will be blocked from the add-in until consent is granted. The following manifest changes will require the admin to consent again.

- Changes to requested [permissions](/javascript/api/manifest/permissions).
- Additional [scopes](/javascript/api/manifest/scopes).
- Additional [Outlook events](../outlook/autolaunch.md).
