## Deploy updates

When you add features or fix bugs in your add-in you'll need to deploy the updates. If your add-in is deployed by one or more admins to their organizations, some manifest changes will require the admin to consent to the updates. Users will be blocked from the add-in until consent is granted. We recommend you test deployment behavior on a staging site before deploying updates live. The following manifest changes will require the admin to consent again.

- Changes to requested [permissions](/javascript/api/manifest/permissions).
- Adding new [scopes](/javascript/api/manifest/scopes).
- Adding new [Outlook events](../outlook/autolaunch.md).
