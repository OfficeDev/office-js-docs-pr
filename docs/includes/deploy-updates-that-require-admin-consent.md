When you add features or fix bugs in your add-in, you'll need to deploy the updates. If your add-in is deployed by one or more admins to their organizations, some manifest changes will require the admin to consent to the updates. Users remain on the existing version of the add-in until the admin consents to the updates. The following manifest changes will require the admin to consent again.

- Changes to requested permissions. See [Requesting permissions for API use in add-ins](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) and [Understanding Outlook add-in permissions](../outlook/understanding-outlook-add-in-permissions.md).
- Additional or changed [Scopes](/javascript/api/manifest/scopes). (Not applicable if the add-in uses the unified manifest for Microsoft 365.)
- Additional or changed [Events](../develop/event-based-activation.md).

> [!NOTE]
> Whenever you make a change to the manifest, you must raise the version number of the manifest.
>
> - If the add-in uses the add-in only manifest, see [Version element](/javascript/api/manifest/version).
> - If the add-in uses the unified manifest, see [version property](/microsoft-365/extensibility/schema/root#version).