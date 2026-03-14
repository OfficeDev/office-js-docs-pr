---
title: Guidance for deploying Office Add-ins on government clouds
description: Learn how to deploy your Office Add-in to secure, government cloud environments
ms.topic: best-practice
ms.date: 03/06/2026
ms.localizationpriority: medium
---

# Guidance for deploying Office Add-ins on government clouds

Microsoft provides the government cloud options for our privacy-sensitive customers in local, state, and national government organizations. This gives partners opportunities to target critical customers with their Office Add-ins. Due to the more restricted nature of these environments, which is important for the customers’ privacy and security needs, not all resources that are typically available in a standard production environment are available within these clouds.

For partners providing Office Add-ins to customers in these restricted cloud environments, consider the important differences from the standard public cloud environment. The following information describes details that require special handling by developers writing Office Add-ins that target customers in these environments.

## All sovereign environments

For all government cloud (sovereign cloud) environments, the public Microsoft 365 and Copilot store isn't available. This limitation means that end users can't acquire Office Add-ins directly from the public store. Administrators also can't deploy Office Add-ins directly from the public store into their Admin Portal. Instead, work with administrators to ensure the following conditions:

- The required resources and services for your solution are available inside the cloud boundary. Either you work with the tenant administrators to provision your service and resources inside of the cloud boundary, or you work with the network administrator to enable access to your resources that reside outside of the cloud boundary.

- The resources the Office Add-in accesses conform to the requirements of the government cloud from which they're being accessed. They also must conform to any additional requirements from the customer tenant for which the solution is being provisioned. These requirements include the transfer, management, and storage of potentially sensitive data, as well as code and resource security and access vetting procedures.

- The Office Add-in manifest that describes the solution and its source location as applicable to the particular government cloud deployment is obtained from the partner and ingested for deployment to the appropriate users via the Admin Portal.

[Centralized deployment](/microsoft-365/admin/manage/centralized-deployment-of-add-ins) of add-ins outside of the store is still supported.

In addition to requirements applicable to all sovereign clouds, each sovereign cloud environment has its own considerations that might affect Office Add-ins targeting the environment. The following sections describe these requirements and recommendations.

### US Government Community Cloud (GCC)

GCC is the least restrictive of the government cloud environments. Solution partners are permitted to reference the public Office JavaScript API library (office.js) resource as they do with their public production solution. However, we recommend that partners reference the library from the following URL.

- **GCC**: `https://appsforoffice.gcc.cdn.office.net/appsforoffice/lib/1/hosted/office.js`

### GCC High (GCCH), US Department of Defense (DOD), or other sovereign managed clouds

These government clouds remain internet-connected, but they severely restrict the set of public endpoints your code can access. One such restricted endpoint is the public endpoint for loading the Office JavaScript API library. Your code can't access the public CDN location for Office.js from within these environments. However, there's an internal, per-cloud Microsoft Office CDN provisioned with the same resource. This means the endpoint URL to access Office.js is different and your Office Add-in might need some level of customization to run. Given the additional restrictions, it's likely that any solution you provide to customers requires hosting provider services within the environment as well, necessitating additional customizations. You need to determine the best way to provide your solution to customers so that it conforms to the additional restrictions imposed on services running within the boundaries of these environments. The Office JavaScript Library CDN URLs are as follows:

- **GCC High**: `https://appsforoffice.gcch.cdn.office.net/appsforoffice/lib/1/hosted/office.js`
- **DOD**: `https://appsforoffice.dod.cdn.office.net/appsforoffice/lib/1/hosted/office.js`

### Air-gapped Sovereign Clouds

These government clouds are essentially disconnected from the public internet entirely. You must custom-provision any resource that you'd normally access from public resources inside these cloud environments. In the GCCH and DOD clouds mentioned previously, you need to provision your services and resources inside the cloud. You can make firewall exceptions that allow access to public services and resources. However, this bypass isn't possible in air-gapped clouds. As with the GCCH and DOD clouds, there's a Microsoft Office CDN inside each cloud environment that hosts required resources such as the Office.js library. You need to work closely with customer tenant administrators to determine the best way to provide your services and resources in a way that conforms to the strict access requirements for air-gapped Sovereign Clouds.

### Office 365 operated by 21Vianet

[!INCLUDE [Information about the China-specific CDN](../includes/21Vianet-CDN.md)]
