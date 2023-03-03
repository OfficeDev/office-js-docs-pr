---
title: Guidance for deploying Office Add-ins on government clouds
description: Learn how to deploy your Office Add-in to secure, government cloud environments
ms.date: 10/17/2022
ms.localizationpriority: medium
---

# Guidance for deploying Office Add-ins on government clouds

Microsoft provides the government cloud options for our privacy-sensitive customers in local, state, and national government organizations. This gives partners opportunities to target critical customers with their Office Add-ins. Due to the more restricted nature of these environments, which is important for the customers’ privacy and security needs, not all resources that are typically available in a standard production environment are available within these clouds.

For partners providing their Office Add-ins to customers in these restricted cloud environments, there are important differences from the standard public cloud environment that must be considered. The following information gives the details that require special handling by developers writing Office Add-ins that target customers in these environments.

## All sovereign environments

For all government cloud (i.e. Sovereign Cloud) environments, the public Office Store is not available. This means that end-users cannot acquire Office Add-ins directly from the public store. Administrators are also unable to deploy Office Add-ins directly from the public store into their Admin Portal. Instead, you must work with administrators to ensure the following:

- The required resources and services for your solution are available inside the cloud boundary. Either you work with the tenant administrators to provision your service and resources inside of the cloud boundary, or you work with the network administrator to enable access to your resources that reside outside of the cloud boundary.

- The resources the Office Add-in accesses conform to the requirements of the government cloud from which they are being accessed. They also must conform to any additional requirements from the customer tenant for which the solution is being provisioned. These requirements include the transfer, management, and storage of potentially sensitive data, as well as code and resource security and access vetting procedures.

- The Office Add-in manifest that describes the solution and its source location as applicable to the particular government cloud deployment is obtained from the partner and ingested for deployment to the appropriate users via the Admin Portal.

## US Government Community Cloud (GCC)

In addition to requirements applicable to all Sovereign Clouds, each Sovereign Cloud environment has its own considerations that may affect Office Add-ins targeting the environment. GCC is the least restrictive of the government cloud environments and some of the resources required by the add-in are available from their usual production endpoints in this environment. One such resource is the Office JavaScript API library. Solution partners can continue to reference the public Office.js resource as they do with their public production solution.

## GCC High (GCCH), US Department of Defense (DOD), or other sovereign managed clouds

These government clouds remain internet-connected, but the set of public endpoints made available is severely restricted. One such restricted endpoint is the public endpoint for loading the Office JavaScript API library. The public CDN location for Office.js will not be accessible from within these environments. However, there will be an internal, per-cloud Microsoft Office CDN provisioned with the same resource. This means the endpoint URL to access Office.js will be different and your Office Add-in may need some level of customization to run. Given the additional restrictions, it's likely that any solution provided to customers will require hosting provider services within the environment as well, necessitating additional customizations. You'll need to determine the best way to provide your solution to customers, such that it conforms to the additional restrictions imposed on services running within the boundaries of these environments.

## Airgapped Sovereign Clouds

These government clouds are essentially disconnected from the public internet entirely. Any resource that would normally be accessed from public resources must instead be custom-provisioned inside these cloud environments. In the GCCH and DOD clouds mentioned previously, most (if not all) solution providers will need to provision their services and resources inside the cloud. There is an option to make firewall exceptions that allows access to public services and resources. However, this bypass is not possible in airgapped clouds. As with the GCCH and DOD clouds, there will be a Microsoft Office CDN provisioned inside each cloud environment that hosts required resources such as the Office.js library. You'll need to work closely with customer tenant administrators to determine the best way to provide your services and resources in a way that conforms to the strict access requirements for airgapped Sovereign Clouds.

## Office 365 operated by 21Vianet

[!INCLUDE [Information about the China-specific CDN](../includes/21Vianet-CDN.md)]
