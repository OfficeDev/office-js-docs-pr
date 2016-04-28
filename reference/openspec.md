# On-going API design specification

thank you for your interest in learning more about new API/features under design. This is a froum through which we will make the early versions of the API specification available for community feedback. Please go thorugh the available topics below and provide your feedback. We value your inputs and will go a long way in ensuring that final design meets the critical use-cases that developers are targetting. 

_**Note:** below listed features are still under design and review phase and hence not yet available as part of the generally available version. The final shape of the API or feature is subject to change based on feasibility, validation, feedback, etc. The final specification will be published once the feature is made available ._


## New Word JavaScript Add-in APIs:
The WordAPI1.3 JavaScript API update contains the largest set of changes since this API was introduced. You’ll now be able to: create and alter documents in memory, create and access list objects, create and access table objects, and more options for accessing and comparing range objects.
These changes have been implemented across nearly all of the WordJS objects. This functionality is now available, or will be shortly, on Word 2016 on the desktop for both Windows and Mac, and on the iPad. So update your clients to the latest monthly build and start implementing these great features!

**Visit this [page](https://github.com/OfficeDev/office-js-docs/tree/WordJs_1.3_Openspec/word) to learn more and provide your feedback.**

## Document properties access
We are working on adding the ability for web Add-ins to access (get, set) document level properties. This feature will allow Add-ins to use document properties to be integrated as part of custom workflows or to simply read/set the document properties for reference purpose. Applications that will support the feature include: Excel, Word, Potentially PowerPoint. The feature will also work for Excel REST API (as the Excel supports REST service). We’ll introduce the basic design idea and work through the use-cases and code snippet of how the API would work when they are added. We welcome any design feedback you have to offer. 

**Visit this [page](https://github.com/OfficeDev/office-js-docs/tree/DocumentProperties_OpenSpec) to learn more and provide your feedback.**



