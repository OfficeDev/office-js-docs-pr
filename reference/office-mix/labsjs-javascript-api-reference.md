
# LabsJS JavaScript API reference
Get an overview of the LabsJS JavaScript object model.

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

The LabsJS reference documents the [TypeScript](http://www.typescriptlang.org/) file, **labs-1.0.42.d.ts**, which sorts the LabsJS object model into modules.

## LabsJS object model

The LabsJS object model is organized into five modules:


- [LabsJS.Labs](../../reference/office-mix/labsjs.labs.md). The Labs module contains the set of key APIs with which to create the labs themselves. They provide the entry point for lab development.
    
- [LabsJS.Labs.Core](../../reference/office-mix/labsjs.labs.core.md). The core interfaces, data structures, and classes that are shared by the LabsJS and the presentation driver (in this case, Office Mix), to create a bridge between the two.
    
- [LabsJS.Labs.Core.Actions](../../reference/office-mix/labsjs.labs.core.actions.md). These APIs represent the operations that a lab, indicating its current behaviors, and are useful to developers who are creating new components (other than the default components), or developing connections with a new driver (other than Office Mix).
    
- [LabsJS.Labs.Core.GetActions](../../reference/office-mix/labsjs.labs.core.getactions.md). These APIs allow you to query for actions that have occurred previously on the server.
    
- [LabsJS.Labs.Components](../../reference/office-mix/labsjs.labs.components.md). These APIs represent the four default components that are available presently available to labs (Activity, Choice, Input, and Dynamic).
    
Each module contains a set of members comprised of one or more of the following member types:


- Classes
    
- Interfaces
    
- Functions
    
- Enumerations
    
- Variables
    



## Additional resources



- [TypeScript](http://www.typescriptlang.org/)
    
