
# ProjectViewTypes enumeration
Specifies the types of views that the  **[getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)** method can recognize.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Added in**|1.0|

```
ProjectViewTypes={
    Gantt           : 1, 
    NetworkDiagram  : 2, 
    TaskDiagram     : 3, 
    TaskForm        : 4, 
    TaskSheet       : 5, 
    ResourceForm    : 6, 
    ResourceSheet   : 7, 
    ResourceGraph   : 8, 
    TeamPlanner     : 9, 
    TaskDetails     : 10, 
    TaskNameForm    : 11, 
    ResourceNames   : 12, 
    Calendar        : 13, 
    TaskUsage       : 14, 
    ResourceUsage   : 15, 
    Timeline        : 16
}
```


## Members


****


|**Member**|**Description**|
|:-----|:-----|
|**Gantt**|The Gantt chart view.|
|**NetworkDiagram**|The Network Diagram view.|
|**TaskDiagram**|The Task Diagram view.|
|**TaskForm**|The Task form view.|
|**TaskSheet**|The Task Sheet view.|
|**ResourceForm**|The Resource Form view.|
|**ResourceSheet**|The Resource Sheet view.|
|**ResourceForm**|The Resource Form view.|
|**ResourceGraph**|The Resource Graph view.|
|**TeamPlanner**|The Team Planner view.|
|**TaskDetails**|The Task Details view.|
|**TaskNameForm**|The Task Name Form view.|
|**ResourceNames**|The Resource Names view.|
|**Calendar**|The Calendar view.|
|**TaskUsage**|The Task Usage view.|
|**ResourceUsage**|The Resource Usage view.|
|**Timeline**|The Timeline view.|

## Remarks

The  **[getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)** method returns the **ProjectViewTypes** constant value and name that corresponds to the active view.


## Support details


A capital Y in the following matrix indicates that this enumeration is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this enumeration.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.0|Introduced|

## See also



#### Other resources


[getSelectedViewAsync method](../../reference/shared/projectdocument.getselectedviewasync.md)
