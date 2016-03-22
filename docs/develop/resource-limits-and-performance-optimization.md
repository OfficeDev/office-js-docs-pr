
# Resource limits and performance optimization for Office Add-ins



To create the best experience for your users, ensure that your Office Add-in performs within specific limits for CPU core and memory usage, reliability and, for Outlook add-ins, the response time for evaluating regular expressions. These run-time resource usage limits apply to add-ins running on Office clients for Windows and OS X, but not Office Online, Outlook Web App,or OWA for Devices. You can also optimize the performance of your add-ins on desktop and mobile devices by optimizing the use of resources in your add-in design and implementation.

## Resource usage limits for add-ins


Run-time resource usage limits apply to all types of Office Add-ins. These limits help ensure performance for your users and mitigate denial-of-service attacks. Be sure to test your Office Add-in on your target host application using range of possible data and measure its performance against the following run-time usage limits:


-  **CPU core usage** - A single CPU core usage threshold of 90%, observed three times in default 5-second intervals.
    
    The default interval for a host rich client to check CPU core usage is every 5 seconds. If the host client detects the CPU core usage of an add-in is above the threshold value, it displays a message asking if the user wants to continue running the add-in. If the user chooses to continue, the host client does not ask the user again during that edit session. Administrators might want to use the  **AlertInterval** registry key to raise the threshold to reduce the display of this warning message if users run CPU-intensive add-ins.
    
-  **Memory usage** - A default memory usage threshold that is dynamically determined based on the available physical memory of the device.
    
    By default, when a host rich client detects that physical memory usage on a device exceeds 80% of the available memory, the client starts monitoring the add-in's memory usage, at a document level for content and task pane add-ins, and at a mailbox level for Outlook add-ins. At a default interval of 5 seconds, the client warns the user if physical memory usage for a set of add-ins at the document or mailbox level exceeds 50%. This memory usage limit uses physical rather than virtual memory to ensure performance on devices with limited RAM, such as tablets. Administrators can override this dynamic setting with an explicit limit by using the  **MemoryAlertThreshold** Windows registry key as a global setting, ir adjust the alert interval by using the **AlertInterval** key as a global setting.
    
-  **Crash tolerance** - A default limit of four crashes for an add-in.
    
    Administrators can adjust the threshold for crashes by using the **RestartManagerRetryLimit** registry key.
    
-  **Application blocking** - Prolonged unresponsiveness threshold of 5 seconds for an add-in.
    
    This affects the user's experiences of the add-in and the host application. When this occurs, the host application automatically restarts all the active add-ins for a document or mailbox (where applicable), and warns the user as to which add-in became unresponsive. Add-ins can reach this threshold when they do not regularly yield processing while performing long-running tasks. There are techniques to ensure that blocking does not occur. Administrators cannot override this threshold.
    
     **Outlook add-ins**
    
    If any Outlook add-in exceeds the preceding thresholds for CPU core or memory usage, or tolerance limit for crashes, Outlook disables the add-in. The Exchange Admin Center displays the disabled status of the app.
    
     >**Note**  Even though only the Outlook rich clients and not Outlook Web App or OWA for Devices monitor resource usage, if a rich client disables an Outlook add-in, that add-in is also disabled for use in Outlook Web App and OWA for Devices.

    In addition to the CPU core, memory, and reliability rules, Outlook add-ins should observe the following rules on activation:
    
      -  **Regular expressions response time** - A default threshold of 1,000 milliseconds for Outlook to evaluate all regular expressions in the manifest of an Outlook add-in. Exceeding the threshold causes Outlook to retry evaluation at a later time.
    
        Using a group policy or application-specific setting in the Windows registry, administrators can adjust this default threshold value of 1,000 milliseconds in the  **OutlookActivationAlertThreshold** setting. For more information, see [Overriding resource usage settings for performance of Office Add-ins](http://msdn.microsoft.com/library/da14ec8c-5075-4035-a951-fc3c2b15c04b%28Office.15%29.aspx).
    
  -  **Regular expressions re-evaluation** - A default limit of three times for Outlook to reevaluate all the regular expressions in a manifest. If evaluation fails all three times by exceeding the applicable threshold (which is either the default of 1,000 milliseconds or a value specified by  **OutlookActivationAlertThreshold**, if that setting exists in the Windows registry), Outlook disables the Outlook add-in. The Exchange Admin Center displays the disabled status, and the add-in is disabled for use in the Outlook rich clients, Outlook Web App and OWA for Devices.
    
    Using a group policy or application-specific setting in the Windows registry, administrators can adjust this number of times to retry evaluation in the  **OutlookActivationManagerRetryLimit** setting. For more information, see [Overriding resource usage settings for performance of Office Add-ins](http://msdn.microsoft.com/library/da14ec8c-5075-4035-a951-fc3c2b15c04b%28Office.15%29.aspx).
    

    **Task pane and content add-ins**
    
    If any content or task pane add-in exceeds the preceding thresholds on CPU core or memory usage, or tolerance limit for crashes, the corresponding host application displays a warning for the user. At this point, the user can do one of the following:
    
  - Restart the add-in.
    
  - Cancel further alerts about exceeding that threshold. Ideally, the user should then delete the add-in from the document; continuing the add-in would risk further performance and stability issues.
    

## Verifying resource usage issues in the Telemetry Log


Office provides a Telemetry Log that maintains a record of certain events (loading, opening, closing, and errors) of Office solutions running on the local computer, including resource usage issues in an Office Add-in. If you have the Telemetry Log set up, you can use Excel to open the Telemetry Log in the following default location on your local drive:

%Users%\ [Current user]\AppData\Local\Microsoft\Office\15.0\Telemetry

For each event that the Telemetry Log tracks for an add-in, there is a date/time of the occurrence, event ID, severity, and short descriptive title for the event, the friendly name and unique ID of the add-in, and the application that logged the event. You can refresh the Telemetry Log to see the current tracked events. The following table shows examples of Outlook add-ins that were tracked in the Telemetry log. 



|**Date/Time**|**Event ID**|**Severity**|**Title**|**File**|**ID**|**Application**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|10/8/2012 5:57:10 PM|7||add-in manifest downloaded successfully|Who's Who|69cc567c-6737-4c49-88dd-123334943a22|Outlook|
|10/8/2012 5:57:01 PM|7||add-in manifest downloaded successfully|LinkedIn|333bf46d-7dad-4f2b-8cf4-c19ddc78b723|Outlook|
 The following table lists the events that the Telemetry Log tracks for Office Add-ins in general.



|**Event ID**|**Title**|**Severity**|**Description**|
|:-----|:-----|:-----|:-----|
|7|Add-in manifest downloaded successfully||The manifest of the Office Add-in was successfully loaded and read by the host application.|
|8|Add-in manifest did not download|Critical|The host application was unable to load the manifest file for the Office Add-in from the SharePoint catalog, corporate catalog, or the Office Store.|
|9|Add-in markup could not be parsed|Critical|The host application loaded the Office Add-in manifest, but could not read the HTML markup of the app.|
|10|Add-in used too much CPU|Critical|The Office Add-in used more than 90% of the CPU resources over a finite period of time.|
|15|Add-in disabled due to string search time-out||Outlook add-ins search the subject line and message of an e-mail to determine whether they should be displayed by using a regular expression. The Outlook add-in listed in the  **File** column was disabled by Outlook because it timed out repeatedly while trying to match a regular expression.|
|18|Add-in closed successfully||The host application was able to close the Office Add-in successfully.|
|19|Add-in encountered runtime error|Critical|The Office Add-in had a problem that caused it to fail. For more details, look at the  **Microsoft Office Alerts** log using the Windows Event Viewer on the computer that encountered the error.|
|20|Add-in failed to verify licensing|Critical|The licensing information for the Office Add-in could not be verified and may have expired. For more details, look at the  **Microsoft Office Alerts** log using the Windows Event Viewer on the computer that encountered the error.|
For more information, see [Deploying Telemetry Dashboard](http://msdn.microsoft.com/en-us/library/f69cde72-689d-421f-99b8-c51676c77717%28Office.15%29.aspx) and [Troubleshooting Office files and custom solutions with the telemetry log](http://msdn.microsoft.com/library/ef88e30e-7537-488e-bc72-8da29810f7aa%28Office.15%29.aspx)


## Design and implementation techniques


While the resources limits on CPU and memory usage, crash tolerance, UI responsiveness apply to Office Add-ins running only on the rich clients, optimizing the usage of these resources and battery should be a priority if you want your add-in to perform satisfactorily on all supporting clients and devices. Optimization is particularly important if your add-in carries out long-running operations or handles large data sets. The following list suggests some techniques to break up CPU-intensive or data-intensive operations into smaller chunks so that your add-in can avoid excessive resource consumption and the host application can remain responsive:


- In a scenario where your add-in needs to read a large volume of data from an unbounded dataset, you can apply paging when reading the data from a table, or reduce the size of data in each shorter read operation, rather than attempting to complete the read in one single operation. 
    
    For a JavaScript and jQuery code sample that shows breaking up a potentially long-running and CPU-intensive series of inputting and outputting operations on unbounded data, see [How can I give control back (briefly) to the browser during intensive JavaScript processing?](http://stackoverflow.com/questions/210821/how-can-i-give-control-back-briefly-to-the-browser-during-intensive-javascript). This example uses the [setTimeout](http://msdn.microsoft.com/en-us/library/ie/ms536753%28v=vs.85%29.aspx) method of the global object to limit the duration of input and output. It also handles the data in defined chunks instead of randomly unbounded data.
    
- If your add-in uses a CPU-intensive algorithm to process a large volume of data, you can use web workers to perform the long-running task in the background while running a separate script in the foreground, such as displaying progress in the user interface. Web workers do not block user activities and allow the HTML page to remain responsive. For an example of web workers, see [The Basics of Web Workers](http://www.mdl5rocks.com/en/tutorials/workers/basics/). See [Web Workers](http://msdn.microsoft.com/en-us/library/IE/hh772807%28v=vs.85%29.aspx) for more information about the Internet Explorer Web Workers API.
    
- If your add-in uses a CPU-intensive algorithm but you can divide the data input or output into smaller sets, consider creating a web service, passing the data to the web service to off-load the CPU, and wait for an asynchronous callback.
    
- Test your add-in against the highest volume of data you expect, and restrict your add-in to process up to that limit.
    

## Additional resources



- [Privacy and security for Office Add-ins](../../docs/develop/privacy-and-security.md)
    
- [Limits for activation and JavaScript API for Outlook add-ins](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    
