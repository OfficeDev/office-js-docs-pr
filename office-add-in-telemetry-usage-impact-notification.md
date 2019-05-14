# Office Add-ins telemetry update Usage Impact Notification
In a recent update of Microsoft Office, we made changes to the app telemetry we provide to Office Add-ins developer community through the developer portal. We recently discovered an issue that affects the telemetry reported for specific add-ins. If your developed add-in is listed below, telemetry for users on office builds (16.0.11231.* through 16.0.11629.*) may not  be properly reported in some cases due to this issue.

|Add-in	                                   |Host App  
|:----------------------------------------:|:-------------------------------------------------:|
| Mini Calendar and Date Picker |	Excel |
| Emoji Keyboard | PowerPoint / Word |
| Web Viewer | PowerPoint |
| Slice Timer | Excel / PowerPoint |
| Lucidchart Diagrams for PowerPoint | PowerPoint |
| Power BI Tiles | PowerPoint
| EasyTimer | PowerPoint
| PhET Sims - Science / Math | PowerPoint
| Slido | PowerPoint
| Vertex42 Template Gallery | Excel / Word
| Create Time Dimension | Excel
| Pixabay Images | PowerPoint / Word
| Adobe Stock | PowerPoint
| Stock Tile | Excel / PowerPoint
| Imagebank - image management for Office | PowerPoint / Word

We are actively working on a fix, and our initial investigation indicates it may take a few weeks to rectify. During this period, you may see a significant reduction in the reported usage of following apps.

We will provide an update in the week of June 3rd with additional information. We regret any inconvenience that this has caused. 
