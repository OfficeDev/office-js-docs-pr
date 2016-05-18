
# Interaction patterns for Office Add-ins


Office Add-ins can enhance authoring and productivity experiences as well as connect content in Office host applications to larger web-based workflows. A number of common scenarios apply to content, task pane, and Outlook add-ins that you might develop. This article describes some of the most common scenarios and provides recommended interaction patterns for the add-in UX. You can break down, combine, or mix and match these interaction patterns depending on your unique scenarios.

 **Common add-in scenarios**

| Add-in type | Common scenarios |
| ------ | ------ |
|  Content  |  Visualizing data <br> Widgets and tools  |
|  Task pane  |  Transforming and processing data <br> Authoring effectively and efficiently <br> Locating content and inserting data <br> Publishing or uploading content to a web service  |
|  Outlook  |  Bridging between mail content and an external application <br> Giving more information about the content in a mail message or appointment <br> Providing information that helps you be more productive  |

## Visualize data with a content add-in


This example shows a content add-in for Excel that generates a chart from data in a spreadsheet.

In this interaction pattern, the add-in doesn't become active until you select and bind data for generating the chart. It's important to communicate the purpose of the add-in and the steps for activating it in the initial view of the add-in. 

**Content add-in for Excel that generates a chart from data in a spreadsheet**
<br>
![Content app for Excel that generates a chart from data in a spreadsheet](../../../images/off15appUXFig01.png)
<br>
<ul><li><p>To reinforce that you must perform an action before choosing the button, display instructions along with a disabled button (A).</p></li><li><p>After you select a range of cells, the <span class="ui">Create Chart</span> button becomes active (B - C).</p></li><li><p>The visualization fills out the container and replaces the previous view (D).</p></li><li><p>Display any additional UI at the bottom edge of the add-in along with a settings button (gear) to take you to a view where you can reset or manage the add-in.</p></li></ul>Best suited for:
<ul><li><p>Add-ins that require you to select data prior to activation.</p></li></ul>

## Transform content with a task pane add-in


This example shows a task pane add-in that translates text in your document into another language.

In this interaction pattern, you must first select the text you want to translate in the document.

**Task pane add-in that translates text in your document into another language**
<br>
![Task pane app that translates text in your document into another language](../../../images/off15appUXFig02.png)
<br>
<ul><li><p>Communicate the purpose of the add-in in a headline and hint at the fact that you must first make a selection (A).</p></li><li><p>The language menu and <span class="ui">Translate</span> button are disabled, reinforcing that you must perform an action before you can progress. After you select content in the document, these two elements become active (D).</p></li><li><p>After you choose <span class="ui">Translate</span>, the UI unfolds showing the translated content along with a button for inserting it back into the document (E).</p></li><li><p>You can provide a <span class="ui">Clear</span> or <span class="ui">Reset</span> button that returns to the initial view.</p></li></ul>Best suited for:
<ul><li><p>Add-ins that require you to select data prior to activation.</p></li><li><p>UI that unfolds or is revealed as you progress through a scenario.</p></li></ul>

## Process data with a task pane add-in


This example shows a task pane add-in that checks data in Excel.

In this interaction pattern, you must select a range of cells in the spreadsheet to begin.

**Task pane add-in that checks data in Excel**
<br>
![Task pane app that checks data in Excel](../../../images/off15appUXFig03.png)
<br>
<ul><li><p>The purpose of the add-in is described in the headline. Instructions help you get started.</p></li><li><p>The <span class="ui">Send selected data</span> button is disabled, reinforcing that you must perform an action in order to progress (A).</p></li><li><p>After you select a range of cells in their spreadsheet (B), the <span class="ui">Send selected data</span> button is activated.</p></li><li><p>After you choose this button, the UI is replaced with the selected range of cells, a progress bar, and a <span class="ui">Cancel</span> button.</p></li><li><p>The progress bar communicates the status of the process, and the <span class="ui">Cancel</span> button lets you interrupt it (D).</p></li><li><p>When the process is finished, the results are automatically displayed (E). Selecting an element in the list activates the corresponding cell in the spreadsheet.</p></li></ul>Best suited for:
<ul><li><p>Processes that take an indeterminate length of time.</p></li></ul>

## Analyze content with a task pane add-in


This example shows a task pane add-in that displays word definitions as you type.

In this interaction pattern, you must first select text in the document to see results.

**Task pane add-in that displays word definitions as you type**
<br>
![Task pane app that displays word definitions as you type](../../../images/off15appUXFig04.png)
<br>
<ul><li><p>A headline explains the purpose of the add-in and how to get started (A).</p></li><li><p>Auto-search is enabled by default with the option to disable it (B).</p></li><li><p>After you make a selection, the add-in displays the corresponding content (D).</p></li><li><p>Provide a link to display more information (E).</p></li></ul>Best suited for:
<ul><li><p>Add-ins that automatically return content as you author.</p></li><li><p>Add-ins that require you to select content prior to activation.</p></li></ul>

## Locate content with a task pane add-in


This example shows a task pane add-in for searching content.

In this interaction pattern, you enter a string in the search box, or select from a list of featured content to begin.

**Task pane add-in for searching content**
<br>
![Task pane app for searching content](../../../images/off15appUXFig05.png)
<br>
<ul><li><p>The initial view contains a <span class="ui">Search</span> box (A) and a list of featured content (B).</p></li><li><p>When you enter a string in the search box, the search icon is replaced with a close icon (C).</p></li><li><p>Choosing the close icon returns you to the initial view.</p></li></ul>Best suited for:
<ul><li><p>Add-ins that automatically return content as you author.</p></li><li><p>Add-ins that require you to select content prior to activation.</p></li></ul>

## Insert media with a task pane add-in


In this interaction pattern, you can select an image from search results to insert into your document.

**Task pane add-in for inserting an image**
<br>
![Task pane app for inserting an image](../../../images/off15appUXFig06.png)
<br>
<ul><li><p>You filtered a list of search returns (A) and selected content to insert (B).</p></li><li><p>A Detail view of the selected content is displayed (C) with a button that takes you back to the list.</p></li><li><p>An <span class="ui">Insert Photo</span> button is located in the footer (D). After you choose this button, the image is inserted into the document.</p></li><li><p>A short description of where the image came from is included with the inserted content (E). </p></li><li><p>UI in the add-in visually confirms the success of the action.</p></li></ul>Best suited for:
<ul><li><p>Add-ins for inserting content.</p></li></ul>

## Insert selected text with a task pane add-in


In this interaction pattern, you select text from search results to insert into the document.

**Task pane add-in for inserting text**
<br>
![Task pane app for inserting text](../../../images/off15appUXFig07.png)
<br>
<ul><li><p>You have already located a piece of content (A).</p></li><li><p>A disabled <span class="ui">Insert Selection</span> button is displayed in the footer (B).</p></li><li><p>When you select a string of text (C), the <span class="ui">Insert Selection</span> button becomes active.</p></li><li><p>After you choose this button, the add-in inserts the selected text into the document along with a reference to the source of the content (E).</p></li></ul>Best suited for:
<ul><li><p>Add-ins for conducting research and inserting content.</p></li></ul>

## Publish to a web service with a task pane add-in


This example shows a task pane add-in for publishing a document as a blog post.

In this interaction pattern, you have finished writing content in a document and want to post it to your blog.

**Task pane add-in for publishing a document as a blog post**
<br>
![Task pane app for publishing a document as a blog post](../../../images/off15appUXFig08.png)
<br>
<ul><li><p>First, a sign-in form is displayed to enter your credentials (A).</p></li><li><p>Links for creating an account and handling typical sign-in troubles are provided (B). Choosing these links opens these pages in a browser.</p></li><li><p>After you are signed in, the add-in displays a form for creating a new blog post (C).</p></li><li><p>The name of the account you signed in to (and will post to) is shown toward the top of the add-in. The add-in provides a link to preview the post (D). Choosing this link displays the preview in a browser.</p></li><li><p>After you choose <span class="ui">Create post</span>, the add-in displays a view confirming that the document content was posted (E).</p></li><li><p>The add-in provides a link to view the post in a browser (F), as well as a button to create another post (G).</p></li></ul>Best suited for:
<ul><li><p>Add-ins that output, publish, or share content to social networks, blogging sites, and web services.</p></li><li><p>Add-ins that require you to sign into a service.</p></li></ul>

## Get more information about people with an Outlook add-in


 **Example 1**

**Results and details page**
<br>
![Results and details page](../../../images/off15appUXFig09.jpg)
<br>
Best suited for:
<ul><li><p>Exposing the breadth of your content if you have large data sets that are useful to showcase.</p></li><li><p>Details pages that use the full size of the add-in container</p></li><li><p>Navigation models that benefit from a "back and forth" flow.</p></li></ul>
 **Example 2**

**Details page with persistent navigation**
<br>
![Details page with persistent navigation](../../../images/off15appUXFig10.jpg)
<br>
Best suited for:
<ul><li><p>Displaying, by default, the first result of a dataset.</p></li><li><p>Navigation structures that work like tabs (single-level linear navigation).</p></li><li><p>Reducing selection actions by keeping navigation available at all times.</p></li><li><p>Providing room to display navigation at all times.</p></li></ul>

## Get more information about content with an Outlook add-in


 **Example 1**

**Results and details page**
<br>
![Results and details page](../../../images/off15appUXFig11.jpg)
<br>
Best suited for:
<ul><li><p>Exposing the breadth of your content if you have large data sets that are useful to display.</p></li><li><p>Requiring you to make a choice or selection before showing more detail.</p></li><li><p>Details pages that use the full size of the add-in container.</p></li><li><p>Navigation models that benefit from a "back and forth" flow.</p></li></ul>
 **Example 2**

**Details page with secondary content**
<br>
![Details page with secondary content](../../../images/off15appUXFig12.jpg)
<br>
Best suited for:
<ul><li><p>Cases where you want to feature one piece of content.</p></li><li><p>Exposing your content without user interaction.</p></li><li><p>Persistent navigation (can be added to this model for a mix of simplicity and ease of navigation).</p></li></ul>

## Connect to an online service and present data


These examples show interaction patterns for getting data and content from an online service. They can be used in all three add-in types: content add-ins, task pane add-ins, and Outlook add-ins.

 **Example 1**

**Carousel**
<br>
![Carousel](../../../images/off15appUXFig13.jpg)
<br>
Best suited for:
<ul><li><p>Small amounts of data that can be exposed one at a time or in groups.</p></li><li><p>Exposing content in a gallery format, such as slideshows or image galleries.</p></li><li><p>When a next/previous navigation model works well.</p></li></ul>
 **Example 2**

**Wizard**
<br>
![Wizard](../../../images/off15appUXFig14.jpg)
<br>
Best suited for:
<ul><li><p>Content that needs to be presented in a specific order.</p></li><li><p>Large amounts of content that is best consumed in a series of small pieces.</p></li><li><p>Book-like consuming experiences.</p></li><li><p>When a series of steps or actions are required to complete a task.</p></li></ul>
 **Example 3**

**Form, results, and details**
<br>
![Form, results, and details](../../../images/off15appUXFig15.jpg)
<br>
Best suited for:
<ul><li><p>Add-ins that require data entry.</p></li><li><p>An entry-point to a results and details pattern.</p></li></ul>

## Additional resources



- [Design guidelines for Office Add-ins](../add-in-design.md)
    
