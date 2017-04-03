# Implementation details, for those who want to know how it *really* works

| | |
|:--|:--|
|[![An image of the Building Office Add-ins using Office.js book cover](../../images/book-cover.png)](https://leanpub.com/buildingofficeaddins)|**This article is an excerpt from the book "[Building Office Add-ins using Office.js](https://leanpub.com/buildingofficeaddins)" by Michael Zlatkovsky, available for purchase as an e-book on [LeanPub.com](https://leanpub.com/buildingofficeaddins).**<br/><br/>Copyright © 2016-2017 by Michael Zlatkovsky, all rights reserved.|

> *In writing this book and receiving feedback from early readers, I've heard a couple requests for a more thorough explanation of what happens under the covers with all this proxy-object / syncing business.  So, if you're the sort of person who likes to see the **implementation details** in order to better understand the outer behavior of an API, read on.  If you're not, feel free to skip to the next section.*


## The Request Context queue

At the heart of the new wave of Office 2016 APIs is a Request Context -- which is the object you receive as a parameter to the batch function, inside of an `Excel.run`.  Fundamentally, you can think of a Request Context object as a central repository that accumulates any changes you'd like to do to the document.  I say "repository", because the Request Context could indeed be likened to a version-control system, where all you send is the *diff*-s between the local state and the remote state.

>**Note:** Git is a particularly well-suited version-control analogy, as local changes are so perfectly isolated from the repository:  until you do a `git push` of your local state, the repository has *no knowledge whatsoever* about the changes being made!  The Request Context and proxy objects of the new Office.js model are very much the same: they are completely unknown to the document until the developer issues a `context.sync()` command!  


The Request Context holds two arrays that allow it do it its work.  One is for **object paths**: descriptions of how to derive one object from another (i.e., "*call method `getRow` with parameter value `2` on <insert-some-preceding-object-path> in order to derive this object*").  The other is for **actions** (i.e., *set the property named "color" to a value of "purple" on the object described by object path #xyz*).  For those who are familiar with the "Command" Design Pattern, this notion of carrying around objects that represent the recipe for a particular action should sound quite familiar.

On the Request Context is a single root object that connects it to the underlying object model.  For Excel, this object is a `workbook`; for Word, it is `document`.  From there, you can derive new objects by calling methods on that root proxy object, or on any of its descendants.  For example, to get a worksheet named "Report", you would ask the `workbook` object for its `worksheets` property (which returns a proxy object corresponding to the worksheets collection in the document), and then use `worksheets` to call a `getItem("Report")` method in order to get a proxy object corresponding to the desired "Report" worksheet.  Each of these objects carries a link to its original Request Context, which in turn keeps track of each object's path info: namely, who was the parent of this new object, and what were the circumstances under which it got created (*was it a property or a method-call? were there any parameters passed in?*).

Whenever a method or property gets called on a proxy object, the call is registered as an **action** on the Request Context. For example, a call to `range.merge()` or the setting of `fill.color = "purple"`, will get put in the queue as action such-and-such on object such-and-such.  Moreover, if the result of the method or property call is another proxy object (for example, `worksheets.getItem("Report")` or `worksheets.add()`, a new proxy object will be generated as a *side-effect* of the method call, and its lineage will be dutifully noted by the omniscient Request Context.

Let's walk through a real example.  Suppose you have the following code:

**Tracing through the Request Context / proxy object operations**
~~~
    Excel.run(async (context) => {
        let range = context.workbook.getSelectedRange();
        range.clear();
        let thirdRow = range.getRow(2);
        thirdRow.format.fill.color = "purple";

        await context.sync();
    }).catch(OfficeHelpers.Utilities.log);
~~~

Now let's analyze it with a fine-toothed comb. With each API call, I will keep a running tally of the object path and their actions (expressed in a friend-liefied and shortened notation, but following closely to what happens internally).

To start, line **#1** -- `Excel.run(async (context) => {` -- uses `Excel.run` to create a Request Context object. The `.run` invocation does a number of other things too, but let's leave it be for the time being. The important thing is that it gives us a brand new `context` object, on which there is already a pre-initialized a `workbook` object (which we'll use momentarily).  

~~~
    objectPaths:
        // markua-start-insert
        1 => global object (workbook)
        // markua-end-insert

    actions: <none>,
~~~


On line **#2** -- `let range = context.workbook.getSelectedRange()` -- we use that `workbook` object to derive a new object, corresponding to the current selection. We assign it to a variable called `range`, but it doesn't matter to the Request Context:  even if we hadn't named it anything, and insitead used it as a pass-through object to get to another destination, it would still get reflected in the Request Context's list. The creation of the object also gets reflected as an object-initialization action in the actions list, for purposes described later in this section.

~~~
    objectPaths:
        P1 => global object (workbook)
        // markua-start-insert
        P2 => (range)
                parent: "P1", type: "method",
                name: "getSelectedRange", args: <none>
        // markua-end-insert

    actions:
        // markua-start-insert
        A1 => action: "init", object: "P2" (range)
        // markua-end-insert
~~~


Line **#3** -- `range.clear()` -- adds the first real document-impacting action: a command to clear the contents of the range:

~~~
    objectPaths:
        P1 => global object (workbook)
        P2 => (range)
                parent: "P1", type: "method",
                name: "getSelectedRange", args: <none>

    actions:
        A1 => action: "init", object: "P2" (range)
        // markua-start-insert
        A2 => action: "method", object: "P2" (range)
                name: "clear", args: <none>
        // markua-end-insert
~~~


Line **#4** -- `let thirdRow = range.getRow(2)` -- follows a similar pattern as line #2, creating a `thirdRow` object that derives from the previously-defined `range` object, and adding another instantiation action:

~~~
    objectPaths:
        P1 => global object (workbook)
        P2 => (range)
                parent: "P1", type: "method",
                name: "getSelectedRange", args: <none>
        // markua-start-insert
        P3 => (thirdRow)
                parent: "P2", type: "method",
                name: "getRow", args: [2]
        // markua-end-insert

    actions:
        A1 => action: "init", object: "P2" (range)
        A2 => action: "method", object: "P2" (range)
                name: "clear", args: <none>
        // markua-start-insert
        A3 => action: "init", object: "P3" (thirdRow)
        // markua-end-insert
~~~


Line **#5** -- `thirdRow.format.fill.color = "purple"` -- is packed with several API calls.  We begin with creating an [anonymous] format object, by following the `format` property of the `thirdRow` variable.  We then do the same for the [anonymous] fill object.  Both follow the same pattern as before, creating an object path and an instatiation action for each.  But then, having reached the desired object, we do another document-impacting action on the object: setting the fill color of the third row to purple (see action "**A6**" below):

~~~
    objectPaths:
        P1 => global object (workbook)
        P2 => (range)
                parent: "P1", type: "method",
                name: "getSelectedRange", args: <none>
        P3 => (thirdRow)
                parent: "P2", type: "method",
                name: "getRow", args: [2]
        // markua-start-insert
        P4 => (format)
                parent: "P3", type: "property",
                name: "format"
        P5 => (fill)
                parent: "P4", type: "property",
                name: "fill"
        // markua-end-insert

    actions:
        A1 => action: "init", object: "P2" (range)
        A2 => action: "method", object: "P2" (range)
                name: "clear", args: <none>
        A3 => action: "init", object: "P3" (thirdRow)
        // markua-start-insert
        A4 => action: "init", object: "P4" (format)
        A5 => action: "init", object: "P5" (fill)
        A6 => action: "setter", object: "P5" (fill),
                name: "color", value: "purple"
        // markua-end-insert
~~~


And finally, on line **#7** (line #6 was a blank one), we get to the magic **`await context.sync()`** incantation.  This command tells the Request Context to pack up all of the relevant information (namely, pending actions, and any associated object path info-s), and send it to the host application for processing.


On the receiving ends of the host application, the host unpacks the actions and begins iterating through them one-by-one.  It keeps a working dictionary of the objects that got derived during this particular `sync` session -- such that, having retrieved the range corresponding to `thirdRow` once, it will not need to re-evaluate it again.  This is done not only for efficiency, but also to prevent mistakes:  you wouldn't want to re-fetch the row at relative index 2 if another few rows got added between it and the first row;  nor would you want to re-fetch the selection every time, since it may well have shifted (e.g., during the adding and activating of a new worksheet), and yet semantically the range should be *imprinted* with the original reference.  Finally, if you have an object derived from calling the `add` method on the worksheet-collection object, you *definitely* wouldn't want to re-derive the object -- and, as a side effect, add a new sheet -- every time that the object was accessed!

If at any point in the chain something goes wrong, the rest of the batch gets aborted. Going with the previous example, if there is no third row in the selection (i.e., it's a 2x2 cell selection), the remaining commands would get ignored (which is probably what you'd expect, anyway). Importantly, though, there is no *atomicity* to the `Excel.run` or the `sync`: any actions that have already been done will *stay* done.  In the case of this example, the document might be left in a *dirtied* state, where the clearing of the selection has already happened, but the formatting of the third row has not been done yet.  While not ideal, this is no different from VBA or VSTO with regards to Office automation; it is simply too difficult to roll back, especially given any user or collaborator actions that may have happened in the meantime. 

## The host application's response

Let's assume that the `sync` did succeed: that every necessary object (the original selection, its third row, the format, the fill) all were created successfully, and that both document-impacting actions were also able to commit to the document.  What happens next?

As mentioned earlier, the host keeps a running dictionary of the object that it's been working with.  But this running dictionary is *only for the duration of the particular `sync`: not for the lifetime of the application*.  To keep and track the objects indefinitely would be a huge performance hit.

Now, let's take the case where an object path is the "add" action on a worksheet collection.   During the processing of the `sync`, the method would only have been executed once (with the appropriate side-effect of creating the worksheet), and the resulting sheet would be cached.  This is great for the current `sync`, but what if the developer wants to access the sheet again at a later `sync`?  This is where the instantiation actions mentioned earlier come in.

For each action, the host application may *optionally* send a response.  For actions like clearing a range or setting a fill color, there is nothing to respond with (the fact that the operation succeeded is obvious through the fact that the queue kept executing to completion).  But for instantiation actions, the host *may* send a response to tell JavaScript to re-map its object path to something less volatile. Thus, while the original path for a newly-created sheet may have been "*execute method '`add` on object xyz*" (where xyz is the worksheet collection), the response might indicate "*from here on out, refer to the sheet as being a "getItem" invocation with parameter "123456789" on that same xyz object".  That is, in creating the object and executing the instantiation action, the host can figure out if there is a more permanent ID it can give back to JavaScript for future references to this object.  (A less drastic example: fetching a sheet by name is somewhat risky, in that names can change, both via user interaction and programmatically; but if the host can re-map the path to a permanent worksheet ID, any future invocations on the object are guaranteed to continue to refer to the same sheet, regardless of its name).

But there is another, even more important use for responses from the host.  Suppose, on the JavaScript side, you have a call to `range.load("formulas")`. In terms of actions, this gets represented as a *query* action on the object, with a parameter whose value is "formulas". To this action, the host will respond by fetching the appropriate object (which is already in its dictionary, thanks to the instantiation action), querying it for the required properties, and returning the requested information.


## Back on the proxy object's territory

Back in JavaScript, the `sync` is patiently waiting for a response from the host application.  And, hopefully, the developer's code is *also* patiently waiting, by either using an `await` or subscribing to the `.then` function-call of the `sync` Promise.

When the response *does* come back, there is a bit of internal processing before the execution gets back to the developer's code.  For example, any of the path-remappings, described in the preceding section, take effect.  There is also some internal processing (e.g., invalidating the paths of objects that were valid during the previous `sync` batch, but cannot be used again -- I'll explain more soon).  And, importantly, the results of any *query* actions take effect, taking the loaded values and putting them back on the corresponding objects and properties.  This ensures that, following the `sync` if the developer's code now references `range.values` for a Range whose values have been loaded, he/she will get the last known snapshot of the values (as opposed to a `PropertyNotLoaded` error).

With the post-processing done, the request context is now ready to be re-used.  Its actions array was reset to a blank slate at the very beginning of the `sync`; and conversely, the object paths array (which is never emptied during the lifetime of the particular request context, as later actions are bound to re-use some of the existing paths) has had any of the object paths tweaked, based on responses from the host and post-processing.  And so a new batch of operations can begin, queuing up until the next `await context.sync()`.


## A special (but common) case: objects without IDs

When working with objects like worksheets (Excel) or content controls (Word), the host application's job is quite easy: in both cases, there is a permanent ID attached to each of those objects, so no matter how the object was created (`getActiveWorksheet()`, or `getItem`, or whatever other invocation), the host can always use the instantiation action to re-map the path back to a permanent ID.  Which means that, as a developer, having created the object once at some point, you can continue to use it in the next `sync`, or even longer thereafter.  No surprises there.

But what about objects that don't have IDs; and that, by definition, have an infinite number of permutations about them?  Both Excel ranges (a particular grouping of cells) and Word ranges  (some text starting at one location and ending in another) are not at all easy to get a concrete reference to: the address/index at which they are might shift, and the ranges might also grow and expand if additional cells or characters are added within them.  The same is true for some other objects.

The host doesn't have issues tracking the objects during the *processing* of the batch, as there it can use an internal dictionary and, having retrieved the object once, continue to cling on to it for the duration of the batch.  But as noted earlier, the host is *stateless* across batches -- it cannot afford to keep a reference to every object ever accessed on it from JavaScript, or else the application will leak memory and crawl to a halt.

To avoid bogging down the host application with thousands of no-longer-needed objects, while still enabling the very common scenario of needing to do a `load` and a `sync` before performing further actions with an object, the original design was as follows:

1. By default, any object without an ID quietly loses connection with the underlying document, and cannot be used again after the current `sync`. It can still be read-from by JavaScript *after* the sync (otherwise, there would be no point to loading it in the first place!), and that might be fine for some scenarios -- but this default behavior would preclude, for example, the highlighting example where the range values are first read, and then colored accordingly.
2. To enable the latter scenario, we expose an API -- `context.trackedObjects.add` -- where a developer can explicitly tell the host that "*I want to track this object longer-term, not just for the current `sync`*.  The developer is then responsible for calling `context.trackedObjects.remove` when the object is no longer needed (no hard penalty if he/she doesn't, but it will slow down the host application over time, so use object-tracking sparingly, and try to release as soon as you're done).

On the JavaScript layer, a call to `context.trackedObjects.add` would add a new type of action to the queue, saying that object with id such-and-such would like to be tracked.  On the host side, this action would be interpreted to create a permanent wrapper around the in-memory object, creating a made-up ID that the object could use as if it were a real ID. This ID would be sent back to the object, very much like the object-path-remapping result of an instantiation action.  And likewise, a call to `context.trackedObjects.remove` would likewise get a special action added on the queue, requesting that the host release the memory for the no-longer-needed object, and marking the object itself as no longer having a valid path.

This design worked -- and in fact, it still works today, if a developer chooses to create a Request Context manually, via `var context = new Excel.RequestContext()`, instead of a `.run`.  But in practice, in both our internal testing and public preview, it tuned out to be very tedious to have to call `context.trackedObjects.add` on an object or two within nearly every scenario.  And even when developers did call it (with some trial-and-error), it was even more tedious (nay, unrealistic) to expect that folks will remember and correctly dispose of the no-longer-needed tracked objects.

In observing folks struggle with this tracked-objects concept, one thing that became clear is that in the vast majority of cases, the developer's intent is *not* to keep the object around for some long-term storage -- rather, the developer generally just needs to track the object so that they can use it across one or two `sync` boundaries, and then they are done with it forever.  This is where the `Excel.run` (`Word.run`, etc.) concept was born:  to allow developers to declare a single semantic unit of automation, even if internally it is comprised of a series of `sync`-s.  And for the framework to handle the tracking and untracking silently.

This means that whenever you do an `Excel.run` (`Word.run`, etc.), after each instantiation action there is *also* an action to track the object.  And at the very end, after a final flush of the queue at the completion of the `Excel.run`, there is a separate internal request made to un-track every non-ID-able and non-derivable object that had been created in the meantime.  So the *true* picture of the "actions" array from above is actually a bit more verbose than shown earlier:

~~~
    actions:
        A1 => action: "init", object: "P2" (range)
        // markua-start-insert
        A2 => action: "track", object: "P2" (range)
        // markua-end-insert
        A3 => action: "method", object: "P2" (range)
                name: "clear", args: <none>
        A4 => action: "init", object: "P3" (thirdRow)
        // markua-start-insert
        A5 => action: "track", object: "P3" (thirdRow)
        // markua-end-insert
        A6 => action: "init", object: "P4" (format)
        A7 => action: "init", object: "P5" (fill)
        A8 => action: "setter", object: "P5" (fill),
                name: "color", value: "purple"
~~~


And then, at the completion of the `run`, after a final flush of the queue, the following would be sent (note that only the relevant object paths are being sent; there is no need to carry extra baggage over the wire/process-boundary):


~~~
    objectPaths:
        P1 => global object (workbook)
        P2 => (range)
                parent: "P1", type: "method",
                name: "getSelectedRange", args: <none>
        P3 => (thirdRow)
                parent: "P2", type: "method",
                name: "getRow", args: [2]

    actions:
        A1 => action: "untrack", object: "P2" (range)
        A2 => action: "untrack", object: "P3" (thirdRow)
~~~


And this -- in a not so small nutshell -- is how the underlying proxy objects work, and how the runtime handles its communication to and from the host application.  

>**This article is an excerpt from the book "[Building Office Add-ins using Office.js](https://leanpub.com/buildingofficeaddins)" by Michael Zlatkovsky**. Read more by purchasing the e-book online at [LeanPub.com](https://leanpub.com/buildingofficeaddins).
