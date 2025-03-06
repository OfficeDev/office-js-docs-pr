---
title: Avoid using the context.sync method in loops
description: Learn how to use the split loop and correlated objects patterns to avoid calling context.sync in a loop.
ms.topic: best-practice
ms.date: 01/31/2025
ms.localizationpriority: medium
---


# Avoid using the context.sync method in loops

> [!NOTE]
> This article assumes that you're beyond the beginning stage of working with at least one of the four application-specific Office JavaScript APIs&mdash;for Excel, Word, OneNote, and Visio&mdash;that use a batch system to interact with the Office document. In particular, you should know what a call to `context.sync` does and you should know what a collection object is. If you're not at that stage, please start with [Understanding the Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md) and the documentation linked to under "application-specific" in that article.

Office Add-ins that use one of the [application-specific API models](../develop/application-specific-api-model.md) may have scenarios that require your code to read or write some property from every member of a collection object. For example, an Excel add-in that gets the values of every cell in a particular table column or a Word add-in that highlights every instance of a string in the document. You will need to iterate over the members in the `items` property of the collection object; but, for performance reasons, you should to avoid calling `context.sync` in every iteration of the loop. Every call of `context.sync` is a round trip from the add-in to the Office document. Repeated round trips hurt performance, especially if the add-in is running in Office on the web because the round trips go across the internet.

> [!NOTE]
> All examples in this article use `for` loops but the practices described apply to any loop statement that can iterate through an array, including the following:
>
> - `for`
> - `for of`
> - `while`
> - `do while`
>
> They also apply to any array method to which a function is passed and applied to the items in the array, including the following:
>
> - `Array.every`
> - `Array.forEach`
> - `Array.filter`
> - `Array.find`
> - `Array.findIndex`
> - `Array.map`
> - `Array.reduce`
> - `Array.reduceRight`
> - `Array.some`

> [!NOTE]
> It's generally a good practice to put have a final `context.sync` just before the closing "}" character of the application `run` function (such as `Excel.run`, `Word.run`, etc.). This is because the `run` function makes a hidden call of `context.sync` as the last thing it does if, and only if, there are queued commands that haven't yet been synchronized. The fact that this call is hidden can be confusing, so we generally recommend that you add the explicit `context.sync`. However, given that this article is about minimizing calls of `context.sync`, it is actually more confusing to add an entirely unnecessary final `context.sync`. So, in this article, we leave it out when there are no unsynchronized commands at the end of the `run`.

## Writing to the document

In the simplest case, you are only writing to members of a collection object, not reading their properties. For example, the following code highlights in yellow every instance of "the" in a Word document.

```javascript
await Word.run(async function (context) {
  let startTime, endTime;
  const docBody = context.document.body;

  // search() returns an array of Ranges.
  const searchResults = docBody.search('the', { matchWholeWord: true });
  searchResults.load('font');
  await context.sync();

  // Record the system time.
  startTime = performance.now();

  for (let i = 0; i < searchResults.items.length; i++) {
    searchResults.items[i].font.highlightColor = '#FFFF00';

    await context.sync(); // SYNCHRONIZE IN EACH ITERATION
  }
  
  // await context.sync(); // SYNCHRONIZE AFTER THE LOOP

  // Record the system time again then calculate how long the operation took.
  endTime = performance.now();
  console.log("The operation took: " + (endTime - startTime) + " milliseconds.");
})
```

The preceding code took 1 full second to complete in a document with 200 instances of "the" in Word on Windows. But when the `await context.sync();` line inside the loop is commented out and the same line just after the loop is uncommented, the operation took only a 1/10th of a second. In Word on the web (with Edge as the browser), it took 3 full seconds with the synchronization inside the loop and only 6/10ths of a second with the synchronization after the loop, about five times faster. In a document with 2000 instances of "the", it took (in Word on the web) 80 seconds with the synchronization inside the loop and only 4 seconds with the synchronization after the loop, about 20 times faster.

> [!NOTE]
> It's worth asking whether the synchronize-inside-the-loop version would execute faster if the synchronizations ran concurrently, which could be done by simply removing the `await` keyword from the front of the `context.sync()`. This would cause the runtime to initiate the synchronization and then immediately start the next iteration of the loop without waiting for the synchronization to complete. However, this isn't as good a solution as moving the `context.sync` out of the loop entirely for the following reasons.
>
> - Just as the commands in a synchronization batch job are queued, the batch jobs themselves are queued in Office, but Office supports no more than 50 batch jobs in the queue. Any more triggers errors. So, if there are more than 50 iterations in a loop, there's a chance that the queue size is exceeded. The greater the number of iterations, the greater the chance of this happening.
> - "Concurrently" doesn't mean simultaneously. It would still take longer to execute multiple synchronization operations than to execute one.
> - Concurrent operations aren't guaranteed to complete in the same order in which they started. In the preceding example, it doesn't matter what order the word "the" gets highlighted, but there are scenarios where it's important that the items in the collection be processed in order.

## Read values from the document with the split loop pattern

Avoiding `context.sync` inside a loop becomes more challenging when the code must *read* a property of the collection items as it processes each one. Suppose your code needs to iterate all the content controls in a Word document and log the text of the first paragraph associated with each control. Your programming instincts might lead you to loop over the controls, load the `text` property of each (first) paragraph, call `context.sync` to populate the proxy paragraph object with the text from the document, and then log it. The following is an example.

```javascript
Word.run(async (context) => {
    const contentControls = context.document.contentControls.load('items');
    await context.sync();

    for (let i = 0; i < contentControls.items.length; i++) {
      // The sync statement in this loop will degrade performance.
      const paragraph = contentControls.items[i].getRange('Whole').paragraphs.getFirst(); 
      paragraph.load('text');
      await context.sync();
      console.log(paragraph.text);
    }
});
```

In this scenario, to avoid having a `context.sync` in a loop, you should use a pattern we call the **split loop** pattern. Let's see a concrete example of the pattern before we get to a formal description of it. Here's how the split loop pattern can be applied to the preceding code snippet. Note the following about this code.

- There are now two loops and the `context.sync` comes between them, so there's no `context.sync` inside either loop.
- The first loop iterates through the items in the collection object and loads the `text` property, just as the original loop did, but the first loop cannot log the paragraph text because it no longer contains a `context.sync` to populate the `text` property of the `paragraph` proxy object. Instead, it adds the `paragraph` object to an array.
- The second loop iterates through the array that was created by the first loop, and logs the `text` of each `paragraph` item. This is possible because the `context.sync` that came between the two loops populated all the `text` properties.

```javascript
Word.run(async (context) => {
    const contentControls = context.document.contentControls.load("items");
    await context.sync();

    const firstParagraphsOfCCs = [];
    for (let i = 0; i < contentControls.items.length; i++) {
      const paragraph = contentControls.items[i].getRange('Whole').paragraphs.getFirst();
      paragraph.load('text');
      firstParagraphsOfCCs.push(paragraph);
    }

    await context.sync();

    for (let i = 0; i < firstParagraphsOfCCs.length; i++) {
      console.log(firstParagraphsOfCCs[i].text);
    }
});
```

The preceding example suggests the following procedure for turning a loop that contains a `context.sync` into the split loop pattern.

1. Replace the loop with two loops.
2. Create a first loop to iterate over the collection and add each item to an array while also loading any property of the item that your code needs to read.
3. Follow the first loop with `context.sync` to populate the proxy objects with any loaded properties.
4. Follow the `context.sync` with a second loop to iterate over the array created in the first loop and read the loaded properties.

## Process objects in the document with the correlated objects pattern

Let's consider a more complex scenario where processing the items in the collection requires data that isn't in the items themselves. The scenario envisions a Word add-in that operates on documents created from a template with some boilerplate text. Scattered in the text are one or more instances of the following placeholder strings: "{Coordinator}", "{Deputy}", and "{Manager}". The add-in replaces each placeholder with some person's name. While the UI of the add-in isn't important to this article, the add-in could have a task pane with three text boxes, each labeled with one of the placeholders. The user enters a name in each text box and then presses a **Replace** button. The handler for the button creates an array that maps the names to the placeholders, and then replaces each placeholder with the assigned name.

You can use the [Script Lab tool](../overview/explore-with-script-lab.md) to follow along with the code snippets shown here. In Word, you can load the "Correlated objects pattern" sample or [import this sample code from the GitHub repo](https://raw.githubusercontent.com/OfficeDev/office-js-snippets/refs/heads/prod/samples/word/90-scenarios/correlated-objects-pattern.yaml).

The following assignment statement creates the mapping array between placeholder and assigned names.

```javascript
const jobMapping = [
        { job: "{Coordinator}", person: "Sally" },
        { job: "{Deputy}", person: "Bob" },
        { job: "{Manager}", person: "Kim" }
    ];
```

The following code shows how you might replace each placeholder with its assigned name if you used `context.sync` inside loops. This corresponds to the `replacePlaceholdersSlow` function in the sample.

```javascript
Word.run(async (context) => {
    // The context.sync calls in the loops will degrade performance.
    for (let i = 0; i < jobMapping.length; i++) {
      let options = Word.SearchOptions.newObject(context);
      options.matchWildcards = false;
      let searchResults = context.document.body.search(jobMapping[i].job, options);
      searchResults.load('items');

      await context.sync(); 

      for (let j = 0; j < searchResults.items.length; j++) {
        searchResults.items[j].insertText(jobMapping[i].person, Word.InsertLocation.replace);

        await context.sync();
      }
    }
});
```

In the preceding code, there's an outer and an inner loop. Each of them contains a `context.sync` call. Based on the first code snippet in this article, you probably see that the `context.sync` in the inner loop can simply be moved after the inner loop. But that would still leave the code with a `context.sync` (two of them actually) in the outer loop. The following code shows how you can remove `context.sync` from the loops. It corresponds to the `replacePlaceholders` function in the sample. We discuss the code later.

```javascript
Word.run(async (context) => {

    const allSearchResults = [];
    for (let i = 0; i < jobMapping.length; i++) {
      let options = Word.SearchOptions.newObject(context);
      options.matchWildcards = false;
      let searchResults = context.document.body.search(jobMapping[i].job, options);
      searchResults.load('items');
      let correlatedSearchResult = {
        rangesMatchingJob: searchResults,
        personAssignedToJob: jobMapping[i].person
      }
      allSearchResults.push(correlatedSearchResult);
    }

    await context.sync()

    for (let i = 0; i < allSearchResults.length; i++) {
      let correlatedObject = allSearchResults[i];

      for (let j = 0; j < correlatedObject.rangesMatchingJob.items.length; j++) {
        let targetRange = correlatedObject.rangesMatchingJob.items[j];
        let name = correlatedObject.personAssignedToJob;
        targetRange.insertText(name, Word.InsertLocation.replace);
      }
    }

    await context.sync();
});
```

Note the code uses the split loop pattern.

- The outer loop from the preceding example has been split into two. (The second loop has an inner loop, which is expected because the code is iterating over a set of jobs (or placeholders) and within that set it is iterating over the matching ranges.)
- There's a `context.sync` after each major loop, but no `context.sync` inside any loop.
- The second major loop iterates through an array that's created in the first loop.

But the array created in the first loop does *not* contain only an Office object as the first loop did in the section [Reading values from the document with the split loop pattern](#read-values-from-the-document-with-the-split-loop-pattern). This is because some of the information needed to process the Word Range objects is not in the Range objects themselves but instead comes from the `jobMapping` array.

So, the objects in the array created in the first loop are custom objects that have two properties. The first is an array of Word ranges that match a specific job title (that is, a placeholder string) and the second is a string that provides the name of the person assigned to the job. This makes the final loop easy to write and easy to read because all of the information needed to process a given range is contained in the same custom object that contains the range. The name that should replace _**correlatedObject**.rangesMatchingJob.items[j]_ is the other property of the same object: _**correlatedObject**.personAssignedToJob_.

We call this variation of the split loop pattern the **correlated objects** pattern. The general idea is that the first loop creates an array of custom objects. Each object has a property whose value is one of the items in an Office collection object (or an array of such items). The custom object has other properties, each of which provides information needed to process the Office objects in the final loop. See the section [Other examples of these patterns](#other-examples-of-these-patterns) for a link to an example where the custom correlating object has more than two properties.

One further caveat: sometimes it takes more than one loop just to create the array of custom correlating objects. This can happen if you need to read a property of each member of one Office collection object just to gather information that will be used to process another collection object. (For example, your code needs to read the titles of all the columns in an Excel table because your add-in is going to apply a number format to the cells of some columns based on that column's title.) But you can always keep the `context.sync`s between the loops, rather than in a loop. See the section [Other examples of these patterns](#other-examples-of-these-patterns) for an example.

## Other examples of these patterns

- For a very simple example for Excel that uses `Array.forEach` loops, see the accepted answer to this Stack Overflow question: [Is it possible to queue more than one context.load before context.sync?](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)
- For a simple example for Word that uses `Array.forEach` loops and doesn't use `async`/`await` syntax, see the accepted answer to this Stack Overflow question: [Iterating over all paragraphs with content controls with Office JavaScript API](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api).
- For an example for Word that is written in TypeScript, see the sample [Word Add-in Angular2 Style Checker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker), especially the file [word.document.service.ts](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts). It has a mixture of `for` and `Array.forEach` loops.
- For an advanced Word sample, import [this gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab) into the [Script Lab tool](../overview/explore-with-script-lab.md). For context in using the gist, see the accepted answer to the Stack Overflow question [Document not in sync after replace text](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text). This sample creates a custom correlating object type that has three properties. It uses a total of three loops to construct the array of correlated objects, and two more loops to do the final processing. There are a mixture of `for` and `Array.forEach` loops.
- Although not strictly an example of the split loop or correlated objects patterns, there's an advanced Excel sample that shows how to convert a set of cell values to other currencies with just a single `context.sync`. To try it, open the [Script Lab tool](../overview/explore-with-script-lab.md) then search for and navigate to the **Currency Converter** sample.

## When should you *not* use the patterns in this article?

Excel can't read more than 5MB of data in a given call of `context.sync`. If this limit is exceeded, an error is thrown. (See the "Excel add-ins section" of [Resource limits and performance optimization for Office Add-ins](resource-limits-and-performance-optimization.md#excel-add-ins) for more information.) It's very rare that this limit is approached, but if there's a chance that this will happen with your add-in, then your code should *not* load all the data in a single loop and follow the loop with a `context.sync`. But you still should avoid having a `context.sync` in every iteration of a loop over a collection object. Instead, define subsets of the items in the collection and loop over each subset in turn, with a `context.sync` between the loops. You could structure this with an outer loop that iterates over the subsets and contains the `context.sync` in each of these outer iterations.
