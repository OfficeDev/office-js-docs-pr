---
title: Usability testing for Office Add-ins
description: Learn how to test your add-in design with real users.
ms.date: 01/13/2025
ms.localizationpriority: medium
---

# Usability testing for Office Add-ins

A great add-in design takes user behaviors into account. Because your own preconceptions influence your design decisions, it's important to test designs with real users to make sure that your add-ins work well for your customers.

You can run usability tests in different ways. For many add-in developers, remote, unmoderated usability studies are the most time and cost effective. Popular testing services include:

- [UserTesting.com](https://www.UserTesting.com)
- [Optimalworkshop.com](https://www.Optimalworkshop.com)
- [Userzoom.com](https://www.Userzoom.com)

These testing services help you to streamline test plan creation and remove the need to seek out participants or moderate the tests.

You need only five participants to uncover most usability issues in your design. Incorporate small tests regularly throughout your development cycle to ensure that your product is user-centered.

> [!NOTE]
> We recommend that you test the usability of your add-in across multiple platforms. To [publish your add-in to Microsoft Marketplace](/partner-center/marketplace-offers/submit-to-appsource-via-partner-center), it must work on all [platforms that support the methods that you define](/javascript/api/requirement-sets).

## 1. Sign up for a testing service

For more information, see [Selecting an Online Tool for Unmoderated Remote User Testing](https://www.nngroup.com/articles/unmoderated-user-testing-tools/).

## 2. Develop your research questions

Research questions define the objectives of your research and guide your test plan. Your questions will help you identify participants to recruit and the tasks they'll perform. Understand when you need specific observations or broad input.

### Specific question examples

- Do users notice the "free trial" link on the landing page?
- When users insert content from the add-in to their document, do they understand where in the document it's inserted?

### Broad question examples

- What are the biggest pain points for the user in our add-in?
- Do users understand the meaning of the icons in our command bar, before they click on them?
- Can users easily find the settings menu?

### User experience aspects

It's important to get data on the entire user journey – from discovering your add-in, to installing and using it. Consider research questions that address the following aspects of the add-in user experience.

- Finding your add-in in Microsoft Marketplace
- Choosing to install your add-in
- First-run experience
- Ribbon commands
- Add-in UI
- How the add-in interacts with the document space of the Office application
- How much control the user has over any content insertion flows

For more information, see [Gathering factual responses vs. subjective data](https://help.usertesting.com/hc/articles/11880238504221).

## 3. Identify participants to target

Remote testing services can give you control over many characteristics of your test participants. Think carefully about what kinds of users you want to target. In your early stages of data collection, it might be better to recruit a wide variety of participants to identify more obvious usability issues. Later, you might choose to target groups like advanced Office users, particular occupations, or specific age ranges.

## 4. Create the participant screener

The screener is the set of questions and requirements you present to prospective test participants to screen them for your test. Keep in mind that participants for services like UserTesting.com have a financial interest in qualifying for your test. It's a good idea to include trick questions in your screener if you want to exclude certain users from the test.

For example, if you want to find participants who are familiar with GitHub, to filter out users who might misrepresent themselves, include fakes in the list of possible answers.

**Which of the following source code repositories are you familiar with?**  
 a. SourceShelf  [*Reject*]  
 b. CodeContainer  [*Reject*]  
 c. GitHub  [*Must select*]  
 d. BitBucket  [*May select*]  
 e. CloudForge  [*May select*]  

If you're planning to test a live build of your add-in, the following questions can screen for users who will be able to do this.

**This test requires you to have the latest version of Microsoft PowerPoint. Do you have the latest version of PowerPoint?**  
 a. Yes [*Must select*]  
 b. No [*Reject*]  
 c. I don't know [*Reject*]  

**This test requires you to install a free add-in for PowerPoint, and create a free account to use it. Are you willing to install an add-in and create a free account?**  
 a. Yes [*Must select*]  
 b. No [*Reject*]  

For more information, see [Screener Questions Best Practices](https://help.usertesting.com/hc/articles/11880418598557).

## 5. Create tasks and questions for participants

Try to prioritize what you want tested so that you can limit the number of tasks and questions for the participant. Some services pay participants only for a set amount of time, so you want to make sure not to go over.

Try to observe participant behaviors instead of asking about them, whenever possible. If you need to ask about behaviors, ask about what participants have done in the past, rather than what they would expect to do in a situation. This tends to give more reliable results.

The main challenge in unmoderated testing is making sure your participants understand your tasks and scenarios. Your directions should be *clear and concise*. Inevitably, if there's potential for confusion, someone will be confused.

Don't assume that your user will be on the screen they're supposed to be on at any given point during the test. Consider telling them what screen they need to be on to start the next task.

For more information, see [Writing Great Tasks](https://help.usertesting.com/hc/articles/11880303389213).

## 6. Create a prototype to match the tasks and questions

You can either test your live add-in, or you can test a prototype. Keep in mind that if you want to test the live add-in, you need to screen for participants that have the latest version of Office, are willing to install the add-in, and are willing to sign up for an account (unless you have logon credentials to provide them.) You'll then need to make sure that they successfully install your add-in.

On average, it takes about 5 minutes to walk users through how to install an add-in. The following is an example of clear, concise installation steps. Adjust the steps based on the specifics of your test.

**Please install the (insert your add-in name here) add-in for PowerPoint, using the following instructions.**

1. Open Microsoft PowerPoint.
1. Select **Blank Presentation.**
1. Select **Home** > **Add-ins**, then select **Get Add-ins**.
1. In the popup window, choose **Store**.
1. Type (Add-in name) in the search box.
1. Choose (Add-in name).
1. Take a moment to look at the Store page to familiarize yourself with the add-in.
1. Choose **Add** to install the add-in.

You can test a prototype at any level of interaction and visual fidelity. For more complex linking and interactivity, consider a prototyping tool like [Figma](https://www.figma.com/). If you just want to test static screens, you can host images online and send participants the corresponding URL, or give them a link to an online PowerPoint presentation.

## 7. Run a pilot test

It can be tricky to get the prototype and your task/question list right. Users might be confused by tasks, or might get lost in your prototype. You should run a pilot test with 1-3 users to work out the inevitable issues with the test format. This will help to ensure that your questions are clear, that the prototype is set up correctly, and that you're capturing the type of data you're looking for.

## 8. Run the test

After you order your test, you'll get email notifications when participants complete it. Unless you've targeted a specific group of participants, the tests are usually completed within a few hours.

## 9. Analyze results

This is the part where you try to make sense of the data you've collected. While watching the test videos, record notes about problems and successes the user has. Avoid trying to interpret the meaning of the data until you have viewed all the results.

A single participant having a usability issue isn't enough to warrant making a change to the design. Two or more participants encountering the same issue suggests that other users in the general population will also encounter that issue.

In general, be careful about how you use your data to draw conclusions. Don't fall into the trap of trying to make the data fit a certain narrative; be honest about what the data actually proves, disproves, or simply fails to provide any insight about. Keep an open mind; user behavior frequently defies designer's expectations.

## See also

- [How to Conduct Usability Testing](https://whatpixel.com/howto-conduct-usability-testing/)  
- [Best Practices for UserTesting](https://help.usertesting.com/hc/articles/11880426022813)  
- [Minimizing Bias](https://downloads.usertesting.com/white_papers/TipSheet_MinimizingBias.pdf)  
