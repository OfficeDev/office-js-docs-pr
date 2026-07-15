---
title: Writing and style guidelines for Office Add-ins
description: Learn how to write clear, conversational UI text for Office Add-ins, including guidelines for tone, word choice, error messages, and labels that improve the user experience.
ms.date: 07/14/2026
ms.topic: best-practice
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Writing and style guidelines for Office Add-ins

The text in your add-in shapes how users perceive its quality and trustworthiness. Clear, concise writing helps users complete tasks faster and builds confidence in your add-in. This article covers voice and tone principles, practical guidelines for common UI elements, and examples to help you write effective add-in text.

## Voice and tone principles

Strive to match the voice and tone of the Office UI, which is conversational, engaging, and accessible to users. Apply these principles consistently across all text in your add-in.

- **Use a natural style.** Write the way you speak. Avoid jargon and overly technical words and phrases. Use terms that are familiar to your users.
- **Use simple, direct language.** Use short words and sentences, and active voice in your text.
- **Be consistent.** Use the same words for the same concepts throughout.
- **Engage the user.** Address the user as "you". Avoid using third person. Use imperatives for user tasks.
- **Be helpful and empathetic.** Make your text positive, polite, supportive, and encouraging. Emphasize what users can accomplish&mdash;not what they can't.
- **Know your customers.** Be mindful of cultural considerations and globalization when you use idioms or colloquialisms.

## Guidelines for UI text

The following guidelines apply to specific text elements you'll write for your add-in.

### Button labels and commands

Write button labels as short verb phrases that describe the action. Lead with a verb when possible.

| Instead of this | Use this |
|---|---|
| OK | Save |
| Click here to submit | Submit |
| New blank document | Create document |

### Error messages

When something goes wrong, tell users what happened, why, and what they can do next. Avoid blaming the user or using technical jargon.

| Instead of this | Use this |
|---|---|
| Invalid ID | You need an ID that looks like this: someone@example.com |
| Error 0x80004005 | We couldn't connect to the server. Check your internet connection and try again. |
| Operation failed | We couldn't save your changes. Try again in a few moments. |

### Descriptions and helper text

Helper text should answer the question "What does this do?" in a single sentence. Front-load the key information so users can scan quickly.

| Instead of this | Use this |
|---|---|
| This feature provides you with the ability to format selected text using pre-configured templates. | Format selected text with a template. |
| By clicking this button you will initiate the process of importing data from your external source. | Import data from an external source. |

### Progress and status messages

Use present tense for ongoing actions and past tense for completed ones. Be specific about what's happening.

## See also

- [Accessibility guidelines for Office Add-ins](accessibility-guidelines.md)
- [Design guidelines for Office Add-ins](add-in-design.md)
- [First-run experience patterns](first-run-experience-patterns.md)
- [Writing for all abilities](/style-guide/accessibility/writing-all-abilities)
- [Top 10 tips for Microsoft style and voice](/style-guide/top-10-tips-style-voice)
- [Word choice](/style-guide/word-choice/)
- [Scannable content](/style-guide/scannable-content/)
- Office Add-in [validation policies](/legal/marketplace/certification-policies)
