> [!NOTE]
> If you are developing on a Mac, enclose the `{url}` in single quotation marks. Do *not* do this on Windows.

```command&nbsp;line
npm run start -- web --document {url}
```

The following are examples.

- `npm run start -- web --document https://contoso.sharepoint.com/:t:/g/EZGxP7ksiE5DuxvY638G798BpuhwluxCMfF1WZQj3VYhYQ?e=F4QM1R`
- `npm run start -- web --document https://1drv.ms/x/s!jkcH7spkM4EGgcZUgqthk4IK3NOypVw?e=Z6G1qp`
- `npm run start -- web --document https://contoso-my.sharepoint-df.com/:t:/p/user/EQda453DNTpFnl1bFPhOVR0BwlrzetbXvnaRYii2lDr_oQ?e=RSccmNP`

If your add-in doesn't sideload in the document, manually sideload it by following the instructions in [Manually sideload add-ins to Office on the web](../testing/sideload-office-add-ins-for-testing.md#manually-sideload-an-add-in-to-office-on-the-web).
