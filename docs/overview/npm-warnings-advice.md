---
title: Warnings and dependencies in the Node.js and npm world
description:  Learn about aspects of npm development that are disconcerting to COM and VSTO developers.
ms.date: 07/07/2025
ms.localizationpriority: medium
---

# Warnings and dependencies in the Node.js and npm world

COM and VSTO add-in developers who are new to the world of [Node Package Manager (npm)](https://www.npmjs.com/) and open source development are often surprised and alarmed at certain aspects of this kind of development. This article is intended to reassure such developers about things they may find disconcerting.

## npm dependency tree

npm is the standard package manager for the JavaScript runtime environment Node.js. It's used to streamline JavaScript development workflows by enabling developers to install open source libraries and tools (collectively called "packages"), and to manage package dependencies.

The npm dependency tree is a hierarchical structure that represents all the packages your Node.js project depends on. Each node in the tree is a package, and its children are the packages it depends on. This structure can become deeply nested, especially in large projects or when using packages with many transitive dependencies.

When you run `npm install`, npm reads the package.json and package-lock.json files in a development project to build this tree and fetch the required packages.

## Understand `npm install` warnings

When you run `npm install`, it's common to see a flurry of warnings in the console. This can be surprising at first, but it's a normal part of working in the Node.js and open source ecosystem. Microsoft tools that call `npm install`, including the [Yeoman Generator for Office Add-ins](../develop/yeoman-generator-overview.md), report these same warnings.

It's beyond the scope of this article to discuss every kind of warning that `npm install` might report, but there are two kinds that are especially likely to be disconcerting to developers who are new to the Node.js world.

### Deprecation warnings

These warnings mean that the managers of a package somewhere in the dependency tree are no longer maintaining it and they may remove it from the Internet at some future time. Neither you, nor Microsoft, has any control over the package deprecation warnings, but you can almost always ignore them. Deprecation doesn't mean that the package has stopped working. It still works and because installation puts a copy of it on your computer, it'll continue to work with your project in the future even if the package is removed from the Internet. The package isn't a web service.

It's very unlikely, but possible, that you'll see a deprecation warning for a package that's at the *top* of the dependency tree. These are the packages that are explicitly listed in the "dependencies" or "devDependencies" sections of the project's package.json file. You can ignore deprecation warnings for "devDependencies" for the same reason given earlier: the code is copied to your development computer. Packages in the "dependencies" section are used by your add-in at runtime, but even deprecation warnings for these can be ignored in projects that are created with Microsoft tools like the Yeoman Generator for Office Add-ins and [Microsoft 365 Agent Toolkit](../develop/agents-toolkit-overview.md) because these tools include copies of the libraries in bundles of JavaScript code that your add-in's web server will serve.

> [!NOTE]
> One situation in which the deprecation of a library in the "dependencies" section is a matter of concern is the following:
>
> - The library is in the "dependencies" section only so you can use it while testing and debugging locally.
> - Your plan, when you deploy the add-in for staging or production, is to not include a copy of the library in the code that your server hosts.
> - Instead, you plan to have the add-in load the library from a web service that hosts npm libraries, such as unpkg.com or cdn.jsdelivr.net.
>
> If this describes your deployment strategy, then there's a danger that your deployed add-in will stop working if the deprecated library is removed from the web service. So, treat the deprecation warning as a notice that you need to redesign your add-in so that it doesn't use the deprecated library.

### Security or audit warnings

Security warnings, also called audit warnings, mean that there's a version of the package in the dependency tree that has a known security vulnerability that a hacker could take advantage of. Microsoft periodically checks for these warnings in projects created by our tools and fixes them, usually by updating the library to a newer version that doesn't have the vulnerability. But new vulnerabilities are discovered and reported almost daily to the security alert databases that `npm install` monitors, and Microsoft cannot always fix them right away. For this reason, it isn't uncommon that running `npm install` in an add-in project reports security vulnerabilities.

When the dependency can be traced to a top-level package in the "devDependencies" section of package.json, then you can ignore it. The code is only running on your computer, and you're not going to hack yourself.

If the dependency traces to a top-level package in the "dependencies" section, or you cannot determine the top-level package, then you should try to fix the vulnerability before you deploy the add-in to production. There's lots of good information on the Internet about how to deal with vulnerabilities in npm packages. We'll mention one technique here. Some vulnerabilities can be fixed automatically by npm. Just run the command `npm audit fix` in the folder where the package.json file is. If there's a newer version of the package that doesn't have the vulnerability and the newer version doesn't have any breaking changes relative to the vulnerable version, then npm will automatically update the package to the safe version.

Another strategy is to take a few minutes every couple of weeks to create a new add-in project with the same Microsoft tool as you created your original project. (Choose the same options for project type,, language, Office application, etc.) If `npm install` no longer reports the security vulnerability on the new project, then Microsoft has fixed it in the project template. You can move the fix to your project with the following steps.

1. Copy the "dependencies" section of from the new project over the same section in the **package.json** of the original project.
1. Delete the **node_modules** folder from the original project.
1. Run `npm install` in the original project.

## Errors

An npm *error*, as distinct from a warning, immediately stops the processing of the npm command, including `npm install`. You must diagnose and fix the problem. Sometimes the error is a side effect of a temporary network problem when npm tries to fetch a package. Try rerunning `npm install`.

> [!NOTE]
> Running `npm install` is the last thing that the Yeoman Generator for Office Add-ins does when it creates a project. If an error is reported, you don't need to rerun the generator because the project has been fully created. You can just rerun `npm install` at the command line.
