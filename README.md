[![NPM Deployment Status](https://travis-ci.org/OfficeDev/office-js.svg?branch=release)](https://travis-ci.org/OfficeDev/office-js/builds)

# Office JavaScript APIs

The JavaScript API for Office enables you to create web applications that interact with the object models in Office host applications. Your application will reference the office.js library, which is a script loader. The office.js library loads the object models applicable to the Office application running the add-in.

<br />

## About the NPM package

The NPM package for Office.js is a copy of what gets published to the official "evergreen" Office.js CDN, at **<https://appsforoffice.microsoft.com/lib/1/hosted/office.js>**.

The Office.js CDN contains all currently available Office.js APIs at any moment in time.

Each Office.js NPM package contains only Office.js APIs available on the Office.js CDN when the NPM package version was created.

### Target scenarios

The NPM package for Office.js is intended as a way to obtain an offline copy (non-CDN) of the Office.js files, which can then be statically serve from your own site instead of using the CDN.

This NPM package is provided for the following scenarios.

1. Development of an add-in behind a firewall, when accessing the Office.js CDN isn't possible.
2. Offline access to the Office.js APIs to facilitate offline debugging.

**Important**: Office Add-ins published to AppSource must use the CDN reference. Offline Office.js references are only appropriate for internal development and debugging scenarios.

### Best practices

Best practices for using the Office.js NPM package include:

- Refreshing the NPM package periodically to access new APIs and bug fixes.

- [Using the NPM package according to the instructions](#use-the-npm-package). Do not import the NPM package as commonly done with other NPM packages.

- Do not use the NPM package in an add-in intended for publication to [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office). Add-ins that are published to AppSource must use the Office.js CDN.

- [Using TypeScript definitions for Office.js](#intellisense-definitions).

### Support

The Office.js CDN is the official supported source for Office Add-ins.

For the NPM package sourced through this repository, only the latest version of the package is supported. No support and no patches will be provided for previous versions of the package. The frequency of the updates to this repository and related NPM package to match the CDN version is not guaranteed.

Outlook add-ins don't support hosting Office.js offline due to network access requirements for dependencies like the Microsoft Ajax library.

<br />

## Install the NPM package

To install "office-js" locally via the NPM package, run the following command.

> npm install @microsoft/office-js --save

<br />

Our policy requires that developers always reference the latest version of Office.js library. This is to ensure that essential product updates are served quickly to the Office Add-ins by always referencing the latest release of the library for a given version, such as Generally Available (GA) version. Given that the latest Office.js release is backward-compatible with prior releases, it's safe to update to the most recent release of the library when one is available. Hence, only the latest version of Office.js NPM package is made available for installation.

## Use the NPM package

Installing the NPM package locally creates a set of static Office.js files in the `node_modules\@microsoft\office-js\dist` folder of the directory where you ran the `npm install` command. To use the NPM package, do the following:

1. Either manually or as part of a build script (e.g., `CopyWebpackPlugin` if you're using Webpack) have the files served from a destination of your choosing (e.g., from the `/assets/office-js/` directory of your web server).

2. Reference that location in a `<script>` tag within the HTML file in your add-in project.

For example, if you add the contents of the `dist` folder to the `assets/office-js` directory of your project, then you'd add the following `<script>` tag to your HTML file:

```html
<script src="/assets/office-js/office.js"></script>
```

<br />

## IntelliSense definitions

TypeScript definitions for Office.js are available.

- latest **RELEASE** version of Office.js:
  - DefinitelyTyped: <https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-js/index.d.ts>
  - @types: `@types/office-js`
  > npm install @types/office-js --save-dev
- latest **BETA** version of Office.js:
  - DefinitelyTyped: <https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-js-preview/index.d.ts>
  - @types: `@types/office-js-preview`.
  > npm install @types/office-js-preview --save-dev
- **any** version:
  - Inside of the NPM package, under `dist/office.d.ts`
  - In this repo: [dist/office.d.ts](dist/office.d.ts)

### Use TypeScript definitions with the NPM package

If you're using the Office.js NPM package for the [firewall scenario](#target-scenarios) and want a d.ts file that precisely matches the JS contents, use the d.ts file that's located within the `/dist/office.d.ts` folder of the NPM package. Achieve this by using a [triple-slash reference](https://www.typescriptlang.org/docs/handbook/triple-slash-directives.html).

- **TIP**: If you create a `references.ts` file at the root of the project, simply point the reference to `office.d.ts` there.

If you don't need the [firewall scenario](#target-scenarios), don't use the Office.js NPM package. Obtain the TypeScript definitions by using `@types/office-js` and reference the Office.js CDN at <https://appsforoffice.microsoft.com/lib/1/hosted/office.js>.

### Enable IntelliSense in Visual Studio

Visual Studio 2017+ can use these same TypeScript definitions, even for regular JavaScript. For JavaScript IntelliSense in earlier versions of Visual Studio, an `office-vsdoc.js` is available alongside the `office.js` file. As long as you have a `Scripts/_references.js` file in your VS project, and as long as you substitute the existing triple-slash reference (`/// <reference path="https://.../office.js" />`) with the new location (the `-vsdoc` part gets substituted automatically, so use it just as you would in a `<script src="">` reference), you should have the corresponding JavaScript IntelliSense.

## NPM Package Versions

Office.js versioning is described in detail in <https://learn.microsoft.com/office/dev/add-ins/develop/referencing-the-javascript-api-for-office-library-from-its-cdn>. Importantly, there's a difference between what's in the JS files versus what are the capabilities of a particular computer (i.e., older or slower-to-update versions of Office).

The NPM package and the repo branches assume the following structure.

| GitHub branch name    | NPM tag name  | Description |
| ------------------    |-------------- |-------------|
| `release`             | `latest`      | Identical to a previous release on <https://appsforoffice.microsoft.com/lib/1/hosted/office.js> <br/><br/> The latest released publicly-available APIs.   |
| `beta`                | `beta`        | Identical to a previous release on <https://appsforoffice.microsoft.com/lib/beta/hosted/office.js> <br/><br/>  Forthcoming APIs, not necessarily ready for public consumption that may change. Possibly available on [Insider Fast (and maybe Insider Slow) builds](https://products.office.com/office-insider). |

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## More info

For more information on Office Add-ins and the Office JavaScript APIs, see:

- [Office Add-ins platform overview](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [JavaScript API for Office reference](https://learn.microsoft.com/javascript/api/overview)

## Join the Microsoft 365 Developer Program

Get a free sandbox, tools, and other resources you need to build solutions for the Microsoft 365 platform.

- [Free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Get a free, renewable 90-day Microsoft 365 E5 developer subscription.
- [Sample data packs](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Automatically configure your sandbox by installing user data and content to help you build your solutions.
- [Access to experts](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Access community events to learn from Microsoft 365 experts.
- [Personalized recommendations](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Find developer resources quickly from your personalized dashboard.
