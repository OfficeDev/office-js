[![NPM Deployment Status](https://travis-ci.org/OfficeDev/office-js.svg?branch=release)](https://travis-ci.org/OfficeDev/office-js/builds)

# Office JavaScript APIs

The JavaScript API for Office enables you to create web applications that interact with the object models in Office host applications. Your application will reference the office.js library, which is a script loader. The office.js library loads the object models that are applicable to the Office application that is running the add-in.

<br />

## About the NPM package

The NPM package for Office.js is a copy of what gets published to the official "evergreen" Office.js CDN, at **<https://appsforoffice.microsoft.com/lib/1/hosted/office.js>**.

While the Office.js CDN contains all currently available Office.js APIs at any moment in time, each version of the NPM package for Office.js contains only the Office.js APIs that were available at the point in time when that version of the NPM package was created.

### Target scenarios

The NPM package for Office.js is intended as a way for you to obtain an offline copy (non-CDN) of the Office.js files, which you can then statically serve from your own site instead of using the CDN. This NPM package is primarily provided to address the following scenarios.

1. If you are developing an add-in behind a firewall, where accessing the Office.js CDN is not possible.
2. If you need offline access to the Office.js APIs (for example, to facilitate offline debugging).

**Important**: If you plan to publish your Office Add-in from AppSource, you must use the CDN reference. Offline Office.js references are only appropriate for internal, development, and debugging scenarios.

### Best practices

Best practices for using the Office.js NPM package include:

- Refresh your NPM package periodically (to gain access to new APIs and/or bug fixes that may not have been available in your current version of the package).

- Use the NPM package according to the instructions in [Using the NPM package](#using-the-npm-package); do not try to import the NPM package as you might commonly do with other NPM packages.

- Do not use the NPM package in an add-in that you submit for publication to [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office). Add-ins that are published to AppSource must use the Office.js CDN.

- Use TypeScript definitions for Office.js as described in [IntelliSense definitions](#intellisense-definitions).

### Support

The Office.js CDN is the official, supported source for Office Add-ins. For the NPM package sourced through this repository, only the latest version of the package is supported. We'll not support nor apply any patches to previous versions of the package. In addition, the frequency of the updates to this repository and related NPM package to match the CDN version is not guaranteed.

Also, Outlook add-ins do not support hosting Office.js offline due to a network access requirement for dependencies like the Microsoft Ajax library.

<br />

## Installing the NPM package

To install "office-js" locally via the NPM package, run the following command:

> npm install @microsoft/office-js --save
<br />

Our policy requires that developers always reference the latest version of Office.js library. This is done to ensure that essential product updates are served quickly to the Office add-ins by always referencing the latest release of the library for a given version, such as Generally Available (GA) version. Given that the latest Office.js release is backward-compatible with prior releases, it is safe to update to the most recent release of the library when one is available. Hence, only the latest version of Office.js NPM package is made available for installation.

## Using the NPM package

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

- For latest **RELEASE** version of Office.js:
  - DefinitelyTyped: <https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-js/index.d.ts>
  - @types: `@types/office-js`.  (Acquire as `npm install @types/office-js --save-dev`)
- For latest **BETA** version of Office.js:
  - DefinitelyTyped: <https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-js-preview/index.d.ts>
  - @types: `@types/office-js-preview`.  (Acquire as `npm install @types/office-js-preview --save-dev`)
- For **any** version:
  - Inside of the NPM package, under `dist/office.d.ts`
  - In this repo: [dist/office.d.ts](dist/office.d.ts)

### Using TypeScript definitions with the NPM package

If you're using the Office.js NPM package for the [firewall scenario](#target-scenarios) and want a d.ts file that precisely matches the JS contents, use the d.ts file that is located within the `/dist/office.d.ts` folder of the NPM package. Achieve this by using a [triple-slash reference](https://www.typescriptlang.org/docs/handbook/triple-slash-directives.html). 

- **Tip**: If you create a `references.ts` file at the root of the project, you can simply point the reference to `office.d.ts` there.

If you don't need the firewall scenario, obtain the TypeScript definitions by using `@types/office-js` and reference the Office.js CDN at <https://appsforoffice.microsoft.com/lib/1/hosted/office.js>. You don't need to use the Office.js NPM package.

### Enabling IntelliSense in Visual Studio

Visual Studio 2017+ can use these same TypeScript definitions, even for regular JavaScript. For JavaScript IntelliSense in earlier versions of Visual Studio, an `office-vsdoc.js` is available alongside the `office.js` file. As long as you have a `Scripts/_references.js` file in your VS project, and as long as you substitute the existing triple-slash reference (`/// <reference path="https://.../office.js" />`) with the new location (the `-vsdoc` part gets substituted automatically, so use it just as you would in a `<script src="">` reference), you should have the corresponding JavaScript IntelliSense.

## Accessing the NPM files via a CDN

In addition to downloading the files locally, they may be used via an external service like <https://unpkg.com>, which provides best-effort (no uptime guarantees) CDN hosting for npm packages.  This is especially useful for trying out custom builds. To do so, simply change the script reference to:

```html
<script src="https://unpkg.com/@microsoft/office-js/dist/office.js"></script>
```

The different versions of the NPM package are listed  on the top right in the dropdown list at <https://unpkg.com/@microsoft/office-js/>.

To view the latest version numbers for each of the tags run the following command on the command-line:

> npm view @microsoft/office-js dist-tags --json

A specific version number, can be used with <https://unpkg.com> by appending `@<version-#>` right after `office-js`; e.g., `.../office-js@1.1.2-custom.0/dist/...`

```html
<script src="https://unpkg.com/@microsoft/office-js@1.1.2-custom.0/dist/office.js"></script>
```

<br />

## Production vs. Beta vs. Custom versions

Office.js versioning is described in detail in <https://docs.microsoft.com/office/dev/add-ins/develop/referencing-the-javascript-api-for-office-library-from-its-cdn>.  Importantly, there is a large difference between what is in the JS files, versus what are the capabilities of a particular computer (i.e., older or slower-to-update versions of office).

The NPM package and the repo branches assume the following structure.

| GitHub branch name    | NPM tag name  | Description |
| ------------------    |-------------- |-------------|
| `release`             | `latest`      | Identical to a previous release on <https://appsforoffice.microsoft.com/lib/1/hosted/office.js> <br/><br/> The latest released publicly-available APIs.   |
| `beta`                | `beta`        |Identical to a previous release on <https://appsforoffice.microsoft.com/lib/beta/hosted/office.js> <br/><br/>  Forthcoming APIs, not necessarily ready for public consumption that may change. Possibly available on [Insider Fast (and maybe Insider Slow) builds](https://products.office.com/office-insider). |
| `custom`              | `custom`      | A custom release deployed for a specific need. Successive versions of this tag are not cumulative updates (e.e. `1.1.2-custom.1` may contain beta JavaScript, and `1.1.2-custom.2` may only contain the publicly-available release APIs with some tweaks) |

## Using a specific version with [Script Lab](https://aka.ms/script-lab)

To use a version of the NPM package with [Script Lab](https://aka.ms/script-lab), substitute the CDN reference and the `@types/office-js` reference with the NPM package name and version.  [Note: Script Lab uses <https://unpkg.com> for resolving the package names, so it's very similar guidance as above].

For example, to use a `1.1.2-custom.0` version, use the following references:

```text
@microsoft/office-js@1.1.2-custom.0/dist/office.js
@microsoft/office-js@1.1.2-custom.0/dist/office.d.ts
```

![Using the NPM package with Script Lab](https://github.com/OfficeDev/office-js/blob/release/.github/images/script-lab-substitute-refs.png)

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## More info

For more information on Office Add-ins and the Office JavaScript APIs, see:

- [Office Add-ins platform overview](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [JavaScript API for Office reference](https://docs.microsoft.com/javascript/api/overview/office?view=office-js)

## Join the Microsoft 365 Developer Program

Get a free sandbox, tools, and other resources you need to build solutions for the Microsoft 365 platform.

- [Free developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Get a free, renewable 90-day Microsoft 365 E5 developer subscription.
- [Sample data packs](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Automatically configure your sandbox by installing user data and content to help you build your solutions.
- [Access to experts](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Access community events to learn from Microsoft 365 experts.
- [Personalized recommendations](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Find developer resources quickly from your personalized dashboard.
