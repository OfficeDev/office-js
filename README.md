[![NPM Deployment Status](https://travis-ci.org/OfficeDev/office-js.svg?branch=release)](https://travis-ci.org/OfficeDev/office-js/builds)

# Office JavaScript APIs

The JavaScript API for Office enables you to create web applications that interact with the object models in Office host applications. Your application will reference the office.js library, which is a script loader. The office.js library loads the object models that are applicable to the Office application that is running the add-in.

<br />

## About the NPM package

The NPM package for Office.js is a copy of what gets published to the official "evergreen" Office.js CDN, at **<https://appsforoffice.microsoft.com/lib/1/hosted/office.js>**.

While the Office.js CDN contains all currently available Office.js APIs at any moment in time, each version of the NPM package for Office.js contains only the Office.js APIs that were available at the point in time when that version of the NPM package was created.

### Target scenarios

The NPM package for Office.js is intended as a way for you to obtain offline copy (non-CDN) of the Office.js files, which you can then statically serve from your own site instead of using the CDN. This NPM package is primarily provided to address the following scenarios:

1. If you are developing an add-in behind a firewall, where accessing the Office.js CDN is not possible.
2. If you need offline access to the Office.js APIs (for example, to facilitate offline debugging).

If you plan to publish your Office Add-in from AppSource, you must use the CDN reference. Offlince Office.js references are only appropriate for internal, development, and debugging scenarios. 

### Best practices

Best practices for using the Office.js NPM package include:

- Refresh your NPM package periodically (to gain access to new APIs and/or bug fixes that may not have been available in your current version of the package).

- Use the NPM package according to the instructions in [Using the NPM package](#using-the-npm-package); do not try to import the NPM package as you might commonly do with other NPM packages.

- Do not use the NPM package in an add-in that you submit for publication to [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office). Add-ins that are published to AppSource must use the Office.js CDN.

- Use TypeScript definitions for Office.js as described in [IntelliSense definitions](#intellisense-definitions).

### Support

The Office.js CDN is the official, supported source for Office Add-ins. For the NPM package sourced through this repository, only the latest version of the package is supported. We'll not support nor apply any patches to previous major/minor versions of the package. In addition, the following are guaranteed for the NPM package: 

- Turnaround time for resolution of issues and bugs. 
- Frequency of updates to match the CDN version. 

Also, Outlook add-ins does not support hosting Office.js offline due to a network access requirement for dependencies like the Microsoft Ajax library.

<br />

## Installing the NPM package

To install "office-js" locally via the NPM package, run the following command:

    npm install @microsoft/office-js --save

<br />

Our policy requires that developers always reference the latest version of Office.js library. This is done to ensure that essential product updates are served quickly to the Office add-ins by always referencing the latest release of the library for a given version, such as Generally Available (GA) version. Given that the latest Office.js release is backward-compatible with prior releases, it is safe to update to the most recent release of the library when one is available. Hence, only the latest version of Office.js NPM package is made available for installation. 

## Using the NPM package

Installing the NPM package locally creates a set of static Office.js files in the `node_modules\@microsoft\office-js\dist` folder of the directory where you ran the `npm install` command. To use the NPM package, do the following:

1. Either manually or as part of a build script (e.g., `CopyWebpackPlugin` if you're using Webpack) have the files served from a destination of your choosing (e.g., from the `/assets/office-js/` directory of your web server).

2. Reference that location in a `<script>` tag within the HTML file in your add-in project.

For example, if you add the contents of the `dist` folder to the `assets/office-js` directory of your project, then you'd add the following `<script>` tag to your HTML file:

    <script src="/assets/office-js/office.js"></script>

<br />

## IntelliSense definitions

TypeScript definitions for Office.js are available.

* For latest **RELEASE** version of Office.js:
  * DefinitelyTyped: <https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-js/index.d.ts>
  * @types: `@types/office-js`.  (Acquire as `npm install @types/office-js --save-dev`)

* For **any version** (including **RELEASE**, but also including **BETA**, etc.):
  * Inside of the NPM package, under `dist/office.d.ts`
  * In this repo: [dist/office.d.ts](dist/office.d.ts)

### Using TypeScript definitions with the NPM package

1. If you are using the Office.js NPM package for the [firewall scenario](#target-scenarios) and want a d.ts file that precisely matches the JS contents, use the d.ts file that is located within the `/dist/office.d.ts` folder of the NPM package. You can achieve this by using a [triple-slash reference](https://www.typescriptlang.org/docs/handbook/triple-slash-directives.html). 

   - **Tip**: If you create a `references.ts` file at the root of the project, you can simply point the reference to `office.d.ts` there.

2. If you are using the Office.js NPM package for **beta**, follow the guidance outlined in the preceding point (#1), but make sure to update often.

If neither of these points applies to your scenario, you can just obtain the TypeScript definitions by using `@types/office-js` and reference the Office.js CDN at <https://appsforoffice.microsoft.com/lib/1/hosted/office.js> -- in which case, you don't need to use the Office.js NPM package.

### Enabling IntelliSense in Visual Studio

Visual Studio 2017+ can use these same TypeScript definitions, even for regular JavaScript. For JavaScript IntelliSense in earlier versions of Visual Studio, an `office-vsdoc.js` is available alongside the `office.js` file. As long as you have a `Scripts/_references.js` file in your VS project, and as long as you substitute the existing triple-slash reference (`/// <reference path="https://.../office.js" />`) with the new location (the `-vsdoc` part gets substituted automatically, so use it just as you would in a `<script src="">` reference), you should have the corresponding JavaScript IntelliSense.

<br />

## Accessing the NPM files via a CDN

In addition to downloading the files locally, you can also use them via an external service like <https://unpkg.com>, which provides best-effort (no uptime guarantees) CDN hosting for npm packages.  This is especially useful for trying out alpha or beta builds.  To do so, simply change the script reference to:

    <script src="https://unpkg.com/@microsoft/office-js/dist/office.js"></script>

You can see the different versions of the NPM package listed in the dropdown on the top right at <https://unpkg.com/@microsoft/office-js/>.  This provides the alpha and beta versions as well.

To view the latest version numbers for each of the tags, you can also run the following command on the command-line:

    npm view @microsoft/office-js dist-tags --json

When you have a version number, can use it as follows with <https://unpkg.com>: (appending `@<version-#>` right after `office-js`; e.g., `.../office-js@1.1.2-alpha.0/dist/...`

    <script src="https://unpkg.com/@microsoft/office-js@1.1.2-alpha.0/dist/office.js"></script>


<br />

## Production vs. Beta vs. Private versions

Office.js versioning is described in detail in <https://docs.microsoft.com/office/dev/add-ins/develop/referencing-the-javascript-api-for-office-library-from-its-cdn>.  Importantly, there is a large difference between what is in the JS files, versus what are the capabilities of a particular computer (i.e., older or slower-to-update versions of office).

The NPM package and the repo branches assume the following structure.

| GitHub branch name | NPM tag name | Description |
| ------------------ |--------------|-------------|
| `release` | `release` (and also `latest`, [a default NPM tag](https://docs.npmjs.com/getting-started/using-tags)) | The latest of the released publicly-available APIs. <br/> This should be identical with what is currently on <https://appsforoffice.microsoft.com/lib/1/hosted/office.js> |
| `beta`   | `beta` |  Forthcoming APIs, not necessarily ready for public consumption yet (and may still change...), but likely available on [Insider Fast (and maybe Insider Slow) builds](https://products.office.com/office-insider).  <br/> This should be identical to what is currently on <https://appsforoffice.microsoft.com/lib/beta/hosted/office.js> |
| `release-next` | `release-next` | A forthcoming update the the "release" branch (typically a couple weeks ahead of "release") |
| `beta-next` | `beta-next` | A forthcoming update the the "beta" branch (typically a couple weeks ahead of "beta") |
| `private` | `private` | Any flavor of a release, but deployed for a very specific need (e.g., try out something experimental) or for a specific partner. Unlike the other tags, successive versions of this tag are not necessarily cumulative updates; it is possible to have a `1.1.2-private.1` that has the beta JS, and then a `1.1.2-private.2` that only contains the publicly-available release APIs (with maybe some tweaks) |


<br />

## Using a Private or Beta version with [Script Lab](https://aka.ms/script-lab)

To use a version of the NPM package with [Script Lab](https://aka.ms/script-lab), substitute the CDN reference and the `@types/office-js` reference with the NPM package name and version.  [Note: Script Lab uses <https://unpkg.com> for resolving the package names, so it's very similar guidance as above].

For example, to use a `1.1.2-beta-next.0` version, use the following references:

    @microsoft/office-js@1.1.2-beta-next.0/dist/office.js
    @microsoft/office-js@1.1.2-beta-next.0/dist/office.d.ts


![Using the NPM package with Script Lab](https://github.com/OfficeDev/office-js/blob/release/.github/images/script-lab-substitute-refs.png)

<br />

## More info

For more information on Office Add-ins and the Office JavaScript APIs, see:

- [Office Add-ins platform overview](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [JavaScript API for Office reference](https://docs.microsoft.com/javascript/api/overview/office?view=office-js)
