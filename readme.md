# Office JavaScript APIs

The JavaScript API for Office enables you to create web applications that interact with the object models in Office host applications. Your application will reference the office.js library, which is a script loader. The office.js library loads the object models that are applicable to the Office application that is running the add-in.

The NPM package for Office.js is a copy of what gets published to the official "evergreen" Office.js CDN, at **<https://appsforoffice.microsoft.com/lib/1/hosted/office.js>**.  The NPM also offers alpha and beta versions for faster-cadence beta-testing (relative to the slower-cadence [official BETA endpoint](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)).


## Installing

To install "office-js" locally via the NPM package, run

    npm install @microsoft/office-js --save

Once installed, the Office.js script reference can be used as

    <src script="node_modules\@microsoft\office-js\dist\office.js"></script>


## Accessing the NPM files via a CDN

In addition to downloading the files locally, you can also use them via an external service like <https://unpkg.com>, which provides best-effort (no uptime guarantees) CDN hosting for npm packages.  This is especially useful for trying out alpha or beta builds.  To do so, simply change the script reference to:

    <src script="https://unpkg.com/@microsoft/office-js/dist/office.js"></script>

You can see the different versions of the NPM package listed in the dropdown on the top right at <https://unpkg.com/@microsoft/office-js/>.  This provides the alpha and beta versions as well.

To view the latest version numbers for each of the tags, you can also run the following command on the command-line:

    npm view @microsoft/office-js dist-tags --json


## More info

For more information on Office Add-ins and the Office JavaScript APIs, see:

- [Office Add-ins platform overview](https://dev.office.com/docs/add-ins/overview/office-add-ins)
- [JavaScript API for Office reference](https://dev.office.com/reference/add-ins/javascript-api-for-office)
