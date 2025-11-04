# Office JavaScript APIs for Office Add-ins

Use Office.js and the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents and in Outlook mail messages and calendar items. With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to build solutions that can run in Office on the web, Windows, Mac, and mobile. You'll find walkthroughs, samples, and more information in the official [Office Add-ins documentation](https://learn.microsoft.com/office/dev/add-ins/).

## This repository

This GitHub repository is primarily used to report issues found in the Office JavaScript APIs. The npm package associated with this repository is no longer officially supported. Your add-in should get the JavaScript library from the Office content delivery network (CDN), as described in the section [Reference Office.js from the CDN](#reference-officejs-from-the-cdn). This ensures that essential product updates are served quickly to Office Add-ins.

## Report issues

If you believe you have found an issue (bug) with the Office JavaScript APIs, please visit the [issues tab](https://github.com/OfficeDev/office-js/issues) of this repo. If your issue is already reported, consider adding additional context or reproduction steps. Otherwise, select **New issue**, choose **Bug report**, and provide as much detail as possible. A member of our team will respond within 1-2 business days.

### Feature requests

Requests for new platform features should be raised at the [Microsoft 365 Developer Platform Tech Community site](https://aka.ms/m365dev-suggestions). Upvoted requests will be considered by our team for inclusion into the product.

### Other questions

Questions about developing add-ins and how to use the APIs should be raised on [Stack Overflow](https://stackoverflow.com/questions/tagged/office-js) with the "office-js" tag or on [Microsoft Q&A](https://learn.microsoft.com/answers/tags/321/office-development) with the "Office Development" tag. These locations are monitored by a community of experts, which includes members of our product team. They will review questions and provide assistance as they are able.

## Reference Office.js from the CDN

The Office CDN is the official supported source for Office Add-ins. Reference the Office.js library in the CDN by adding the following `<script>` tag within the `<head>` section of your HTML page.

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

This will download and cache the Office JavaScript API files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version. For more information, including how to reference preview APIs, see [Referencing the Office JavaScript API library](https://learn.microsoft.com/office/dev/add-ins/develop/referencing-the-javascript-api-for-office-library-from-its-cdn).

Government clouds may need to reference a special version of the CDN. For more information, see [Guidance for deploying Office Add-ins on government clouds](https://learn.microsoft.com/office/dev/add-ins/publish/government-cloud-guidance).

## IntelliSense definitions

TypeScript definitions for Office.js are available on DefinitelyTyped.

- latest **RELEASE** version of Office.js:
  - DefinitelyTyped: <https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-js/index.d.ts>
  - @types: `@types/office-js`
  > npm install @types/office-js --save-dev
- latest **PREVIEW** version of Office.js:
  - DefinitelyTyped: <https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-js-preview/index.d.ts>
  - @types: `@types/office-js-preview`
  > npm install @types/office-js-preview --save-dev

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any questions or comments.

## More information

For more information on Office Add-ins and the Office JavaScript APIs, see:

- [Office Add-ins platform overview](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- [JavaScript API for Office reference](https://learn.microsoft.com/javascript/api/overview)

## Join the Microsoft 365 Developer Program

Join the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram) to get resources and information to help you build solutions for the Microsoft 365 platform, including recommendations tailored to your areas of interest.

You might also qualify for a free developer subscription that's renewable for 90 days and comes configured with sample data; for details, see the [FAQ](https://learn.microsoft.com/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-).
