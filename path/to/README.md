# Office JavaScript APIs for Office Add-ins

Use Office.js and the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents and in Outlook mail messages and calendar items. With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to build solutions that can run in Office on the web, Windows, Mac, and mobile. You'll find walkthroughs, samples, and more information in the official [Office Add-ins documentation](https://learn.microsoft.com/office/dev/add-ins/).

## This repository

This GitHub repository is primarily used to report issues found in the Office JavaScript APIs. The npm package associated with this repository is no longer officially supported. Your add-in should get the JavaScript library from the Office content delivery network (CDN), as described in the section [Reference Office.js from the CDN](#reference-officejs-from-the-cdn). This ensures that essential product updates are served quickly to Office Add-ins.

## Report issues

If you believe you've found an issue (bug) with the Office JavaScript APIs, please visit the [issues tab](https://github.com/OfficeDev/office-js/issues) of this repo. If your issue is already reported, consider adding additional context or reproduction steps. Otherwise, select **New issue**, choose **Bug report**, and provide as much detail as possible. A member of our team will respond within 1-2 business days.

## Reference Office.js from the CDN

The Office CDN is the official supported source for Office Add-ins. Reference the Office.js library in the CDN by adding the following `<script>` tag within the `<head>` section of your HTML page.