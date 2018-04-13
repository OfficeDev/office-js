import { isString, isNil, isArray } from "lodash";
import * as semver from "semver";
import * as handlebars from "handlebars";
import * as moment from "moment-timezone";
import * as fs from "fs-extra";
import * as jsyaml from "js-yaml";

import { fetchAndThrowOnError, runNpmCommand, stripSpaces } from "./util";

export async function getNextVersionNumber(npmTag: string): Promise<string> {
    const currentReleaseVersionString = await getReleasePackageVersonStringFromGitHub();

    if (npmTag === "release") {
        // For release, increment it not by one patch version, but by TWO.
        // That way, the release-next/beta/beta-next/adhoc versions don't get mixed up with the main one
        // (e.g., in a dropdown or alphabetical list.)

        const oneUp = semver.inc(currentReleaseVersionString, "patch")!;
        const twoUp = semver.inc(oneUp, "patch")!;
        return twoUp;

    } else {
        return await getNextSuffixedVersion(semver.inc(currentReleaseVersionString, "patch")!, npmTag);
    }
}

export function incrementLastNumber(version: string, npmPublishTag: string): string {
    const regex = /^(.*\.)(\d+)$/;
    /* Matches the last digit of all of the following:
        1.1.3-beta.1
        1.1.3-beta.1111
        1.1.4
    */
    const result = regex.exec(version)!;
    let lastNumber = Number.parseInt(result[2]);
    lastNumber = lastNumber + (npmPublishTag === "release" ? 2 : 1);

    return result[1] + lastNumber;
}

export function updatePackageJson(version: string): void {
    const packageJsonPath = "package.json";
    const packageJsonContentsArray = fs.readFileSync(packageJsonPath).toString().split("\n");
    const versionRegex = /^(\s+"version": ")(.*)(",\s*)$/;
    let versionEntryIndex = packageJsonContentsArray.findIndex(line => versionRegex.test(line));
    if (versionEntryIndex <= 0) {
        const errorMessage = "Could not find a line with the package version number, this can't be correct.";
        console.error(errorMessage);
        console.warn(packageJsonContentsArray.join("\n"));
        throw new Error(errorMessage);
    }
    const regexResult = versionRegex.exec(packageJsonContentsArray[versionEntryIndex])!;
    const substitutedVersion = regexResult[1] + version + regexResult[3];
    packageJsonContentsArray[versionEntryIndex] = substitutedVersion;
    fs.writeFileSync(packageJsonPath, packageJsonContentsArray.join("\n"));
}

export function generateDeploymentYamlText(partialContext: {
    version: string,
    travisBuildNumber: string,
    travisBuildId: string,
    npmPublishTag: string,
    includeTagUrls: boolean
    historyInfo: {}
}): string {
    const context = {
        ...partialContext,
        deployedAt: `${moment().utc().format('YYYY-MM-DD h:mm a')} UTC  (${moment().tz("America/Los_Angeles").format('YYYY-MM-DD h:mm a')} Pacific Time)`,
        historyBlockString: jsyaml.safeDump({ history: partialContext.historyInfo }, { indent: 4 })
    };

    const template = stripSpaces(`
        version: {{{version}}}
        tag: {{{npmPublishTag}}}
        githubReleaseUrl: https://github.com/OfficeDev/office-js/releases/tag/v{{{version}}}
        githubViewUrl: https://github.com/OfficeDev/office-js/tree/v{{{version}}}
        deployedAt: {{{deployedAt}}}

        {{{historyBlockString}}}

        unpkgUrls: |-
            this specific build number:
                https://unpkg.com/@microsoft/office-js@{{{version}}}/dist/office.js
                https://unpkg.com/@microsoft/office-js@{{{version}}}/dist/office.d.ts
        {{#if includeTagUrls}}
            builds using this same tag ("{{{npmPublishTag}}}"):
                https://unpkg.com/@microsoft/office-js@{{{npmPublishTag}}}/dist/office.js
                https://unpkg.com/@microsoft/office-js@{{{npmPublishTag}}}/dist/office.d.ts
        {{/if}}

        buildLog: https://travis-ci.org/OfficeDev/office-js/builds/{{{travisBuildId}}}
    `);

    return handlebars.compile(template)(context);
}

export function generateMarkdownDescription(context: {
    commitMessage: string;
    version: string;
    travisBuildId: string;
    npmPublishTag: string;
    includeTagUrls: boolean;
    NPM_DEPLOYMENT_INFO_YAML_FILENAME: string;
}): string {
    const template = stripSpaces(`
        ### From version "{{{commitMessage}}}"

        * Full version details & commit history: https://github.com/OfficeDev/office-js/blob/v{{{version}}}/{{{DEPLOYMENT_YAML_FILENAME}}}
        * Build log: https://travis-ci.org/OfficeDev/office-js/builds/{{{travisBuildId}}}

        <br />

        ### Unpkg CDN URLs
        > #### This specific build number:
        > * https://unpkg.com/@microsoft/office-js@{{{version}}}/dist/office.js
        > * https://unpkg.com/@microsoft/office-js@{{{version}}}/dist/office.d.ts
        {{#if includeTagUrls}}
        > #### Builds using this same tag ("{{{npmPublishTag}}}"):
        > https://unpkg.com/@microsoft/office-js@{{{npmPublishTag}}}/dist/office.js
        > https://unpkg.com/@microsoft/office-js@{{{npmPublishTag}}}/dist/office.d.ts
        {{/if}}

        ### IntelliSense

        > #### Script Lab:
        > * Copy-paste both of the above URLs (**office.js** *and* **office.d.ts**) into the **Libraries** tab.  That's it.
        > #### TypeScript-based project:
        > 1. \`npm install @microsoft/{{{version}}}\`  (for this specific build number)
        {{#if isOfficialBuild}}
        >         &nbsp;&nbsp;&nbsp;&nbsp; or
        >    \`npm install @microsoft/office-js@{{{npmPublishTag}}}\`  (for any version of the "npmPublishTag" tag)
        {{/if}}
        > 2. At the top of the TS file, add:
        > \`/// <reference path="./node_modules/@microsoft/office-js/dist/office.d.ts" />\`
        >
        > #### VS JavaScript-based project
        > * VS **2017**:  Download the **d.ts** file from above link and stick it inside your project.
        > * VS **2015** or earlier:  Under \`Scripts\\_references.js\`, replace the URL inside \`/// <reference path="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" />\` with the **.js** (*not* d.ts) URL above.
    `);

    return handlebars.compile(template)(context);
}


async function getReleasePackageVersonStringFromGitHub(): Promise<string> {
    const url = `https://raw.githubusercontent.com/OfficeDev/office-js/release/package.json`;
    const versionString = (await fetchAndThrowOnError<{ version: string }>(url, "json")).version;

    if (isNil(versionString) || !isString(versionString) || versionString.length <= 0) {
        throw new Error(`Missing or invalid package version number at URL "${url}"`);
    }

    if (!semver.valid(versionString)) {
        throw new Error("Invalid release build number, should never happen");
    }

    return versionString;
}

async function getNextSuffixedVersion(mainVersionNumberString: string, tagName: string) {
    const versionsResult = await runNpmCommand<any>("view", "@microsoft/office-js", "versions", "--json");
    if (Object.keys(versionsResult).length !== 1) {
        throw new Error("Unexpected result for versions");
    }

    const versionsArray: string[] = versionsResult[Object.keys(versionsResult)[0]]["versions"];
    if (!versionsArray || !isArray(versionsArray)) {
        throw new Error("Unexpected result for versions");
    }

    const matchingVersions = versionsArray
        .filter(item => item.startsWith(mainVersionNumberString + "-" + tagName + "."));

    if (matchingVersions.length === 0) {
        return `${mainVersionNumberString}-${tagName}.0`;
    }

    const suffixRegex = new RegExp(`(.*-${tagName}\\.)(\\d+)`);
    // Note: don't need to escape tagName, since it's going to be devoid of any special characters
    // (just letters and a dash, e.g., for "release-next"), and a dash is OK (at least in this context)
    const largestNumber = Math.max(...matchingVersions.map(item => {
        let suffix = suffixRegex.exec(item)![2];
        return Number.parseInt(suffix);
    }));

    return `${mainVersionNumberString}-${tagName}.${largestNumber + 1}`;
}
