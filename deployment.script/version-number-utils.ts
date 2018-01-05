import { isString, isNil, isArray } from "lodash";
import * as semver from "semver";
import * as handlebars from "handlebars";
import * as moment from "moment-timezone";
import * as fs from "fs-extra";

import { fetchAndThrowOnError, runNpmCommand, stripSpaces } from "./util";


export async function getPackageVersonStringFromRepo(
    branch: "release"
): Promise<string> {
    const url = `https://raw.githubusercontent.com/OfficeDev/office-js/${branch}/package.json`;
    const version = (await fetchAndThrowOnError<{ version: string }>(url, "json")).version;
    if (isNil(version) || !isString(version) || version.length <= 0) {
        throw new Error(`Missing or invalid package version number at URL "${url}"`);
    }

    return version;
}

export async function getNextPrivateVersionNumber() {
    const releaseVersionString = await getPackageVersonStringFromRepo("release");
    if (!semver.valid(releaseVersionString)) {
        throw new Error("Invalid release build number, should never happen");
    }

    const versionsResult = await runNpmCommand<any>("view", "@microsoft/office-js", "versions", "--json");
    if (Object.keys(versionsResult).length !== 1) {
        throw new Error("Unexpected result for versions");
    }

    const versionsArray: string[] = versionsResult[Object.keys(versionsResult)[0]]["versions"];
    if (!versionsArray || !isArray(versionsArray)) {
        throw new Error("Unexpected result for versions");
    }

    const privateNumStart = semver.inc(releaseVersionString, "patch")!;
    const matchingVersions = versionsArray
        .filter(item => item.startsWith(privateNumStart + "-private."));

    if (matchingVersions.length === 0) {
        return privateNumStart + "-private.0";
    }

    const largestNumber = Math.max(...matchingVersions.map(item => {
        let suffix = /(.*-private\.)(\d+)/.exec(item)![2];
        return Number.parseInt(suffix);
    }));

    return privateNumStart + "-private." + (largestNumber + 1);
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
    branchName: string,
    commitHash: string,
    commitMessage: string
}): string {
    const context = {
        ...partialContext,
        deployedAt: `${moment().utc().format('YYYY-MM-DD h:mm a')} UTC  (${moment().tz("America/Los_Angeles").format('YYYY-MM-DD h:mm a')} Pacific Time)`,
        isOfficialBuild: partialContext.npmPublishTag !== "private"
    };

    const template = stripSpaces(`
        version: {{{version}}}
        deployedAt: {{{deployedAt}}}

        history:
            privateBranchName: {{{branchName}}}
            basedOnDistFolderFromCommitHash: {{{commitHash}}}
            commitMessage: {{{commitMessage}}}

        unpkgUrls: |-
        {{#if isOfficialBuild}}
            builds using this same tag ("{{{tag}}}"):
                https://unpkg.com/@microsoft/office-js@{{{tag}}}/dist/office.js
                https://unpkg.com/@microsoft/office-js@{{{tag}}}/dist/office.debug.js  (unminified)
        {{/if}}
            this specific build number:
                https://unpkg.com/@microsoft/office-js@{{{version}}}/dist/office.js
                https://unpkg.com/@microsoft/office-js@{{{version}}}/dist/office.debug.js  (unminified)

        scriptLabReferences: |-
        {{#if isOfficialBuild}}
            builds using this same tag ("{{{tag}}}"):
                @microsoft/office-js@{{{tag}}}/dist/office.js
                @microsoft/office-js@{{{tag}}}/dist/office.d.ts
        {{/if}}
            this specific build number:
                @microsoft/office-js@{{{version}}}/dist/office.js
                @microsoft/office-js@{{{version}}}/dist/office.d.ts

        travisCI:
            buildNumber: {{{travisBuildNumber}}}
            buildId: {{{travisBuildId}}}
            log: https://travis-ci.org/OfficeDev/office-js/builds/{{{travisBuildId}}}
    `);

    return handlebars.compile(template)(context);
}
