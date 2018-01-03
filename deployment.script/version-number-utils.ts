import { isString, isNil, isArray } from "lodash";
import * as semver from 'semver';
import * as handlebars from 'handlebars';
import * as moment from "moment";
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
    const versionRegex = /(^\s+"version": ")(.*)(",\s+$)/;
    let versionEntryIndex = packageJsonContentsArray.findIndex(line => versionRegex.test(line));
    if (versionEntryIndex <= 0) {
        throw new Error("Could not find a line with the package version number, this can't be correct");
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
    tag: string,
    branchName: string,
    commitHash: string,
    commitMessage: string
}): string {
    const context = {
        ...partialContext,
        deployedAt: moment().utc().format('YYYY-MM-DD HH:mm a') + ' UTC',
        isOfficialBuild: partialContext.tag !== "private"
    };

    const template = stripSpaces(`
        version: {{{version}}}
        deployedAt: {{{deployedAt}}}

        travisCI:
            buildNumber: {{{travisBuildNumber}}}
            buildId: {{{travisBuildId}}}
            log: https://travis-ci.org/OfficeDev/script-lab/builds/{{{travisBuildId}}}

        history:
            privateBranchName: {{{branchName}}}
            basedOnDistFolderFromCommitHash: {{{commitHash}}}
            commitMessage: {{{commitMessage}}}

        unpkgUrls: |-
        {{#if isOfficialBuild}}
            builds using this same tag ("{{{tag}}}"):
                https://unpkg.com/@microsoft/office-js@{{{tag}}}/office.js
        {{/if}}
            this specific build number:
                https://unpkg.com/@microsoft/office-js@{{{version}}}/office.js

        scriptLabReferences: |-
        {{#if isOfficialBuild}}
            builds using this same tag ("{{{tag}}}"):
                @microsoft/office-js@{{{tag}}}/office.js
                @microsoft/office-js@{{{tag}}}/office.d.ts
        {{/if}}
            this specific build number:
                @microsoft/office-js@{{{version}}}/office.js
                @microsoft/office-js@{{{version}}}/office.d.ts
    `);

    return handlebars.compile(template)(context);
}
