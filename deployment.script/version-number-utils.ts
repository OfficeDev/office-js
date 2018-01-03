import { isString, isNil, isArray } from "lodash";
import * as semver from 'semver';
import * as handlebars from 'handlebars';
import * as moment from "moment";

import { fetchAndThrowOnError, runNpmCommand, stripSpaces } from "./util";

export interface IDeploymentYamlFileContext {
    version: string,
    travisBuildNumber: string,
    travisBuildId: string,
    tag: string
}


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

export function writeDeploymentYamlFile(partialContext: IDeploymentYamlFileContext) {
    console.log(generateDeploymentYamlText(partialContext));
}


function generateDeploymentYamlText(partialContext: IDeploymentYamlFileContext): string {
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
