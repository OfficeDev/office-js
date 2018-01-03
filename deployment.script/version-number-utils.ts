import { isString, isNil, isArray } from "lodash";
import * as semver from 'semver';

import { fetchAndThrowOnError, runNpmCommand } from "./util";

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
