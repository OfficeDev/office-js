import * as environment from "./EnvironmentVariables";
import * as standardFile from "./standardFile";

import { EnvironmentVariables } from "./EnvironmentVariables";
import { debug, banner } from "./debug";
import { deploymentPrerequisitesPassed } from "./deploymentPrerequisitesPassed";
import { getReleaseTypeFromBranchName, ReleaseType } from "./ReleaseType";
import { getNpmPackageTag } from "./getNpmPackageTag";
import {deployNpmPackage} from "./deployNpmPackage";
import { isUndefined } from "util";

/*
General Overview:

The script runs on any commit to a branch.

Based on the branch name a different deployment is triggered.

Three types of packages:

    version: corresponds to the npm package version (needs to be set in the package.json)
    tag: corresponds to the npm package tag

release:
    version: x.y.z
    tag: release

beta:
    version: x.y.(z+1)-beta.q
    tag: beta

custom:
    version: x.y.(z+1)-custom.p
    tag: custom
*/
console.log("Deployment Script: Start");
const env: environment.EnvironmentVariables = environment.getEnvironmentVariables();

// Printing for debug purposes
// debug(env);

if (!deploymentPrerequisitesPassed(env)){
    process.exit(0);
}

// Base actions on the branch name
const release_type: ReleaseType = getReleaseTypeFromBranchName(env.TRAVIS_BRANCH);
const tag = getNpmPackageTag(release_type);

const packageDirectory = env.TRAVIS_BUILD_DIR;
const packageName = "@microsoft/office-js"; // could pull from the package.json
const packageTag = tag;
const npmAuthToken = env.NPM_TOKEN;

const deployedPackageVersion: string | undefined = deployNpmPackage(packageDirectory, packageName, packageTag, npmAuthToken);
console.log("Deployment Script: Complete");


const deploymentSucceeded = !isUndefined(deployedPackageVersion);
// report in an extra ovious way
console.log(banner(`DEPLOYMENT [${deploymentSucceeded ? "SUCCEEDED" : "FAILED"}]`))
if (deploymentSucceeded) {
console.log(`

Unpkg CDN URLs:
https://unpkg.com/@microsoft/office-js@${deployedPackageVersion}/dist/office.js
https://unpkg.com/@microsoft/office-js@${deployedPackageVersion}/dist/office.d.ts`
);

}
