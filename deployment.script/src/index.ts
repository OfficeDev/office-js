import * as environment from "./EnvironmentVariables";
import { isString, isUndefined } from "util";
import * as semver from "semver";
import { SemVer, coerce } from "semver";
import {executeCommand} from "./executeCommand"
import { relative } from "path";
import {findLatestNpmPackageVersion} from "./findLatestNpmPackageVersion";
import * as getNextNpmPackageVersion from "./getNextNpmPackageVersion";
import * as standardFile from "./standardFile";
import * as path from "path";
import * as fs from "fs";

const env: environment.EnvironmentVariables = environment.getEnvironmentVariables();

// Print environment variables

const fieldsToPrint: (keyof environment.EnvironmentVariables)[] = [
    "TRAVIS",
    "TRAVIS_BRANCH",
    //"TRAVIS_BUILD_ID",
    //"TRAVIS_BUILD_NUMBER",
    //"TRAVIS_COMMIT_MESSAGE",
    "TRAVIS_PULL_REQUEST",

    // TRAVIS_BUILD_DIR Intentionally left out, since it serves no use to see, and causes issues if you copy-paste the output of these Travis parameters from the log into "launch.json"
];

const fields = fieldsToPrint.map(item => `"${item}": "${env[item]}"`).join(",\n");
console.log(fields);



enum ReleaseType {
    release = "release",
    beta = "beta",
    custom = "custom",
    none = "none"
}

// Base deployment action on the branch name
const DEPLOYMENT_BRANCH_NAME_RELEASE = ReleaseType.release as string;
const DEPLOYMENT_BRANCH_NAME_BETA =  ReleaseType.beta as string;
const DEPLOYMENT_BRANCH_NAME_CUSTOM_PREFIX =  ReleaseType.custom as string;

function getReleaseTypeFromBranchName(branchName: string): ReleaseType {

    if (branchName === DEPLOYMENT_BRANCH_NAME_RELEASE) {
        return ReleaseType.release;
    }

    if (branchName === DEPLOYMENT_BRANCH_NAME_BETA) {
        return ReleaseType.beta;
    }

    if (branchName.startsWith(DEPLOYMENT_BRANCH_NAME_CUSTOM_PREFIX)){
        return ReleaseType.custom;
    }

    return ReleaseType.none;
}

function deploymentPrerequisitesPassed(): boolean {
    let shouldDeploy = true;

    if (!process.env.TRAVIS) {
        console.log(`Deployment skipped - Not running inside of Travis.`);
        shouldDeploy = false;
    }
    
    // Do not run for pull requests.
    // Careful! Need this check to ensure a pull request does NOT trigger a deployment.
    if (process.env.TRAVIS_PULL_REQUEST !== 'false') {
        console.log(`Deployment skipped - Pull requests must NOT trigger a deployment.`);
        shouldDeploy = false;
    }

    // Only deploy from deployment branches
    if (getReleaseTypeFromBranchName(env.TRAVIS_BRANCH) === ReleaseType.none) {
        console.log(`Deployment skipped - Not a deployment branch.`);
        shouldDeploy = false;
    }

    if (env.NPM_TOKEN.length <= 0) {
        console.log(`Deployment skipped - [NPM_TOKEN] is a required global variables.`);
        shouldDeploy = false;
    }

    return shouldDeploy;
}

if (!deploymentPrerequisitesPassed()){
    // TODO: Actually exit
    // console.log("TODO: Actually skip deployment")
    process.exit(0);
}

/*
General Overview

The script runs on any commit to a branch

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

/**
 * The tag that should be applied to the package
 * @param release_type 
 */
function getNpmPackageTag(release_type: ReleaseType): string | undefined {
    if (release_type === ReleaseType.release) {
        return undefined;
    } else {
        return release_type as string;
    }
}


// Base actions on the branch name
const release_type = getReleaseTypeFromBranchName(env.TRAVIS_BRANCH);
const tag = getNpmPackageTag(release_type);


function updatePackageVersion(packageJsonPath: string, version: string) {
    type Package = {version: string};
    
    const packageData: Package = standardFile.readFileJson<Package>(packageJsonPath);

    packageData.version = version;

    standardFile.writeFileJson(packageJsonPath, packageData);
}

// update the package.json

/*
[env.TRAVIS_BUILD_DIR, path.join(env.TRAVIS_BUILD_DIR, "..")].forEach((dir: string)=> {

    console.log(`dir: [${dir}]`);

    if (dir === "" || dir === undefined || dir === "..") {

    } else if (fs.existsSync(dir) && standardFile.IsDirectory(dir)) {
        console.log(`subdirectories: [${dir}]`);
        console.log(standardFile.getSubDirectories(env.TRAVIS_BUILD_DIR));
        
        console.log(`files: [${dir}]`);
        console.log(standardFile.getFilesInDirectory(env.TRAVIS_BUILD_DIR));
    }

});
*/
interface AdditionalInfoError extends Error {
    additionalInfo: string;
}

function isPublishOverPreviouslyPublishVersionErrorString(errorText: string): boolean {
    const phraseMatchIndex = [
        "npm ERR! You cannot publish over the previously published version",
        "npm ERR! Cannot publish over previously published version"
    ]
        .map(phrase => phrase.toLowerCase())
        .findIndex(phrase => errorText.toLowerCase().indexOf(phrase) >= 0);
    return phraseMatchIndex >= 0;
}

function deployNpmPackage(packageDirectory: string, packageName: string, packageTag: string | undefined, npmAuthToken: string): void {
    console.log("Write .npmrc Deployment Token:")
    fs.writeFileSync(".npmrc", `//registry.npmjs.org/:_authToken=${env.NPM_TOKEN}`);
    
    const packageJsonPath = path.join(packageDirectory, "package.json");

    let remainingPublishAttempts = 1;
    let npmDeploymentSucceeded = false;

    while (!npmDeploymentSucceeded && remainingPublishAttempts > 0){
        remainingPublishAttempts -= 1;

        // dist folder is underneath
        const nextNpmPackageVersion = getNextNpmPackageVersion.getNextNpmPackageVersion("@microsoft/office-js", tag);
        console.log(`Package Version: [${nextNpmPackageVersion}]`);
        
        console.log("update package:");
        
        updatePackageVersion(packageJsonPath, nextNpmPackageVersion);
        
        console.log("Publish:")
        const tagParameter = tag === undefined ? "" : `--tag ${tag}`;

        try {
            executeCommand(`npm publish ${tagParameter}`, packageDirectory);
            npmDeploymentSucceeded = true;
        } catch (e) {
            const wasFailureDueToPreviouslyPublishedDeletedVersion =
                (e as AdditionalInfoError).additionalInfo &&
                isPublishOverPreviouslyPublishVersionErrorString((e as AdditionalInfoError).additionalInfo);

            if (wasFailureDueToPreviouslyPublishedDeletedVersion) {
                console.log(`Previous version was taken, trying again with an incremented version number...`);
            } else {
                throw e;
            }
        }
    }



}



const packageDirectory = env.TRAVIS_BUILD_DIR;
const packageName = "@microsoft/office-js";
const packageTag = tag;
const npmAuthToken = env.NPM_TOKEN;

console.log(`subdirectories: [${packageDirectory}]`);
console.log(standardFile.getSubDirectories(packageDirectory));

deployNpmPackage(packageDirectory, packageName, packageTag, npmAuthToken);


// view all versions

//console.log(data);

// find a similar version



//console.log(versions);

// find similar versions


// [undefined, "beta", "adhoc"].forEach((version: string | undefined) => {
//     console.log(`version: [${version}] [${findLatestNpmPackageVersion("@microsoft/office-js", version)}]`);
// });

// getNextNpmPackageVersion.test();

// function getNextPackageVersion(release_type: ReleaseType, tag?: string) {
//     // read package.json
// }


// read package.json



// What to figure out

// * What tag the npm package should have
// * What the version should be

// need to update the package.json appropriately

/*
semver.inc('1.2.3', 'prerelease', 'beta')
// '1.2.4-beta.0'



prerelease(v): Returns an array of prerelease components, or null if none exist. Example: prerelease('1.2.3-alpha.1') -> ['alpha', 1]
*/


/**
 * Figure out what to label the npm package and increment appropriately
 * 
 * release: x.y.z
 * 
 * beta: x.y.(z+1)-beta.q
 * 
 * custom: x.y.(z+1)-custom.p
 */

// let t: semver.SemVer = new semver.SemVer("1.1.1");
// console.log(t.format());

// t = t.inc('patch');
// console.log(t.format());

// t = t.inc('prerelease');
// console.log(t.format());

// t = t.inc("prerelease", "beta");
// console.log(t.format());

// t = t.inc("prerelease");
// console.log(t.format());

// const v = semver.coerce("1.1.3-beta.1");
// console.log(v === null ? "null" : v.format());



//semver.gte(basis, semver.coerce(tagVersion)

// pretty easy to increment last digit
// t= semver.inc('1.2.3-beta.0', 'prerelease') as string;
// console.log(t);
// t = semver.inc(t, 'prerelease') as string;
// console.log(t);

//function updatePackageJsonVersion(file: string, tag?: string, number: string)


//     "npm ERR! You cannot publish over the previously published version XYZ"
// error, we can revamp the version number and keep trying until we get it right.
/*
        fs.writeFileSync(".npmrc", `//registry.npmjs.org/:_authToken=${process.env.NPM_TOKEN}`);
        try {
            execCommand(`npm publish --tag ${npmPublishTag}`);
            npmDeploymentSuceeded = true;
        } catch (e) {
            const wasFailureDueToPreviouslyPublishedDeletedVersion =
                (e as AdditionalInfoError).additionalInfo &&
                isPublishOverPreviouslyPublishVersionErrorString((e as AdditionalInfoError).additionalInfo);

            if (wasFailureDueToPreviouslyPublishedDeletedVersion) {
                version = VersionUtils.incrementLastNumber(version, npmPublishTag);
                console.log(chalk.cyan.bold(`Previous version was taken, trying again with an incremented version number...`));
            } else {
                throw e;
            }
        }


*/
