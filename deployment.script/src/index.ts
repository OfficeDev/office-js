import * as environment from "./EnvironmentVariables";
import { isString, isUndefined } from "util";
import * as semver from "semver";
import { SemVer } from "semver";
import {executeCommand} from "./executeCommand"
import { relative } from "path";
import {findLatestNpmPackageVersion} from "./findLatestNpmPackageVersion";

const env: environment.EnvironmentVariables = environment.getEnvironmentVariables();

// Print environment variables

const fieldsToPrint: (keyof environment.EnvironmentVariables)[] = [
    "TRAVIS",
    "TRAVIS_BRANCH",
    "TRAVIS_BUILD_ID",
    "TRAVIS_BUILD_NUMBER",
    "TRAVIS_COMMIT_MESSAGE",
    "TRAVIS_PULL_REQUEST",

    // TRAVIS_BUILD_DIR Intentionally left out, since it serves no use to see, and causes issues if you copy-paste the output of these Travis parameters from the log into "launch.json"
];


const fields = fieldsToPrint.map(item => `"${item}": "${process.env[item]}"`).join(",\n");
console.log(fields);

// Base deployment action on the branch name
const DEPLOYMENT_BRANCH_NAME_RELEASE = "release";
const DEPLOYMENT_BRANCH_NAME_BETA = "beta";
const DEPLOYMENT_BRANCH_NAME_CUSTOM_PREFIX = "custom";

enum ReleaseType {
    release = "release",
    beta = "beta",
    custom = "custom",
    none = "none"
}

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


    const REQUIRED_ADDITIONAL_FIELDS: Array<keyof environment.EnvironmentVariables> = ["GH_TOKEN"];

    REQUIRED_ADDITIONAL_FIELDS.forEach(key => {
        if (!isString(process.env[key]) || (process.env[key] as string).trim().length <= 0) {
            console.log(`Deployment skipped - [${key}] is a required global variables.`);
            shouldDeploy = false;
        }
    });

    return shouldDeploy;
}

if (!deploymentPrerequisitesPassed()){
    // TODO: Actually exit
    console.log("TODO: Actually skip deployment")
    // process.exit(0);
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




// Base actions on the branch name
const release_type = getReleaseTypeFromBranchName(env.TRAVIS_BRANCH);

// view all versions

//console.log(data);

// find a similar version



//console.log(versions);

// find similar versions


[undefined, "beta", "adhoc"].forEach((version: string | undefined) => {
    console.log(`version: [${version}] [${findLatestNpmPackageVersion("@microsoft/office-js", version)}]`);
});


function getNextPackageVersion(release_type: ReleaseType, tag?: string) {
    // read package.json
}


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

let t: semver.SemVer = new semver.SemVer("1.1.1");
console.log(t.format());

t = t.inc('patch');
console.log(t.format());

t = t.inc('prerelease');
console.log(t.format());

t = t.inc("prerelease", "beta");
console.log(t.format());

t = t.inc("prerelease");
console.log(t.format());

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
