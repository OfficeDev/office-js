#!/usr/bin/env node --harmony

import { isString } from "lodash";
import * as chalk from 'chalk';

import { banner, stripSpaces } from "./util";
import * as VersionUtils from "./version-number-utils";

declare var process: {
    env: IEnvironmentVariables
    exit: (status: number) => void;
};

const TRAVIS_AUTO_COMMIT_TEXT = "[TravisCI AUTO-COMMIT]";

const REQUIRED_ADDITIONAL_FIELDS: Array<keyof IEnvironmentVariables> =
    [];
//['GH_ACCOUNT', 'GH_REPO', 'GH_TOKEN'];

interface IEnvironmentVariables {
    TRAVIS: string,
    TRAVIS_BRANCH: string,
    TRAVIS_PULL_REQUEST: string,
    TRAVIS_COMMIT_MESSAGE: string,
    TRAVIS_BUILD_ID: string,
    TRAVIS_BUILD_NUMBER: string,
}

const OFFICIAL_BRANCHES = ["release", "release-next", "beta", "beta-next"];


(async () => {
    try {
        printBuildStartInfo();

        precheckOrExit();

        if (process.env.TRAVIS_BRANCH.startsWith("__private")) {
            await deployPrivateBuild();
        } else if (OFFICIAL_BRANCHES.indexOf(process.env.TRAVIS_BRANCH) >= 0) {
            await deployOfficialBranchBuild();
        } else {
            const message = stripSpaces(`
                Branch "${process.env.TRAVIS_BRANCH}" neither starts with "__private",
                    nor matches any of the following: [${
                OFFICIAL_BRANCHES.map(item => `"${item}"`).join(", ")
                }].
            `);
            banner('UNKNOWN BRANCH, SKIPPING DEPLOYMENT', message, chalk.yellow.bold);
        }

        banner('SUCCESS, DEPLOYMENT COMPLETE!', null, chalk.green.bold);
        process.exit(0);

    } catch (error) {
        banner('AN ERROR OCCURRED', error.message || error, chalk.bold.red);
        console.error(error);

        banner('DEPLOYMENT DID NOT GET TRIGGERED', null, chalk.bold.red);
        process.exit(1);
    }
})();

function printBuildStartInfo() {
    const fieldsToPrint: (keyof IEnvironmentVariables)[] = [
        "TRAVIS",
        "TRAVIS_BRANCH",
        "TRAVIS_BUILD_ID",
        "TRAVIS_BUILD_NUMBER",
        "TRAVIS_COMMIT_MESSAGE",
        "TRAVIS_PULL_REQUEST"
    ];

    const fieldsString = fieldsToPrint
        .map(item => item + ": " + process.env[item])
        .join("\n");

    banner('TravisCI build started', fieldsString, chalk.green.bold);
}

function precheckOrExit(): void {
    /* Check if the code is running inside of travis.ci. If not abort immediately. */
    if (!process.env.TRAVIS) {
        banner('Deployment skipped', 'Not running inside of Travis.', chalk.yellow.bold);
        process.exit(0);
    }

    if (process.env.TRAVIS_COMMIT_MESSAGE && process.env.TRAVIS_COMMIT_MESSAGE.startsWith(TRAVIS_AUTO_COMMIT_TEXT)) {
        banner('Deployment skipped',
            `Skipping builds for commit messages labeled as "${TRAVIS_AUTO_COMMIT_TEXT}"`,
            chalk.yellow.bold);
        process.exit(0);
    }

    // Careful! Need this check because otherwise, a pull request against master would immediately trigger a deployment.
    if (process.env.TRAVIS_PULL_REQUEST !== 'false') {
        banner('Deployment skipped', 'Skipping deploy for pull requests.', chalk.yellow.bold);
        process.exit(0);
    }

    REQUIRED_ADDITIONAL_FIELDS.forEach(key => {
        if (!isString(process.env[key]) || (process.env[key] as string).trim().length <= 0) {
            throw new Error(`"${key}" is a required global variables.`);
        }
    });
}

async function deployPrivateBuild(): Promise<void> {
    let version = await VersionUtils.getNextPrivateVersionNumber();
    console.log("Will be deploying build number " + version);

    VersionUtils.writeDeploymentYamlFile({
        tag: "private",
        travisBuildId: process.env.TRAVIS_BUILD_ID,
        travisBuildNumber: process.env.TRAVIS_BUILD_NUMBER,
        version
    });
}

async function deployOfficialBranchBuild(): Promise<void> {
    const message = stripSpaces(`
        Sorry, the auto-deployment of official branches isn't supported yet.
    `);
    banner('NOT YET IMPLEMENTED, SKIPPING DEPLOYMENT', message, chalk.yellow.bold);
}
