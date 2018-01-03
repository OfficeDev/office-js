#!/usr/bin/env node --harmony

import { isString } from "lodash";
import * as chalk from 'chalk';
import * as shell from 'shelljs';
import * as fs from "fs-extra";

import { banner, stripSpaces, execCommand } from "./util";
import * as VersionUtils from "./version-number-utils";

declare var process: {
    env: IEnvironmentVariables
    exit: (status: number) => void;
};

const TRAVIS_AUTO_COMMIT_TEXT = "[TRAVIS CI AUTO-COMMIT]";
const TOKENIZED_GITHUB_PUSH_URL = `https://<<<token>>>@github.com/OfficeDev/office-js.git`;
const REPO_LOCAL_FOLDER = "office-js"; // Just a local folder name relative to current path.
const DEPLOYMENT_YAML_FILENAME = "NPM.DEPLOYMENT.INFO.yaml";

const REQUIRED_ADDITIONAL_FIELDS: Array<keyof IEnvironmentVariables> = ['GH_TOKEN'];

interface IEnvironmentVariables {
    TRAVIS: string,
    TRAVIS_BRANCH: string,
    TRAVIS_PULL_REQUEST: string,
    TRAVIS_COMMIT: string,
    TRAVIS_COMMIT_MESSAGE: string,
    TRAVIS_BUILD_ID: string,
    TRAVIS_BUILD_NUMBER: string,

    /**
     * GitHub token generated using https://github.com/settings/tokens,
     *     bearing permissions for "repo:status", "repo_deployment", and "public_repo".
     * This is a personal access token, so the commits always happen on behalf
     *     of the person who created the token.
     * The token is then entered as a hidden value in https://travis-ci.org/OfficeDev/office-js/settings */
    GH_TOKEN: string
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
    const version = await VersionUtils.getNextPrivateVersionNumber();
    console.log("Will be deploying as office.js NPM version " + version);

    await cloneRepo({
        version,
        afterCloneBeforePush: async () => {
            const deploymentFileContents = VersionUtils.generateDeploymentYamlText({
                tag: "private",
                travisBuildId: process.env.TRAVIS_BUILD_ID,
                travisBuildNumber: process.env.TRAVIS_BUILD_NUMBER,
                branchName: process.env.TRAVIS_BRANCH,
                commitHash: process.env.TRAVIS_COMMIT,
                commitMessage: process.env.TRAVIS_COMMIT_MESSAGE,
                version
            });
            fs.writeFileSync(DEPLOYMENT_YAML_FILENAME, deploymentFileContents);
            execCommand(`git add ${DEPLOYMENT_YAML_FILENAME}`);

            VersionUtils.updatePackageJson(version);

            execCommand(`git commit --allow-empty -m "${TRAVIS_AUTO_COMMIT_TEXT}\n${process.env.TRAVIS_COMMIT_MESSAGE}"`);
        }
    });
}

async function deployOfficialBranchBuild(): Promise<void> {
    const message = stripSpaces(`
        Sorry, the auto-deployment of official branches isn't supported yet.
    `);
    banner('NOT YET IMPLEMENTED, SKIPPING DEPLOYMENT', message, chalk.yellow.bold);
}

async function cloneRepo(options: {
    version: string;
    afterCloneBeforePush: () => Promise<any>,
}) {
    fs.removeSync(REPO_LOCAL_FOLDER);

    execCommand(`git clone ${TOKENIZED_GITHUB_PUSH_URL} ${REPO_LOCAL_FOLDER}`, {
        token: process.env.GH_TOKEN
    });

    shell.pushd(REPO_LOCAL_FOLDER);


    execCommand(`git checkout ${process.env.TRAVIS_BRANCH}`);
    execCommand('git config --add user.name "Travis CI"');
    execCommand('git config --add user.email "travis.ci@microsoft.com"');

    await options.afterCloneBeforePush();

    execCommand(`git push`);

    const gitTagName = "v" + options.version;
    execCommand(`git tag -a ${gitTagName} -m "${TRAVIS_AUTO_COMMIT_TEXT}\n${process.env.TRAVIS_COMMIT_MESSAGE}"`);
    execCommand(`git push origin ${gitTagName}`);


    shell.popd();
}
