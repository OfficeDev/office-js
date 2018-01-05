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
    TRAVIS_BUILD_DIR: string,

    /**
     * GitHub token generated using https://github.com/settings/tokens,
     *     bearing permissions for "repo:status", "repo_deployment", and "public_repo".
     * This is a personal access token, so the commits always happen on behalf
     *     of the person who created the token.
     * The token is then entered as a hidden value in https://travis-ci.org/OfficeDev/office-js/settings */
    GH_TOKEN: string,

    /** A token for publishing to NPM.  It can be generated using "npm token create"
     * Note that you'll need NPM version 5.5.1+ to run this command.
     * https://docs.npmjs.com/getting-started/working_with_tokens
    */
    NPM_TOKEN: string
}

const OFFICIAL_BRANCHES = ["release", "release-next", "beta", "beta-next"];

interface IDeploymentParams {
    npmPublishTag: string;
    version: string;
    afterCloneBeforeCommit?: () => Promise<any>;
}

(async () => {
    try {
        printBuildStartInfo();

        precheckOrExit();

        let deploymentParams: IDeploymentParams;
        if (process.env.TRAVIS_BRANCH.startsWith("__private")) {
            deploymentParams = await getPrivateBranchDeploymentParams();
        } else if (OFFICIAL_BRANCHES.indexOf(process.env.TRAVIS_BRANCH) >= 0) {
            deploymentParams = await getOfficialBranchDeploymentParams();
        } else {
            const message = stripSpaces(`
                Branch "${process.env.TRAVIS_BRANCH}" neither starts with "__private",
                    nor matches any of the following: [${
                OFFICIAL_BRANCHES.map(item => `"${item}"`).join(", ")
                }].
            `);
            banner('UNKNOWN BRANCH, SKIPPING DEPLOYMENT', message, chalk.yellow.bold);
            process.exit(0);
            return;
        }

        let deploymentResultOutput = await doDeployment(deploymentParams);

        banner('SUCCESS, DEPLOYMENT COMPLETE!', deploymentResultOutput, chalk.green.bold);
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
        "TRAVIS_PULL_REQUEST",

        // "TRAVIS_BUILD_DIR":  Intentionally *NOT* outputting it,
        // since it serves no use to see, but causes issues if you copy-paste
        // the output of these Travis parameters from the log into "launch.json"
    ];

    const fieldsString = fieldsToPrint
        .map(item => `"${item}": "${process.env[item]}"`)
        .join(",\n");

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

async function doDeployment(params: IDeploymentParams): Promise<string> {
    const { version, npmPublishTag } = params;
    const gitTagName = "v" + params.version;

    banner("This deployment's target NPM version", "Target package version: " + version, chalk.magenta.bold);

    const deploymentFileContents = VersionUtils.generateDeploymentYamlText({
        npmPublishTag,
        version,
        travisBuildId: process.env.TRAVIS_BUILD_ID,
        travisBuildNumber: process.env.TRAVIS_BUILD_NUMBER,
        branchName: process.env.TRAVIS_BRANCH,
        commitHash: process.env.TRAVIS_COMMIT,
        commitMessage: process.env.TRAVIS_COMMIT_MESSAGE,
    });

    const repoLocalFolderPath = process.env.TRAVIS_BUILD_DIR + "/" + "office-js/";
    fs.removeSync(repoLocalFolderPath);

    execCommand(`git clone ${TOKENIZED_GITHUB_PUSH_URL} ${repoLocalFolderPath}`, {
        token: process.env.GH_TOKEN
    });

    shell.pushd(repoLocalFolderPath);


    execCommand(`git checkout ${process.env.TRAVIS_BRANCH}`);
    execCommand('git config --add user.name "Travis CI"');
    execCommand('git config --add user.email "travis.ci@microsoft.com"');

    if (params.afterCloneBeforeCommit) {
        await params.afterCloneBeforeCommit();
    }

    fs.writeFileSync(DEPLOYMENT_YAML_FILENAME, deploymentFileContents);
    execCommand(`git add ${DEPLOYMENT_YAML_FILENAME}`);

    VersionUtils.updatePackageJson(version);

    execCommand(`git commit --allow-empty -m "${TRAVIS_AUTO_COMMIT_TEXT} ${process.env.TRAVIS_COMMIT_MESSAGE}"`);
    execCommand(`git push`);


    // Now that the repo is updated, publish to NPM:

    fs.writeFileSync(".npmrc", `//registry.npmjs.org/:_authToken=${process.env.NPM_TOKEN}`);
    execCommand(`npm publish --tag ${npmPublishTag}`);


    // If NPM succeeded, tag it and also add an NPM release:
    console.log(`Also tag the branch, and make a GitHub release: https://github.com/OfficeDev/office-js/releases/tag/${gitTagName}`);

    execCommand(`git tag -a ${gitTagName} -m "${TRAVIS_AUTO_COMMIT_TEXT} ${process.env.TRAVIS_COMMIT_MESSAGE}"`);
    execCommand(`git push origin ${gitTagName}`);

    const releaseNotesWithNbsp = deploymentFileContents.split("\n").map(line => {
        let regex = /^(\s*)(.+?)(\|\-)?$/;
        // Match 0 or more starting spaces, followed by a non-greedy (lazy) anything, followed by optionally a "|-" at the end.
        let result = regex.exec(line);
        if (!result) {
            return line;
        }
        return (result[1].length > 0 ? ("&nbsp;".repeat(result[1].length - 1) + " ") : "") + result[2];
    }).join("\n");

    // Documentation: https://developer.github.com/v3/repos/releases/#create-a-release
    const response = await fetch("https://api.github.com/repos/OfficeDev/office-js/releases", {
        method: "POST",
        headers: new Headers({
            "Authorization": `token ${process.env.GH_TOKEN}`
        }),
        body: JSON.stringify({
            "tag_name": gitTagName,
            "name": gitTagName,
            "body": releaseNotesWithNbsp,
            "prerelease": true,
            "draft": false
        })
    });

    if (response.status !== 201) {
        throw new Error(`Failed to create GitHub release; ${response.status}: ${response.statusText}`);
    }


    shell.popd();

    let removeLocalFolderAtCompletion = true;
    if (removeLocalFolderAtCompletion) {
        fs.removeSync(repoLocalFolderPath);
    }

    return releaseNotesWithNbsp.replace(/&nbsp;/g, " ");
}

async function getPrivateBranchDeploymentParams(): Promise<IDeploymentParams> {
    const version = await VersionUtils.getNextPrivateVersionNumber();
    const npmPublishTag = "private";

    return {
        version,
        npmPublishTag
    };
}

async function getOfficialBranchDeploymentParams(): Promise<IDeploymentParams> {
    const message = stripSpaces(`
        Sorry, the auto-deployment of official branches isn't supported yet.
    `);
    banner('NOT YET IMPLEMENTED, SKIPPING DEPLOYMENT', message, chalk.yellow.bold);
    process.exit(0);

    return {
        version: "unknown",
        npmPublishTag: "unknown"
    };
}
