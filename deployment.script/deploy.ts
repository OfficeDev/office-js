#!/usr/bin/env node --harmony

import { isString, isNil, cloneDeep } from "lodash";
import * as chalk from 'chalk';
import * as shell from 'shelljs';
import * as fs from "fs-extra";
import * as path from 'path';
import * as jsyaml from 'js-yaml';

import { banner, stripSpaces, execCommand, fetchAndThrowOnError, AdditionalInfoError } from "./util";
import * as VersionUtils from "./version-number-utils";

declare var process: {
    env: IEnvironmentVariables
    exit: (status: number) => void;
};

const TRAVIS_AUTO_COMMIT_TEXT = "[TRAVIS CI AUTO-COMMIT]";
const TOKENIZED_GITHUB_PUSH_URL = `https://<<<token>>>@github.com/OfficeDev/office-js.git`;
const DEPLOYMENT_YAML_FILENAME = "NPM.DEPLOYMENT.INFO.yaml";
const DEPLOY_REQUEST_FILENAME = "DEPLOY_REQUEST.yaml";

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
const OFFICIAL_TAGS = cloneDeep(OFFICIAL_BRANCHES).concat("latest");
const DEPLOYMENT_QUEUE_BRANCH = "deployment-queue";
const ADHOC_BRANCH_PREFIX = "__adhoc";
const DEFAULT_ADHOC_TAG = "adhoc";

interface IDeploymentInfoFromSubmittedRepo {
    tag: string,
    history: IDeploymentInfoFromSubmittedRepoHistory,
}
interface IDeploymentInfoFromSubmittedRepoHistory {
    commitMessage: string,
    adhocBranchName: string,
    fullCommitHistory: string,
}

const WORKING_DIRECTORY = path.resolve(process.env.TRAVIS_BUILD_DIR, "..", "working-travis-output-dir");

interface IOfficialBranchDeployRequest {
    targetBranch: string;
    from: string;
    deleteAdhocBranchOnSuccessfulDeployment: boolean;
}

interface IDeploymentParams {
    isOfficialBuild: boolean;
    branchToCheckOut: string;
    npmPublishTag: string;
    commitMessagePartial: string;

    historyInfo: IDeploymentInfoFromSubmittedRepoHistory;

    /** A script to run after cloning. Note that the current working directory at that point
     * is still the original one from the start of the script */
    afterCloneBeforeCommit?: (repoLocalFolderPath: string) => Promise<any>;

    rightBeforeCompletion?: () => Promise<any>;
}

(async () => {
    try {
        await attemptDeployScript();
        process.exit(0);
    }
    catch (error) {
        banner('AN ERROR OCCURRED', error.message || error, chalk.bold.red);
        console.error(error);

        banner('DEPLOYMENT DID NOT GET TRIGGERED', null, chalk.bold.red);
        process.exit(1);
    }
})();

async function attemptDeployScript() {
    printBuildStartInfo();
    makeWorkingDirectory();

    precheckOrExit();

    if (process.env.TRAVIS_BRANCH.startsWith(ADHOC_BRANCH_PREFIX)) {
        const deploymentInfoFromSubmittedRepoState = await getDeploymentInfoFromGithubRepoState(process.env.TRAVIS_BRANCH);
        const npmPublishTag = deploymentInfoFromSubmittedRepoState.tag;
        await doDeployment({
            isOfficialBuild: false,
            commitMessagePartial: process.env.TRAVIS_COMMIT_MESSAGE,
            branchToCheckOut: process.env.TRAVIS_BRANCH,
            npmPublishTag,
            historyInfo: deploymentInfoFromSubmittedRepoState.history
        });
        return;

    } else if (process.env.TRAVIS_BRANCH === DEPLOYMENT_QUEUE_BRANCH) {
        await doOfficialDeployment();
        return;

    } else if (OFFICIAL_BRANCHES.indexOf(process.env.TRAVIS_BRANCH) >= 0) {
        const message = stripSpaces(`
            Deployment to one of the official branches must happen through the
            "${DEPLOYMENT_QUEUE_BRANCH}" branch. Please see
            https://github.com/OfficeDev/office-js/blob/deployment-queue/README.md
            for more info.
        `);
        banner('SKIPPING DEPLOYMENT', message, chalk.yellow.bold);
        return;

    } else {
        const message = stripSpaces(`
            UNKNOWN BRANCH: Branch "${process.env.TRAVIS_BRANCH}" does not match any of the following:
                * A branch that starts with "__${ADHOC_BRANCH_PREFIX}".
                * The "${DEPLOYMENT_QUEUE_BRANCH}" branch.
                * Any of the following official branches: [${OFFICIAL_BRANCHES.map(item => `"${item}"`).join(", ")}].
        `);
        banner('SKIPPING DEPLOYMENT', message, chalk.yellow.bold);
        return;
    }
}

function makeWorkingDirectory() {
    if (!fs.existsSync(WORKING_DIRECTORY)) {
        fs.mkdirSync(WORKING_DIRECTORY);
    } else {
        fs.emptyDirSync(WORKING_DIRECTORY);
    }

    banner("Working directory", WORKING_DIRECTORY);

    shell.pushd(WORKING_DIRECTORY);
    execCommand('dir');
    shell.popd();
}

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
        // NOTE: in practice, such builds should also be skipped because they'll have the text "skip ci" in them.
        // But this serves as a double-guarantee.
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

async function doDeployment(params: IDeploymentParams): Promise<void> {
    const { npmPublishTag, historyInfo, isOfficialBuild, commitMessagePartial } = params;

    if (!isOfficialBuild && OFFICIAL_TAGS.indexOf(npmPublishTag) >= 0) {
        throw new Error("Private build may not use an official NPM tag!");
    }

    const repoLocalFolderPath = WORKING_DIRECTORY + "/" + "office-js/";
    fs.removeSync(repoLocalFolderPath);

    execCommand(`git clone --single-branch --depth 1 -b ${params.branchToCheckOut} ${TOKENIZED_GITHUB_PUSH_URL} ${repoLocalFolderPath}`, {
        token: process.env.GH_TOKEN
    });

    if (params.afterCloneBeforeCommit) {
        await params.afterCloneBeforeCommit(repoLocalFolderPath);
    }

    shell.pushd(repoLocalFolderPath);

    execCommand('git config --add user.name "Travis CI"');
    execCommand('git config --add user.email "travis.ci@microsoft.com"');


    let npmDeploymentSuceeded = false;
    let version: string = await VersionUtils.getNextVersionNumber(npmPublishTag);
    let deploymentFileContents: string;
    let markdownReleasesNotes: string;

    while (!npmDeploymentSuceeded) {
        console.log(chalk.cyan.bold(`Tenantive version: ${version}`));

        const commonHandlebarsParams = {
            npmPublishTag,
            version,
            travisBuildId: process.env.TRAVIS_BUILD_ID,
            includeTagUrls: npmPublishTag !== DEFAULT_ADHOC_TAG,
        };
        deploymentFileContents = VersionUtils.generateDeploymentYamlText({
            ...commonHandlebarsParams,
            historyInfo: historyInfo,
            travisBuildNumber: process.env.TRAVIS_BUILD_NUMBER,
        });
        markdownReleasesNotes = VersionUtils.generateMarkdownDescription({
            ...commonHandlebarsParams,
            DEPLOYMENT_YAML_FILENAME,
            commitMessage: historyInfo.commitMessage,
        });

        fs.writeFileSync(DEPLOYMENT_YAML_FILENAME, deploymentFileContents!);
        VersionUtils.updatePackageJson(version);

        // Publish to NPM first (even before writing to repo).
        // Rationale:  that way, if there is an unpublished version, and run into an
        //     "npm ERR! You cannot publish over the previously published version XYZ"
        // error, we can revamp the version number and keep trying until we get it right.

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
    }


    console.log(`Having successfully published to NPM, finish off any remaining NPM tasks:`);

    // For "office-js" package, the "release" tag is same as "latest" -- so for release, tag it as the "latest" too:
    if (npmPublishTag === "release") {
        execCommand(`npm dist-tag add @microsoft/office-js@${version!} latest`);
    }

    console.log(chalk.magenta(`FYI, if you need to unpublish (must be done within the first 24 hours), run:`));
    console.log(chalk.magenta(`    npm unpublish @microsoft/office-js@${version!}`));



    console.log(`Now commit and push to the repo`);

    const commitMessage = `${TRAVIS_AUTO_COMMIT_TEXT} ${commitMessagePartial} [skip ci]`;
    // Note: "skip CI" will skip travis running on the build.  https://docs.travis-ci.com/user/customizing-the-build/#Skipping-a-build

    console.log("");
    execCommand(`git add -A`);
    execCommand(`git commit --allow-empty -m "${commitMessage}"`);
    execCommand(`git push`);


    const gitTagName = "v" + version;
    console.log(`Also tag the branch, and make a GitHub release: https://github.com/OfficeDev/office-js/releases/tag/${gitTagName}`);
    execCommand(`git tag -a ${gitTagName} -m "${commitMessage}"`);
    execCommand(`git push origin ${gitTagName}`);

    console.log(chalk.magenta(`FYI, if you need to delete the tag, run`));
    console.log(chalk.magenta(`    git push --delete origin ${gitTagName}`));
    console.log(chalk.magenta(`You'll also need to discard the resulting draft in "https://github.com/OfficeDev/office-js/releases"`));
    console.log(chalk.magenta(`    (be sure that you're logged in to see it)`));

    // Documentation: https://developer.github.com/v3/repos/releases/#create-a-release
    const response = await fetch("https://api.github.com/repos/OfficeDev/office-js/releases", {
        method: "POST",
        headers: new Headers({
            "Authorization": `token ${process.env.GH_TOKEN}`
        }),
        body: JSON.stringify({
            "tag_name": gitTagName,
            "name": gitTagName,
            "body": markdownReleasesNotes!,
            "prerelease": !isOfficialBuild,
            "draft": false
        })
    });

    if (response.status !== 201) {
        throw new Error(`Failed to create GitHub release; ${response.status}: ${response.statusText}`);
    }


    shell.popd();


    if (params.rightBeforeCompletion) {
        await params.rightBeforeCompletion();
    }


    banner('SUCCESS, DEPLOYMENT COMPLETE!', markdownReleasesNotes!.replace(/&nbsp;/g, ' '), chalk.green.bold);

    banner(`GitHub Releases page for v${version}`, `https://github.com/OfficeDev/office-js/releases/tag/v${version}`, chalk.green.bold);
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

async function getDeploymentInfoFromGithubRepoState(lookupBranch: string): Promise<IDeploymentInfoFromSubmittedRepo> {
    const url = `https://raw.githubusercontent.com/OfficeDev/office-js/${lookupBranch}/NPM.DEPLOYMENT.INFO.yaml`;
    const contents = await fetchAndThrowOnError(url, "text");

    const result = jsyaml.safeLoad(contents) as IDeploymentInfoFromSubmittedRepo;
    if (result.history && result.tag) {
        return result;
    } else {
        throw new Error(`Missing required fields from in-repo "${DEPLOYMENT_YAML_FILENAME}" file of branch "${lookupBranch}".` +
            "\n\n" + url + "\n\n" + contents);
    }
}

async function doOfficialDeployment(): Promise<void> {
    console.log(`First off: is there a request for a "targetBranch" and "from" in the ${DEPLOY_REQUEST_FILENAME} file?`);
    let currentYaml: IOfficialBranchDeployRequest = jsyaml.safeLoad(
        fs.readFileSync(process.env.TRAVIS_BUILD_DIR + "/" + DEPLOY_REQUEST_FILENAME).toString());

    if (isNil(currentYaml.targetBranch) || isNil(currentYaml.from)) {
        banner('SKIPPING DEPLOYMENT', `Nothing to deploy: missing "targetBranch" and/or "from" parameters.`, chalk.yellow.bold);
        return;
    }

    if (OFFICIAL_BRANCHES.indexOf(currentYaml.targetBranch) < 0) {
        throw new Error(`Invalid target branch "${currentYaml.targetBranch}"; does not belong to the list of official branches`);
    }

    banner("DEPLOYMENT REQUEST DETECTED", stripSpaces(`
        Acknowledging request to deploy to
            "${currentYaml.targetBranch}"
        from
            ${currentYaml.from}
    `));

    // for purposes of the official tags, the NPM publish tag will be the same as the branch name (beta, release-next, etc)
    const npmPublishTag = currentYaml.targetBranch;
    const historyInfo = (await getDeploymentInfoFromGithubRepoState(currentYaml.from)).history;

    await doDeployment({
        isOfficialBuild: true,
        commitMessagePartial: `Deploy to '${currentYaml.targetBranch}' from '${currentYaml.from}'`,
        branchToCheckOut: currentYaml.targetBranch,
        npmPublishTag,
        historyInfo,
        afterCloneBeforeCommit: async (repoLocalFolderPath: string) => {
            console.log(`Delete all files except ".git" and the "dist" folder`);
            fs.readdirSync(repoLocalFolderPath)
                .filter(filename => [".git", "dist"].indexOf(filename) < 0)
                .forEach(filename => fs.removeSync(repoLocalFolderPath + '/' + filename));


            (() => {
                console.log(`Now copy over all files from the release branch except ".git" and "dist":`);
                const repoReleaseCopyFolderPath = `${WORKING_DIRECTORY}/office-js-repo-release-copy-${new Date().getTime()}/`;
                fs.removeSync(repoReleaseCopyFolderPath);
                execCommand(`git clone --single-branch --depth 1 ${TOKENIZED_GITHUB_PUSH_URL} ${repoReleaseCopyFolderPath}`, {
                    token: process.env.GH_TOKEN
                });
                fs.readdirSync(repoReleaseCopyFolderPath)
                    .filter(filename => [".git", "dist"].indexOf(filename) < 0)
                    .forEach(filename => fs.copySync(
                        repoReleaseCopyFolderPath + '/' + filename,
                        repoLocalFolderPath + '/' + filename,
                        {
                            preserveTimestamps: true
                        }
                    ));
            })();

            (() => {
                // And do this again, this time with copying ONLY the DIST folder from the "from" branch:
                const repoFromBranchCopyFolderPath = `${WORKING_DIRECTORY}/office-js-repo-release-copy-${new Date().getTime()}/`;
                execCommand(`git clone --single-branch --depth 1 -b ${currentYaml.from} ${TOKENIZED_GITHUB_PUSH_URL} ${repoFromBranchCopyFolderPath}`, {
                    token: process.env.GH_TOKEN
                });
                fs.copySync(
                    repoFromBranchCopyFolderPath + '/dist',
                    repoLocalFolderPath + '/dist',
                    {
                        preserveTimestamps: true
                    }
                );
            })();
        },
        rightBeforeCompletion: async () => {
            (() => {
                console.log(`Reset the ${DEPLOY_REQUEST_FILENAME} "from" and "targetBranch" keys:`);
                const repoDeployCopyFolderPath = `${WORKING_DIRECTORY}/office-js-repo-deploy-copy-${new Date().getTime()}/`;
                fs.removeSync(repoDeployCopyFolderPath);
                execCommand(`git clone --single-branch --depth 1 -b ${DEPLOYMENT_QUEUE_BRANCH} ${TOKENIZED_GITHUB_PUSH_URL} ${repoDeployCopyFolderPath}`, {
                    token: process.env.GH_TOKEN
                });
                shell.pushd(repoDeployCopyFolderPath);

                const repoDeployCopyDeployRequestFilePath = repoDeployCopyFolderPath + "/" + DEPLOY_REQUEST_FILENAME;
                const sanitizedDeployRequestFileContents =
                    (fs.readFileSync(repoDeployCopyDeployRequestFilePath).toString().split("\n"))
                        .map(line => {
                            if (line.startsWith("from:")) {
                                return "from:";
                            } else if (line.startsWith("targetBranch:")) {
                                return "targetBranch:";
                            }
                            return line;
                        })
                        .join("\n");
                fs.writeFileSync(repoDeployCopyDeployRequestFilePath, sanitizedDeployRequestFileContents);
                execCommand(`git add -A`);
                execCommand(`git commit --allow-empty -m "Remove ${DEPLOY_REQUEST_FILENAME} parameters, to ready for next job [skip ci]"`);
                execCommand(`git push origin ${DEPLOYMENT_QUEUE_BRANCH}`);

                if (currentYaml.deleteAdhocBranchOnSuccessfulDeployment) {
                    execCommand(`git push origin --delete ${currentYaml.from}`);
                }
            })();
        }
    });
}
