#!/usr/bin/env node --harmony

import { isString } from "lodash";
import * as chalk from 'chalk';

import { banner } from "./util";

declare var process: {
    env: IEnvironmentVariables
    exit: (status: number) => void;
};

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

(async () => {
    try {
        printBuildStartInfo();

        if (!precheck()) {
            return;
        }

        await deploy();

        process.exit(0);
    } catch (e) {
        console.error(e);
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

function precheck(): boolean {
    /* Check if the code is running inside of travis.ci. If not abort immediately. */
    if (!process.env.TRAVIS) {
        banner('Deployment skipped', 'Not running inside of Travis.', chalk.yellow.bold);
        return false;
    }

    // Careful! Need this check because otherwise, a pull request against master would immediately trigger a deployment.
    if (process.env.TRAVIS_PULL_REQUEST !== 'false') {
        banner('Deployment skipped', 'Skipping deploy for pull requests.', chalk.yellow.bold);
        return false;
    }

    REQUIRED_ADDITIONAL_FIELDS.forEach(key => {
        if (!isString(process.env[key]) || (process.env[key] as string).trim().length <= 0) {
            throw new Error(`"${key}" is a required global variables.`);
        }
    });

    return true;
}

function deploy(): void {

}
