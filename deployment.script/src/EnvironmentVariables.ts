import { isUndefined } from "util";

export interface EnvironmentVariables {
  TRAVIS: boolean;
  TRAVIS_BRANCH: string;
  TRAVIS_PULL_REQUEST: boolean;
  TRAVIS_COMMIT: string;
  TRAVIS_COMMIT_MESSAGE: string;
  TRAVIS_BUILD_ID: string;
  TRAVIS_BUILD_NUMBER: string;
  TRAVIS_BUILD_DIR: string;

  /**
   * GitHub token generated using https://github.com/settings/tokens,
   *     bearing permissions for "repo:status", "repo_deployment", and "public_repo".
   * This is a personal access token, so the commits always happen on behalf
   *     of the person who created the token.
   * The token is then entered as a hidden value in https://travis-ci.org/OfficeDev/office-js/settings */
  GH_TOKEN: string;

  /** A token for publishing to NPM.  It can be generated using "npm token create"
   * Note that you'll need NPM version 5.5.1+ to run this command.
   * https://docs.npmjs.com/getting-started/working_with_tokens
   */
  NPM_TOKEN: string;
}

function env(variable: string): string {
  const value = process.env[variable];
  return isUndefined(value) ? "" : value;
}

export function getEnvironmentVariables(): EnvironmentVariables {
  const environment: EnvironmentVariables = {
    TRAVIS: env("TRAVIS") === "true",
    TRAVIS_BRANCH: env("TRAVIS_BRANCH"),
    TRAVIS_PULL_REQUEST: env("TRAVIS_PULL_REQUEST") === "true",
    TRAVIS_COMMIT: env("TRAVIS_COMMIT"),
    TRAVIS_COMMIT_MESSAGE: env("TRAVIS_COMMIT_MESSAGE"),
    TRAVIS_BUILD_ID: env("TRAVIS_BUILD_ID"),
    TRAVIS_BUILD_NUMBER: env("TRAVIS_BUILD_NUMBER"),
    TRAVIS_BUILD_DIR: env("TRAVIS_BUILD_DIR"),
    GH_TOKEN: env("GH_TOKEN"),
    NPM_TOKEN: env("NPM_TOKEN"),
  };

  return environment;
}
