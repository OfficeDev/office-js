import { EnvironmentVariables } from "./EnvironmentVariables";
import { getReleaseTypeFromBranchName, ReleaseType } from "./ReleaseType";

export function deploymentPrerequisitesPassed(env: EnvironmentVariables): boolean {
    let shouldDeploy = true;

    if (!env.TRAVIS) {
        console.log(`Deployment skipped - Not running inside of Travis.`);
        shouldDeploy = false;
    }
    
    // Do not run for pull requests.
    // Careful! Need this check to ensure a pull request does NOT trigger a deployment.
    if (env.TRAVIS_PULL_REQUEST) {
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