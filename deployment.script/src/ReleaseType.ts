
export enum ReleaseType {
    release = "release",
    beta = "beta",
    custom = "custom",
    none = "none"
}

// Base deployment action on the branch name
const DEPLOYMENT_BRANCH_NAME_RELEASE = ReleaseType.release as string;
const DEPLOYMENT_BRANCH_NAME_BETA =  ReleaseType.beta as string;
const DEPLOYMENT_BRANCH_NAME_CUSTOM_PREFIX =  ReleaseType.custom as string;

export function getReleaseTypeFromBranchName(branchName: string): ReleaseType {

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
