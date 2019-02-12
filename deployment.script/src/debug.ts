import { EnvironmentVariables } from "./EnvironmentVariables";
import * as standardFile from "./standardFile";

export function debug(env: EnvironmentVariables) {
    // Print environment variables
    const fieldsToPrint: (keyof EnvironmentVariables)[] = [
        "TRAVIS",
        "TRAVIS_BRANCH",
        "TRAVIS_PULL_REQUEST",];

    const fields = fieldsToPrint.map(item => `"${item}": "${env[item]}"`).join(",\n");
    console.log(fields);

    console.log(`subdirectories: [${env.TRAVIS_BUILD_DIR}]`);
    console.log(standardFile.getSubDirectories(env.TRAVIS_BUILD_DIR));
}