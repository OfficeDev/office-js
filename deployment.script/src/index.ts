import {EnvironmentVariables} from "./EnvironmentVariables";

// import {process} from "./process"

const fieldsToPrint: (keyof EnvironmentVariables)[] = [
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

