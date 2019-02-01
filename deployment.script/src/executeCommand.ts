import { isUndefined } from "util";
import * as child_process from 'child_process';

export function executeCommand(command: string, workingDirectory?: string): string {
    console.log(command);

    const options:{cwd?: string, encoding?: string} = {};
    if (!isUndefined(workingDirectory)) {
        options.cwd = workingDirectory;
    }

    options.encoding = "utf-8";

    return child_process.execSync(command, options).toString();
}
