import { isUndefined } from "util";
import * as child_process from 'child_process';

export function executeCommand(command: string, workingDirectory?: string, ignoreError?: boolean): string {
    console.log(command);

    const options:{cwd?: string, encoding?: string, stdio?: any[]} = {};
    if (!isUndefined(workingDirectory)) {
        options.cwd = workingDirectory;
    }

    if (!isUndefined(ignoreError) && ignoreError) {
        options.stdio = ["ignore", "pipe", "ignore"];
    }
    
    options.encoding = "utf-8";

    return child_process.execSync(command, options).toString();
}
