import * as chalk from 'chalk';
import * as shell from "shelljs";
require('isomorphic-fetch');


/**
 * Creates a chalk based section with the specific color.
 * @param title Title of the banner.
 * @param message Message of the banner.
 * @param chalkFunction Chalk color function.
 */
export const banner = (title: string, message: string | null = null, chalkFn: chalk.ChalkChain | null = null) => {
    if (!chalkFn) {
        chalkFn = chalk.bold;
    }

    const dashes = Array(Math.max(title.length + 1, 30)).join('-');
    console.log("\n\n");
    console.log(chalkFn(`${dashes}`));
    console.log(chalkFn(`${title}`));
    if (message) {
        console.log(chalkFn(dashes));
        console.log(message);
    }
    console.log(chalkFn(`${dashes}`));
    console.log("\n");
};


export async function fetchAndThrowOnError(url: string, format: 'text'): Promise<string>;
export async function fetchAndThrowOnError<T>(url: string, format: 'json'): Promise<T>;
export async function fetchAndThrowOnError(url: string, format: 'text' | 'json') {
    let response = await fetch(url);
    if (response.status >= 400) {
        throw new Error(`Bad response from server for URL ${url}`);
    }

    switch (format) {
        case 'text':
            return await response.text();
        case 'json':
            return await response.json();
        default:
            throw new Error("Invalid format specified");
    }
}

export function shellExecOrFail(command: string) {
    console.log(">> " + command);
    let result = shell.exec(command);
    if (result.code !== 0) {
        console.error(`An error occurred while executing \`${command}\``);
        console.error(result.stderr);
        process.exit(1);
    }
}

export async function runNpmCommand<T>(command: string, ...args: any[]): Promise<T> {
    console.log(chalk.white.bold(`npm ${command} ${args.join(" ")}`));

    chalk.reset();

    return new Promise<T>((resolve, reject) => {
        const npm = require('npm');
        npm.load((error: any) => {
            if (error) {
                reject(error);
            }

            npm.commands[command](args, (error: any, data: T) => {
                if (error) {
                    reject(error);
                }

                resolve(data);
            });
        });
    });
}

export function stripSpaces(text: string) {
    let lines: string[] = text.split('\n');

    // Replace each tab with 4 spaces.
    for (let i: number = 0; i < lines.length; i++) {
        lines[i].replace('\t', '    ');
    }

    let isZeroLengthLine: boolean = true;
    let arrayPosition: number = 0;

    // Remove zero length lines from the beginning of the snippet.
    do {
        let currentLine: string = lines[arrayPosition];
        if (currentLine.trim() === '') {
            lines.splice(arrayPosition, 1);
        } else {
            isZeroLengthLine = false;
        }
    } while (isZeroLengthLine || (arrayPosition === lines.length));

    arrayPosition = lines.length - 1;
    isZeroLengthLine = true;

    // Remove zero length lines from the end of the snippet.
    do {
        let currentLine: string = lines[arrayPosition];
        if (currentLine.trim() === '') {
            lines.splice(arrayPosition, 1);
            arrayPosition--;
        } else {
            isZeroLengthLine = false;
        }
    } while (isZeroLengthLine);

    // Get smallest indent for align left.
    let shortestIndentSize: number = 1024;
    for (let line of lines) {
        let currentLine: string = line;
        if (currentLine.trim() !== '') {
            let spaces: number = line.search(/\S/);
            if (spaces < shortestIndentSize) {
                shortestIndentSize = spaces;
            }
        }
    }

    // Align left
    for (let i: number = 0; i < lines.length; i++) {
        if (lines[i].length >= shortestIndentSize) {
            lines[i] = lines[i].substring(shortestIndentSize);
        }
    }

    // Convert the array back into a string and return it.
    let finalSetOfLines: string = '';
    for (let i: number = 0; i < lines.length; i++) {
        if (i < lines.length - 1) {
            finalSetOfLines += lines[i] + '\n';
        }
        else {
            finalSetOfLines += lines[i];
        }
    }
    return finalSetOfLines;
}
