import * as fs from "fs";
import * as path from "path";

export function standardNewlines(s: string): string {
    return s.replace(/\r/gm, "");
}


/**
 * Read utf-8 file and transform to standard new lines.
 * @param path file path
 */
export function readFile(path: string): string {
    return standardNewlines(fs.readFileSync(path, "utf-8"));
}

/**
 * write data to path with standard newlined.
 * @param path 
 * @param data 
 */
export function writeFile(path: string, data: string) {
    const cleanData = standardNewlines(data);
    fs.writeFileSync(path, cleanData);
}

/**
 * Read a file that contains json and turn it into an object
 *
 * Note: no validation is done on the data.
 * TODO: add validation of json to ensure it conforms to schema to type.
 *
 * @param path path to the json file
 */
export function readFileJson<T>(path: string): T {
    const data: string = readFile(path);
    const object: T = JSON.parse(data);
    return object;
}

export function IsDirectory(directory: string): boolean {
    return fs.lstatSync(directory).isDirectory();
}

function IsFile(file: string): boolean {
    return fs.lstatSync(file).isFile();
}


/**
 * Transform a data object to a string and write it to the specified path.
 * @param path 
 * @param data 
 */
export function writeFileJson(path: string, data: {}): void {
    const json: string = JSON.stringify(data, undefined, 2);
    writeFile(path, json);
}


export function getSubDirectories(directory: string): string[] {
    const all = fs.readdirSync(directory);
    const directories = all.filter((sub) => IsDirectory(path.join(directory, sub)));
    return directories.sort();
}

export function getFilesInDirectory(directory: string): string[] {
    const all = fs.readdirSync(directory);
    const files = all.filter((file) => IsFile(path.join(directory, file)));
    return files.sort();
}