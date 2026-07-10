# Import necessary libraries
import * as fs from 'fs';

# Define a function to get the standard file
export function getStandardFile(): string {
    try {
        // Get the standard file
        const standardFile = fs.readFileSync('standard.txt', 'utf8');

        // Return the standard file
        return standardFile;
    } catch (error) {
        // Return an error message if the standard file is not found
        return 'Error: Standard file not found';
    }
}