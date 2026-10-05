# Import necessary libraries
import * as os from 'os';

# Define a function to debug the Office.js library
export function debugOfficeJs(): string {
    try {
        // Get the Office.js library version
        const officeJsVersion = getOfficeJsVersion();

        // Check if the Office.js library version is valid
        if (officeJsVersion === 'Error: Office.js library version not found') {
            return 'Error: Office.js library version not found';
        } else {
            // Return a success message
            return 'Office.js library version: ' + officeJsVersion;
        }
    } catch (error) {
        // Return an error message if an error occurs
        return 'Error: ' + error.message;
    }
}