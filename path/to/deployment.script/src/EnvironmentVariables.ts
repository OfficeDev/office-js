# Import necessary libraries
import * as os from 'os';

# Define a function to get the Office.js library version
export function getOfficeJsVersion(): string {
    try {
        // Get the Office.js library version from the environment variables
        const officeJsVersion = os.environ['OFFICE_JS_VERSION'];

        // Return the Office.js library version
        return officeJsVersion;
    } catch (error) {
        // Return an error message if the Office.js library version is not found
        return 'Error: Office.js library version not found';
    }
}