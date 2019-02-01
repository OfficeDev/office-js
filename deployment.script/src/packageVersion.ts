// /*
// Figure out what the package version should be

// */






// /**
//  * Figure out what to label the npm package and increment appropriately
//  * 
//  * release: x.y.z
//  * 
//  * beta: x.y.z-beta.q
//  * 
//  * custom: x.y.z-custom.p
//  */
// export function getPackageVersion(): string {
    
//     return "";
// }


// async function getReleasePackageVersonStringFromGitHub(): Promise<string> {
//     const url = `https://raw.githubusercontent.com/OfficeDev/office-js/release/package.json`;
//     const versionString = (await fetchJson<{ version: string }>(url, "json")).version;

//     if (isNull(versionString) || !isString(versionString) || versionString.length <= 0) {
//         throw new Error(`Missing or invalid package version number at URL "${url}"`);
//     }

//     if (!semver.valid(versionString)) {
//         throw new Error("Invalid release build number, should never happen");
//     }

//     return versionString;
// }


// async function fetchJson<T>(url: string, format: 'text' | 'json'): Promise<T> {
//     let response = await fetch(url);
//     if (response.status >= 400) {
//         throw new Error(`Bad response from server for URL ${url}`);
//     }

//     switch (format) {
//         case 'json':
//             return await response.json();
//         default:
//             throw new Error("Invalid format specified");
//     }
// }