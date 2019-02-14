import * as semver from "semver";
import { isUndefined } from "util";
import { executeCommand } from "./executeCommand";

/**
 * Find the latest npm package version with a specific tag
 *
 * NPM packages are labeled as follows: x.y.z or x.y.z-tag.q
 *
 * Find the highest version package.
 *
 * @param tag optional tag the package is tagged with.
 */
export function findLatestNpmPackageVersion(
  packageName: string,
  tag?: string,
): string | undefined {
  const data = executeCommand(`npm view ${packageName} versions --json`);
  const versions: string[] = JSON.parse(data);

  let filter: (version: string) => boolean = () => false;

  if (isUndefined(tag)) {
    filter = (version: string) => !version.includes("-");
  } else {
    filter = (version: string) => version.includes(`-${tag}.`);
  }

  const all_matching = versions.filter(filter).sort(semver.compare);

  // Note: it's possible that there are no matches
  return all_matching.pop();
}
