import { findLatestNpmPackageVersion } from "./findLatestNpmPackageVersion";
import * as semver from "semver";

/**
 * get the next package version with the following versioning scheme:
 *
 * release: x.y.z
 *
 * tag: x.y.(z+1)-tag.q
 *
 * The tags x.y.z are always the same as the release version.
 * note: the tags patch version (z) is always one ahead of the release version.
 *
 * @param packageName name of the npm package to update
 * @param tag tag to apply to the package
 */
export function getNextNpmPackageVersion(packageName: string, tag?: string): string {
  const currentVersionRelease = findLatestNpmPackageVersion(packageName);
  // console.log(`currentVersionRelease: ${currentVersionRelease}`);

  // Handle case: release
  if (tag === undefined) {
    return getNextVersionRelease(currentVersionRelease);
  }

  // Handle case: tag
  const currentVersionTag = findLatestNpmPackageVersion(packageName, tag);
  return getNextVersionTag(tag, currentVersionRelease, currentVersionTag);
}

/**
 * Increments the patch version or if the version is undefined returns 0.0.0
 *
 * @param currentVersion x.y.z or undefined
 * @returns x.y.z -> x.y.(z+1) or undefined -> 0.0.0
 */
export function getNextVersionRelease(currentVersion?: string) {
  if (currentVersion === undefined) {
    // first base version ever
    return "0.0.0";
  }

  return semver.inc(currentVersion, "patch") as string;
}

/**
 * Tags are versioned with the following versioning scheme:
 *
 * if x.y.z is equal to or ahead of tx.ty.tz then x.y.z becomes the new base version of the tag:
 *
 * x.y.(z+1)-tag.0
 *
 * otherwise the prerelease version is incremented:
 *
 * tx.ty.tz-tag-(q+1)
 *
 * @param tag tag string
 * @param currentVersionRelease x.y.z or undefined
 * @param currentVersionTag tx.ty.tz-tag.q or undefined
 * @returns the next tag version
 */
export function getNextVersionTag(
  tag: string,
  currentVersionRelease?: string,
  currentVersionTag?: string,
) {
  const basis = currentVersionRelease || "0.0.0";

  if (currentVersionTag === undefined) {
    // first version for the tag
    return new semver.SemVer(basis).inc("prerelease", tag).format();
  }

  const tagVersionCoerce = semver.coerce(currentVersionTag);
  const tagVersionWithoutTag = tagVersionCoerce === null ? "0.0.0" : tagVersionCoerce;
  if (semver.gte(basis, tagVersionWithoutTag)) {
    // release version is ahead ot the tag
    // update the tag base to match
    return new semver.SemVer(basis).inc("prerelease", tag).format();
  }

  return new semver.SemVer(currentVersionTag).inc("prerelease").format();
}

export function test() {
  const failures: string[] = [];
  if (
    !(
      getNextVersionRelease(undefined) === "0.0.0" ||
      getNextVersionRelease("1.1.1") === "1.1.2"
    )
  ) {
    failures.push("Failed basic getNextVersionRelease");
  }

  type TestParameters = {
    tag: string;
    releaseVersion?: string;
    tagVersion?: string;
    expectedVersion: string;
  };
  const tests: TestParameters[] = [
    {
      tag: "tag",
      releaseVersion: undefined,
      tagVersion: undefined,
      expectedVersion: "0.0.1-tag.0",
    }, // first time
    {
      tag: "tag",
      releaseVersion: "0.0.0",
      tagVersion: "0.0.1-tag.0",
      expectedVersion: "0.0.1-tag.1",
    }, // increment tag version
    {
      tag: "tag",
      releaseVersion: "0.0.0",
      tagVersion: "1.0.1-tag.0",
      expectedVersion: "1.0.1-tag.1",
    }, // Follow tag version
    {
      tag: "tag",
      releaseVersion: "1.0.0",
      tagVersion: "0.0.0-tag.0",
      expectedVersion: "1.0.1-tag.0",
    }, // Follow release version
    {
      tag: "tag",
      releaseVersion: "1.0.0",
      tagVersion: "1.0.0-tag.1",
      expectedVersion: "1.0.1-tag.0",
    }, // Follow release version
  ];

  tests.forEach((t: TestParameters) => {
    const actualVersion = getNextVersionTag(t.tag, t.releaseVersion, t.tagVersion);
    if (actualVersion !== t.expectedVersion) {
      failures.push(
        `failed: ${JSON.stringify(t)} expected: [${
          t.expectedVersion
        }] actual: [${actualVersion}]`,
      );
    }
  });

  if (failures.length > 0) {
    throw failures.join("\n");
  }
}
