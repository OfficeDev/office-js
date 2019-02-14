import { ReleaseType } from "./ReleaseType";

/**
 * The tag that should be applied to the package
 * @param release_type
 */
export function getNpmPackageTag(release_type: ReleaseType): string | undefined {
  if (release_type === ReleaseType.release) {
    return undefined;
  } else {
    return release_type as string;
  }
}
