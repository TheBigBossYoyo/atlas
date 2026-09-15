export function resolveIsDev(deps: {
  atlasDevEnv: string | undefined;
  distIndexExists: () => boolean;
}): boolean;
