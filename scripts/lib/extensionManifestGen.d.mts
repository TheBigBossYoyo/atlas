export interface AssociationLike {
  readonly ext: string
  readonly associationName: string
  readonly associationDescription: string
  readonly associationRole: string
}

export function buildExtensionList(manifest: ReadonlyArray<{ readonly ext: string }>): string[]
export function buildGeneratedCjsSource(manifest: ReadonlyArray<{ readonly ext: string }>): string
export function buildYamlAssociationsBlock(manifest: ReadonlyArray<AssociationLike>): string
export function spliceYamlAssociations(ymlText: string, block: string): string
export const YAML_BEGIN_MARKER: string
export const YAML_END_MARKER: string
