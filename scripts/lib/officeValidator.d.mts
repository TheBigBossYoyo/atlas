export interface OfficeValidationIssue {
  readonly severity: 'error' | 'warning'
  readonly code: string
  readonly part: string | null
  readonly message: string
}

export interface OfficeFormat {
  readonly family: 'opc' | 'odf' | 'unknown'
  readonly kind: 'docx' | 'pptx' | 'xlsx' | 'xlsb' | 'odt' | 'odp' | 'ods' | 'opc-unknown' | 'odf-unknown' | 'unknown'
  readonly mimetype?: string
}

export interface OfficeValidationResult {
  readonly format: OfficeFormat
  readonly issues: OfficeValidationIssue[]
}

export function validateOfficeFile(buffer: Buffer): OfficeValidationResult
export function formatIssuesReport(filePath: string, result: OfficeValidationResult): string
