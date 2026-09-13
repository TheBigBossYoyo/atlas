export type SlideData = {
  readonly id: string
  readonly index: number
  readonly texts: ReadonlyArray<{
    readonly text: string
    readonly x?: number
    readonly y?: number
    readonly w?: number
    readonly h?: number
  }>
  readonly images: ReadonlyArray<{
    readonly src: string
    readonly x: number
    readonly y: number
    readonly w: number
    readonly h: number
  }>
  readonly width: number
  readonly height: number
}
