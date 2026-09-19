export type FileOpenAction = 'send' | 'queue';

export interface FileOpenRoutingState {
  hasWindow: boolean;
  rendererReady: boolean;
}

export function decideFileOpenAction(state: FileOpenRoutingState): FileOpenAction;
