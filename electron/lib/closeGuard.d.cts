export type CloseDecision = 'allow' | 'prompt';
export type ClosePromptResult = 'save' | 'discard' | 'cancel';

export const CLOSE_DECISION: {
  readonly ALLOW: 'allow';
  readonly PROMPT: 'prompt';
};

export const CLOSE_PROMPT_BUTTONS: readonly ['Save', 'Discard', 'Cancel'];

export const CLOSE_PROMPT_CHOICE: {
  readonly SAVE: 0;
  readonly DISCARD: 1;
  readonly CANCEL: 2;
};

export function decideOnClose(isDirty: boolean): CloseDecision;

export function decideAfterPromptChoice(choice: number | undefined): ClosePromptResult;
