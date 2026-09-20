import { afterEach, describe, expect, it, vi } from 'vitest';
import { prefersReducedMotion, scrollIntoViewRespectingMotionPreference } from '../motionPreference';

function mockMatchMedia(matches: boolean) {
  const matchMedia = vi.fn().mockReturnValue({ matches } as MediaQueryList);
  vi.stubGlobal('matchMedia', matchMedia);
  return matchMedia;
}

describe('prefersReducedMotion', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('reflects the (prefers-reduced-motion: reduce) media query', () => {
    mockMatchMedia(true);
    expect(prefersReducedMotion()).toBe(true);
  });

  it('is false when the media query does not match', () => {
    mockMatchMedia(false);
    expect(prefersReducedMotion()).toBe(false);
  });
});

describe('scrollIntoViewRespectingMotionPreference', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('downgrades a smooth scroll to instant when reduced motion is on', () => {
    mockMatchMedia(true);
    const element = document.createElement('div');
    element.scrollIntoView = vi.fn();

    scrollIntoViewRespectingMotionPreference(element, { behavior: 'smooth', block: 'center' });

    expect(element.scrollIntoView).toHaveBeenCalledWith({ behavior: 'auto', block: 'center' });
  });

  it('leaves the requested behavior untouched when reduced motion is off', () => {
    mockMatchMedia(false);
    const element = document.createElement('div');
    element.scrollIntoView = vi.fn();

    scrollIntoViewRespectingMotionPreference(element, { behavior: 'smooth', block: 'center' });

    expect(element.scrollIntoView).toHaveBeenCalledWith({ behavior: 'smooth', block: 'center' });
  });

  it('passes through with no options at all', () => {
    mockMatchMedia(true);
    const element = document.createElement('div');
    element.scrollIntoView = vi.fn();

    scrollIntoViewRespectingMotionPreference(element);

    expect(element.scrollIntoView).toHaveBeenCalledWith(undefined);
  });
});
