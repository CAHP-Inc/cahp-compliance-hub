export * from './limits';
export * from './rentRoll';
export * from './manualRoll';
export * from './entity';
export * from './analyze';
export * from './exhibit';
export * from './letter';

/** Filesystem-safe slug for output filenames. */
export function slugify(name: string): string {
  return name.replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}
