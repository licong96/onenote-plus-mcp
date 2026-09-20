import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { mkdir, mkdtemp, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { readLocalAttachment } from '@/tools/createPage.js';

describe('readLocalAttachment', () => {
  let dir: string;
  let originalCwd: string;

  beforeEach(async () => {
    originalCwd = process.cwd();
    dir = await mkdtemp(join(tmpdir(), 'onenote-mcp-attach-'));
    process.chdir(dir);
  });

  afterEach(async () => {
    process.chdir(originalCwd);
    await rm(dir, { recursive: true, force: true });
  });

  it('reads a file inside cwd', async () => {
    await writeFile('inside.bin', Buffer.from([1, 2, 3]));
    const bytes = await readLocalAttachment('inside.bin');
    expect(Array.from(bytes)).toEqual([1, 2, 3]);
  });

  it('rejects relative path traversal (..)', async () => {
    await expect(readLocalAttachment('../escape.bin')).rejects.toThrow(
      /outside the working directory/,
    );
  });

  it('rejects absolute paths outside cwd', async () => {
    await expect(readLocalAttachment('/etc/passwd')).rejects.toThrow(
      /outside the working directory/,
    );
  });

  it('allows files whose names start with two dots (not a traversal)', async () => {
    await writeFile('..dotfile', Buffer.from([7, 8]));
    const bytes = await readLocalAttachment('..dotfile');
    expect(Array.from(bytes)).toEqual([7, 8]);
  });

  it('rejects a symlink inside cwd whose target escapes cwd', async () => {
    await writeFile(join(dir, 'secret-outside.bin'), Buffer.from([9]));
    const sandbox = join(dir, 'sandbox');
    await mkdir(sandbox);
    process.chdir(sandbox);
    await symlink(join(dir, 'secret-outside.bin'), join(sandbox, 'link.bin'));

    await expect(readLocalAttachment('link.bin')).rejects.toThrow(
      /outside the working directory/,
    );
  });

  it('still reads a symlink whose target stays inside cwd', async () => {
    await writeFile('target.bin', Buffer.from([4, 5]));
    await symlink(join(process.cwd(), 'target.bin'), join(process.cwd(), 'alias.bin'));

    const bytes = await readLocalAttachment('alias.bin');
    expect(Array.from(bytes)).toEqual([4, 5]);
  });

  it('rejects ~ home-relative paths (treated as a literal filename, not expanded — still escapes)', async () => {
    // Node doesn't expand ~; resolved against cwd it stays inside. So this *would*
    // try to read a literal `~` file (which doesn't exist). Document that
    // behavior: the guard catches the cross-cwd cases above; `~` is not a
    // shell escape concern here.
    await expect(readLocalAttachment('~')).rejects.toThrow();
  });
});
