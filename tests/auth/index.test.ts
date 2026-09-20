import { describe, it, expect, vi, beforeEach } from 'vitest';
import type { AccountInfo, AuthenticationResult } from '@azure/msal-node';

const mockGetAllAccounts = vi.fn();
const mockRemoveAccount = vi.fn();
const mockAcquireTokenSilent = vi.fn();
const mockAcquireTokenByDeviceCode = vi.fn();

vi.mock('../../src/auth/msal.js', () => ({
  getMsalClient: () => ({
    getTokenCache: () => ({
      getAllAccounts: mockGetAllAccounts,
      removeAccount: mockRemoveAccount,
    }),
    acquireTokenSilent: mockAcquireTokenSilent,
    acquireTokenByDeviceCode: mockAcquireTokenByDeviceCode,
  }),
}));

vi.mock('node:fs/promises', () => ({
  unlink: vi.fn(),
}));

vi.mock('../../src/config.js', () => ({
  SCOPES: ['Notes.ReadWrite', 'offline_access'],
  getTokenCachePath: () => '/fake/.config/onenote-mcp/tokens.json',
}));

import { getAccessToken, login, logout } from '@/auth/index.js';
import { unlink } from 'node:fs/promises';

const fakeAccount: AccountInfo = {
  homeAccountId: 'home-id',
  environment: 'login.microsoftonline.com',
  tenantId: 'tenant-id',
  username: 'user@example.com',
  localAccountId: 'local-id',
};

const fakeAuthResult = (
  overrides: Partial<AuthenticationResult> = {},
): AuthenticationResult => ({
  authority: 'https://login.microsoftonline.com/common',
  uniqueId: 'unique-id',
  tenantId: 'tenant-id',
  scopes: ['Notes.ReadWrite', 'offline_access'],
  account: fakeAccount,
  idToken: 'id-token',
  idTokenClaims: {},
  accessToken: 'access-token-abc',
  fromCache: false,
  expiresOn: new Date(Date.now() + 3600_000),
  tokenType: 'Bearer',
  correlationId: 'corr-id',
  ...overrides,
});

beforeEach(() => {
  vi.clearAllMocks();
});

describe('getAccessToken', () => {
  it('returns the access token when a cached account acquires silently', async () => {
    mockGetAllAccounts.mockResolvedValueOnce([fakeAccount]);
    mockAcquireTokenSilent.mockResolvedValueOnce(fakeAuthResult());

    const token = await getAccessToken();

    expect(token).toBe('access-token-abc');
    expect(mockAcquireTokenSilent).toHaveBeenCalledWith({
      account: fakeAccount,
      scopes: ['Notes.ReadWrite', 'offline_access'],
    });
  });

  it('throws when no cached account exists', async () => {
    mockGetAllAccounts.mockResolvedValueOnce([]);

    await expect(getAccessToken()).rejects.toThrow('Not signed in');
  });

  it('throws when silent acquisition fails', async () => {
    mockGetAllAccounts.mockResolvedValueOnce([fakeAccount]);
    mockAcquireTokenSilent.mockRejectedValueOnce(new Error('interaction_required'));

    await expect(getAccessToken()).rejects.toThrow('Not signed in');
  });

  it('throws when silent acquisition returns null', async () => {
    mockGetAllAccounts.mockResolvedValueOnce([fakeAccount]);
    mockAcquireTokenSilent.mockResolvedValueOnce(null);

    await expect(getAccessToken()).rejects.toThrow('Not signed in');
  });

  it('throws when silent acquisition returns a result with no accessToken', async () => {
    mockGetAllAccounts.mockResolvedValueOnce([fakeAccount]);
    mockAcquireTokenSilent.mockResolvedValueOnce(fakeAuthResult({ accessToken: '' }));

    await expect(getAccessToken()).rejects.toThrow('Not signed in');
  });
});

describe('login', () => {
  it('returns the cached account when silent acquisition succeeds', async () => {
    mockGetAllAccounts.mockResolvedValueOnce([fakeAccount]);
    mockAcquireTokenSilent.mockResolvedValueOnce(fakeAuthResult());

    const account = await login();

    expect(account).toEqual(fakeAccount);
    expect(mockAcquireTokenByDeviceCode).not.toHaveBeenCalled();
  });

  it('falls back to device code when no cached account exists', async () => {
    mockGetAllAccounts.mockResolvedValueOnce([]);
    mockAcquireTokenByDeviceCode.mockResolvedValueOnce(fakeAuthResult());

    const notify = vi.fn();
    const account = await login(notify);

    expect(account).toEqual(fakeAccount);
    expect(mockAcquireTokenByDeviceCode).toHaveBeenCalledTimes(1);
  });

  it('falls back to device code when silent acquisition fails', async () => {
    mockGetAllAccounts.mockResolvedValueOnce([fakeAccount]);
    mockAcquireTokenSilent.mockRejectedValueOnce(new Error('expired'));
    mockAcquireTokenByDeviceCode.mockResolvedValueOnce(fakeAuthResult());

    const account = await login();

    expect(account).toEqual(fakeAccount);
    expect(mockAcquireTokenByDeviceCode).toHaveBeenCalledTimes(1);
  });

  it('falls back to device code when silent acquisition returns null account', async () => {
    mockGetAllAccounts.mockResolvedValueOnce([fakeAccount]);
    mockAcquireTokenSilent.mockResolvedValueOnce(
      fakeAuthResult({ account: null as unknown as AccountInfo }),
    );
    mockAcquireTokenByDeviceCode.mockResolvedValueOnce(fakeAuthResult());

    const account = await login();

    expect(account).toEqual(fakeAccount);
    expect(mockAcquireTokenByDeviceCode).toHaveBeenCalledTimes(1);
  });

  it('invokes the device code callback with structured info', async () => {
    mockGetAllAccounts.mockResolvedValueOnce([]);
    mockAcquireTokenByDeviceCode.mockImplementationOnce(
      async (opts: { deviceCodeCallback: (info: unknown) => void }) => {
        opts.deviceCodeCallback({
          userCode: 'ABC123',
          verificationUri: 'https://microsoft.com/devicelogin',
          message: 'Go to the page and enter the code',
          expiresIn: 900,
        });
        return fakeAuthResult();
      },
    );

    const notify = vi.fn();
    await login(notify);

    expect(notify).toHaveBeenCalledWith({
      userCode: 'ABC123',
      verificationUri: 'https://microsoft.com/devicelogin',
      message: 'Go to the page and enter the code',
      expiresIn: 900,
    });
  });

  it('throws when device code returns no result', async () => {
    mockGetAllAccounts.mockResolvedValueOnce([]);
    mockAcquireTokenByDeviceCode.mockResolvedValueOnce(null);

    await expect(login()).rejects.toThrow('no result');
  });

  it('throws when device code returns a result with no account', async () => {
    mockGetAllAccounts.mockResolvedValueOnce([]);
    mockAcquireTokenByDeviceCode.mockResolvedValueOnce(
      fakeAuthResult({ account: null as unknown as AccountInfo }),
    );

    await expect(login()).rejects.toThrow('no account was returned');
  });
});

describe('logout', () => {
  it('removes all cached accounts and deletes the token file', async () => {
    const account2: AccountInfo = { ...fakeAccount, homeAccountId: 'other' };
    mockGetAllAccounts.mockResolvedValueOnce([fakeAccount, account2]);

    await logout();

    expect(mockRemoveAccount).toHaveBeenCalledTimes(2);
    expect(mockRemoveAccount).toHaveBeenCalledWith(fakeAccount);
    expect(mockRemoveAccount).toHaveBeenCalledWith(account2);
    expect(unlink).toHaveBeenCalledWith('/fake/.config/onenote-mcp/tokens.json');
  });

  it('handles an already-deleted token file gracefully', async () => {
    mockGetAllAccounts.mockResolvedValueOnce([]);
    const err = Object.assign(new Error('ENOENT'), { code: 'ENOENT' });
    vi.mocked(unlink).mockRejectedValueOnce(err);

    await expect(logout()).resolves.toBeUndefined();
  });

  it('re-throws non-ENOENT file system errors', async () => {
    mockGetAllAccounts.mockResolvedValueOnce([]);
    const err = Object.assign(new Error('EACCES'), { code: 'EACCES' });
    vi.mocked(unlink).mockRejectedValueOnce(err);

    await expect(logout()).rejects.toThrow('EACCES');
  });

  it('works when there are no cached accounts', async () => {
    mockGetAllAccounts.mockResolvedValueOnce([]);

    await logout();

    expect(mockRemoveAccount).not.toHaveBeenCalled();
    expect(unlink).toHaveBeenCalledTimes(1);
  });
});
