import { unlink } from 'node:fs/promises';
import type { AuthenticationResult, AccountInfo } from '@azure/msal-node';
import { SCOPES, getTokenCachePath } from '@/config.js';
import { getMsalClient } from './msal.js';

export type DeviceCodeNotifier = (info: {
  userCode: string;
  verificationUri: string;
  message: string;
  expiresIn: number;
}) => void;

const DEFAULT_NOTIFIER: DeviceCodeNotifier = ({ message }) => {
  process.stderr.write(`${message}\n`);
};

const getCachedAccount = async (): Promise<AccountInfo | undefined> => {
  const cache = getMsalClient().getTokenCache();
  const accounts = await cache.getAllAccounts();
  return accounts[0];
};

const acquireTokenSilently = async (
  account: AccountInfo,
): Promise<AuthenticationResult | null> => {
  try {
    return await getMsalClient().acquireTokenSilent({
      account,
      scopes: [...SCOPES],
    });
  } catch {
    return null;
  }
};

const acquireTokenByDeviceCode = async (
  notify: DeviceCodeNotifier,
): Promise<AuthenticationResult> => {
  const result = await getMsalClient().acquireTokenByDeviceCode({
    scopes: [...SCOPES],
    deviceCodeCallback: (response) => {
      notify({
        userCode: response.userCode,
        verificationUri: response.verificationUri,
        message: response.message,
        expiresIn: response.expiresIn,
      });
    },
  });
  if (!result) {
    throw new Error('Device-code authentication returned no result.');
  }
  return result;
};

export const getAccessToken = async (): Promise<string> => {
  const account = await getCachedAccount();
  if (account) {
    const silent = await acquireTokenSilently(account);
    if (silent?.accessToken) return silent.accessToken;
  }
  throw new Error(
    'Not signed in. Run `onenote-mcp login` to authenticate before starting the MCP server.',
  );
};

export const login = async (
  notify: DeviceCodeNotifier = DEFAULT_NOTIFIER,
): Promise<AccountInfo> => {
  const existing = await getCachedAccount();
  if (existing) {
    const silent = await acquireTokenSilently(existing);
    if (silent?.account) return silent.account;
  }
  const result = await acquireTokenByDeviceCode(notify);
  if (!result.account) {
    throw new Error('Sign-in completed but no account was returned.');
  }
  return result.account;
};

export const logout = async (): Promise<void> => {
  const cache = getMsalClient().getTokenCache();
  for (const account of await cache.getAllAccounts()) {
    await cache.removeAccount(account);
  }
  try {
    await unlink(getTokenCachePath());
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code !== 'ENOENT') throw err;
  }
};
