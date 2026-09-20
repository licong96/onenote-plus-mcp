import { PublicClientApplication, type Configuration, LogLevel } from '@azure/msal-node';
import { getAuthority, getClientId } from '@/config.js';
import { createFileCachePlugin } from './tokenCache.js';

let cachedClient: PublicClientApplication | undefined;

export const getMsalClient = (): PublicClientApplication => {
  if (cachedClient) return cachedClient;

  const config: Configuration = {
    auth: {
      clientId: getClientId(),
      authority: getAuthority(),
    },
    cache: {
      cachePlugin: createFileCachePlugin(),
    },
    system: {
      loggerOptions: {
        // MSAL is chatty; route to stderr so it doesn't clobber MCP stdio.
        loggerCallback: (level, message) => {
          if (level <= LogLevel.Warning) {
            process.stderr.write(`[msal] ${message}\n`);
          }
        },
        piiLoggingEnabled: false,
        logLevel: LogLevel.Warning,
      },
    },
  };

  cachedClient = new PublicClientApplication(config);
  return cachedClient;
};
