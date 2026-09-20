import { stat } from 'node:fs/promises';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { getAuthSnapshot } from '@/auth/index.js';
import {
  CONFIG_DIR_ENV,
  CLIENT_ID_ENV,
  SCOPES,
  getAuthority,
  getConfigDir,
  getTokenCachePath,
} from '@/config.js';

const fileExists = async (path: string): Promise<boolean> => {
  try {
    await stat(path);
    return true;
  } catch {
    return false;
  }
};

export const register = (server: McpServer): void => {
  server.registerTool(
    'auth_status',
    {
      title: 'Auth Status',
      description:
        'Reports which Microsoft account is signed in, the token cache location, and whether the cached token still works. Use it first when a OneNote call fails with an authentication error.',
      inputSchema: {},
    },
    async () => {
      const [snapshot, cacheExists] = await Promise.all([
        getAuthSnapshot(),
        fileExists(getTokenCachePath()),
      ]);

      const payload = {
        signedIn: snapshot.signedIn,
        account: snapshot.account
          ? {
              username: snapshot.account.username,
              name: snapshot.account.name,
              tenantId: snapshot.account.tenantId,
              environment: snapshot.account.environment,
            }
          : null,
        accessTokenExpiresOn: snapshot.expiresOn?.toISOString(),
        // A client ID is a public identifier, not a secret — safe to report.
        clientId: process.env[CLIENT_ID_ENV]?.trim() || null,
        clientIdEnvVar: CLIENT_ID_ENV,
        authority: getAuthority(),
        scopes: [...SCOPES],
        tokenCache: {
          path: getTokenCachePath(),
          exists: cacheExists,
          configDirOverrideEnvVar: CONFIG_DIR_ENV,
          configDirOverride: process.env[CONFIG_DIR_ENV]?.trim() || null,
          configDir: getConfigDir(),
        },
        remediation: snapshot.signedIn
          ? undefined
          : 'Not signed in, or the cached token could not be refreshed. Run `onenote-plus-mcp login` in a terminal, then restart this MCP server.',
      };

      return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] };
    },
  );
};
