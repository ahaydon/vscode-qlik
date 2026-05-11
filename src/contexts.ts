import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import * as yaml from 'js-yaml';
import type { HostConfig } from '@qlik/api/auth';

export type ContextSource = 'yaml' | 'oauth';

export interface QlikContext {
  name: string;
  server: string;
  serverType: string;
  source: ContextSource;
  hostConfig: HostConfig;
  /** For OAuth contexts: the clientId, surfaced so callers can build the runtime hostConfig. */
  clientId?: string;
}

export interface SavedOauthClient {
  name: string;
  host: string;
  clientId: string;
}

export const OAUTH_CLIENTS_KEY = 'qlikcloud.oauthClients';

interface RawContext {
  server?: string;
  'server-type'?: string;
  headers?: Record<string, string>;
  'oauth-client-id'?: string;
  'oauth-client-secret'?: string;
  insecure?: boolean;
}

interface ContextsFile {
  'current-context'?: string;
  contexts?: Record<string, RawContext>;
}

export function loadYamlContexts(): { contexts: QlikContext[]; currentContext: string | undefined } {
  const contextPath = path.join(os.homedir(), '.qlik', 'contexts.yml');
  if (!fs.existsSync(contextPath)) {
    return { contexts: [], currentContext: undefined };
  }

  const raw = yaml.load(fs.readFileSync(contextPath, 'utf8')) as ContextsFile;
  const currentContext = raw['current-context'];
  const contexts: QlikContext[] = [];

  for (const [name, cfg] of Object.entries(raw.contexts ?? {})) {
    if (!cfg.server) continue;

    const serverType = (cfg['server-type'] ?? 'cloud').toLowerCase();
    const host = cfg.server.replace(/\/$/, '');
    let hostConfig: HostConfig | undefined;

    if (cfg.headers?.Authorization) {
      // Bearer token / API key — strip the "Bearer " prefix for apiKey
      const apiKey = cfg.headers.Authorization.replace(/^Bearer\s+/i, '');
      hostConfig = { authType: 'apikey', host, apiKey } as HostConfig;
    } else if (cfg['oauth-client-id'] && cfg['oauth-client-secret']) {
      // OAuth2 client credentials (M2M)
      hostConfig = {
        authType: 'oauth2',
        host,
        clientId: cfg['oauth-client-id'],
        clientSecret: cfg['oauth-client-secret'],
      } as HostConfig;
    } else {
      // No usable auth (e.g. interactive-only OAuth, or empty headers) — skip
      continue;
    }

    contexts.push({ name, server: host, serverType, source: 'yaml', hostConfig });
  }

  return { contexts, currentContext };
}

/**
 * Build placeholder QlikContext entries for each saved OAuth client. The
 * runtime hostConfig (with performInteractiveLogin and accessTokenStorage)
 * must be constructed by the caller using the extension's SecretStorage.
 */
export function oauthClientsToContexts(clients: SavedOauthClient[]): QlikContext[] {
  return clients.map(c => ({
    name: c.name,
    server: c.host,
    serverType: 'cloud',
    source: 'oauth',
    clientId: c.clientId,
    // Placeholder — extension.ts replaces this before creating the client.
    hostConfig: { authType: 'oauth2', host: c.host, clientId: c.clientId } as HostConfig,
  }));
}

export function loadContexts(
  savedOauthClients: SavedOauthClient[] = [],
): { contexts: QlikContext[]; currentContext: string | undefined } {
  const { contexts: yamlContexts, currentContext } = loadYamlContexts();
  return {
    contexts: [...yamlContexts, ...oauthClientsToContexts(savedOauthClients)],
    currentContext,
  };
}
