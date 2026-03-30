import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import * as yaml from 'js-yaml';
import type { HostConfig } from '@qlik/api/auth';

export interface QlikContext {
  name: string;
  server: string;
  serverType: string;
  hostConfig: HostConfig;
}

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

export function loadContexts(): { contexts: QlikContext[]; currentContext: string | undefined } {
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

    contexts.push({ name, server: host, serverType, hostConfig });
  }

  return { contexts, currentContext };
}
