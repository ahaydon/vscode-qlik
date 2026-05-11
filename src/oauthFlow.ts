import * as http from 'http';
import * as vscode from 'vscode';
import type { HostConfig } from '@qlik/api/auth';
import type { SavedOauthClient } from './contexts';

export const OAUTH_REDIRECT_URI = 'http://localhost:5173/callback';
const LOGIN_TIMEOUT_MS = 5 * 60_000;

interface SecretStorageLike {
  store(key: string, value: string): Promise<void>;
  get(key: string): Promise<string | undefined>;
  delete(key: string): Promise<void>;
}

function secretTopic(host: string, clientId: string): string {
  return `qlikcloud:oauth:${host}:${clientId}`;
}

function makeSecretStorage(secrets: vscode.SecretStorage, topic: string): SecretStorageLike {
  const k = (key: string) => `${topic}:${key}`;
  return {
    store: (key, value) => Promise.resolve(secrets.store(k(key), value)),
    get: key => Promise.resolve(secrets.get(k(key))),
    delete: key => Promise.resolve(secrets.delete(k(key))),
  };
}

async function runLocalCallbackServer(
  redirectUri: string,
  token: vscode.CancellationToken,
): Promise<{ code: string; state: string }> {
  const url = new URL(redirectUri);
  return new Promise((resolve, reject) => {
    const server = http.createServer((req, res) => {
      const reqUrl = new URL(req.url ?? '/', redirectUri);
      if (reqUrl.pathname !== url.pathname) {
        res.writeHead(404).end();
        return;
      }
      const code = reqUrl.searchParams.get('code');
      const state = reqUrl.searchParams.get('state');
      const error = reqUrl.searchParams.get('error');
      const errorDesc = reqUrl.searchParams.get('error_description');

      res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' });
      if (error) {
        res.end(
          `<!doctype html><html><body style="font-family:system-ui;padding:2rem">
            <h3>Login failed</h3>
            <p>${error}${errorDesc ? `: ${errorDesc}` : ''}</p>
            <p>You can close this tab and try again from VS Code.</p>
          </body></html>`,
        );
      } else {
        res.end(
          `<!doctype html><html><body style="font-family:system-ui;padding:2rem">
            <h3>Login complete</h3>
            <p>You can close this tab and return to VS Code.</p>
          </body></html>`,
        );
      }

      server.close();
      if (error) reject(new Error(errorDesc ? `${error}: ${errorDesc}` : error));
      else if (code && state) resolve({ code, state });
      else reject(new Error('Missing code or state in OAuth callback.'));
    });

    server.on('error', err => {
      if ((err as NodeJS.ErrnoException).code === 'EADDRINUSE') {
        reject(new Error(
          `Port ${url.port} is already in use. Close the process using it and try again.`,
        ));
      } else {
        reject(err);
      }
    });

    const port = url.port ? Number(url.port) : 80;
    server.listen(port, url.hostname);

    const timer = setTimeout(() => {
      server.close();
      reject(new Error('OAuth login timed out after 5 minutes.'));
    }, LOGIN_TIMEOUT_MS);
    server.once('close', () => clearTimeout(timer));

    token.onCancellationRequested(() => {
      server.close();
      reject(new Error('OAuth login was cancelled.'));
    });
  });
}

/**
 * Build a HostConfig for a saved OAuth client. The performInteractiveLogin
 * callback will open a browser tab and run a one-shot local HTTP server to
 * capture the redirect.
 */
export function buildOauthHostConfig(
  client: SavedOauthClient,
  secrets: vscode.SecretStorage,
): HostConfig {
  const topic = secretTopic(client.host, client.clientId);
  return {
    authType: 'oauth2',
    host: client.host,
    clientId: client.clientId,
    redirectUri: OAUTH_REDIRECT_URI,
    accessTokenStorage: makeSecretStorage(secrets, topic),
    performInteractiveLogin: async ({ getLoginUrl }) => {
      const loginUrl = await getLoginUrl({ redirectUri: OAUTH_REDIRECT_URI });
      return await vscode.window.withProgress(
        {
          location: vscode.ProgressLocation.Notification,
          title: `Waiting for OAuth login (${client.name})…`,
          cancellable: true,
        },
        async (_progress, token) => {
          const opened = await vscode.env.openExternal(vscode.Uri.parse(loginUrl));
          if (!opened) {
            throw new Error('Failed to open the browser for OAuth login.');
          }
          return await runLocalCallbackServer(OAUTH_REDIRECT_URI, token);
        },
      );
    },
  } as HostConfig;
}

/**
 * Clear cached access/refresh tokens for an OAuth client.
 *
 * The keys here must mirror what @qlik/api writes via its SecretStorage
 * contract: `qlik-qmfe-api-<clientId>_<scope>-{access,refresh}-token`,
 * passed through our adapter's `<topic>:` prefix. We use the default
 * scope (`user_default`) since `buildOauthHostConfig` does not set one.
 */
export async function clearOauthTokens(
  client: SavedOauthClient,
  secrets: vscode.SecretStorage,
): Promise<void> {
  const topic = secretTopic(client.host, client.clientId);
  const libTopic = `${client.clientId}_user_default`;
  await secrets.delete(`${topic}:qlik-qmfe-api-${libTopic}-access-token`);
  await secrets.delete(`${topic}:qlik-qmfe-api-${libTopic}-refresh-token`);
}
