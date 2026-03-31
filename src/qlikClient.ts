import { getSpaces as apiGetSpaces } from '@qlik/api/spaces';
import { getItems } from '@qlik/api/items';
import { getAppScript, getAppScriptHistory, updateAppScript } from '@qlik/api/apps';
import type { Space } from '@qlik/api/spaces';
import type { ScriptMeta } from '@qlik/api/apps';
import type { HostConfig } from '@qlik/api/auth';

export type { ScriptMeta };

export type { Space as QlikSpace };

export interface QlikApp {
  id: string;
  name: string;
  spaceId?: string;
  updatedAt?: string;
}

export interface QlikClient {
  getSpaces(): Promise<Space[]>;
  searchApps(spaceIds: string[], query: string): Promise<QlikApp[]>;
  getCurrentScript(appId: string): Promise<string>;
  saveScript(appId: string, script: string, versionMessage: string): Promise<void>;
  getScriptHistory(appId: string): Promise<ScriptMeta[]>;
  getScriptVersion(appId: string, scriptId: string): Promise<string>;
}

export function createClient(hostConfig: HostConfig): QlikClient {
  /** Shared call options: pass hostConfig with every request */
  const opts = () => ({ hostConfig });

  return {
    async getSpaces(): Promise<Space[]> {
      const spaces: Space[] = [];
      let response = await apiGetSpaces({ limit: 100 }, opts());
      spaces.push(...(response.data.data ?? []));
      while (response.next) {
        response = await response.next(opts());
        spaces.push(...(response.data.data ?? []));
      }
      return spaces;
    },

    async searchApps(spaceIds: string[], query: string): Promise<QlikApp[]> {
      const all: QlikApp[] = [];
      const seen = new Set<string>();

      // Empty array → search all spaces (single unfiltered request)
      // '__personal__' → map to 'personal' (supported spaceId value in items API)
      const spaceFilters: Array<string | undefined> =
        spaceIds.length === 0
          ? [undefined]
          : spaceIds.map(id => (id === '__personal__' ? 'personal' : id));

      for (const spaceId of spaceFilters) {
        let response = await getItems(
          {
            resourceType: 'app',
            limit: 100,
            sort: '+name',
            ...(spaceId !== undefined ? { spaceId } : {}),
            ...(query ? { name: query } : {}),
          },
          opts(),
        );

        const collect = (items: typeof response.data.data) => {
          for (const item of items) {
            if (!item.resourceId || seen.has(item.resourceId)) continue;
            seen.add(item.resourceId);
            all.push({
              id: item.resourceId,
              name: item.name,
              spaceId: item.spaceId,
              updatedAt: item.resourceUpdatedAt,
            });
          }
        };

        collect(response.data.data);
        while (response.next) {
          response = await response.next(opts());
          collect(response.data.data);
        }
      }

      return all.sort((a, b) => a.name.localeCompare(b.name));
    },

    async getCurrentScript(appId: string): Promise<string> {
      const response = await getAppScript(appId, 'current', opts());
      return response.data.script ?? '';
    },

    async saveScript(appId: string, script: string, versionMessage: string): Promise<void> {
      await updateAppScript(appId, { script, versionMessage }, opts());
    },

    async getScriptHistory(appId: string): Promise<ScriptMeta[]> {
      const response = await getAppScriptHistory(appId, { limit: 100 }, opts());
      return response.data.scripts ?? [];
    },

    async getScriptVersion(appId: string, scriptId: string): Promise<string> {
      const response = await getAppScript(appId, scriptId, opts());
      return response.data.script ?? '';
    },
  };
}
