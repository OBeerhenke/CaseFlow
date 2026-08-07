import { sp } from '@pnp/sp';
import { IConfigStore } from './ConfigService';

interface IConfigListItem {
  ID: number;
  Title: string;
  Value: string;
}

/**
 * SharePoint-backed implementation of IConfigStore, reading/writing the
 * CaseFlow_Config key-value list. Kept in its own file (separate from
 * ConfigService itself) so ConfigService — and anything that tests it —
 * never needs to import `@pnp/sp`.
 */
export class SharePointConfigStore implements IConfigStore {
  public constructor(private readonly listName: string) {}

  public async getAll(): Promise<Record<string, string>> {
    const items: IConfigListItem[] = await sp.web.lists.getByTitle(this.listName).items
      .select('ID', 'Title', 'Value')
      .top(2000)
      .get();

    const result: Record<string, string> = {};
    for (const item of items) {
      result[item.Title] = item.Value;
    }
    return result;
  }

  public async setValue(key: string, value: string): Promise<void> {
    const list = sp.web.lists.getByTitle(this.listName);
    const escapedKey = key.replace(/'/g, "''");
    const existing: IConfigListItem[] = await list.items
      .filter(`Title eq '${escapedKey}'`)
      .select('ID')
      .top(1)
      .get();

    if (existing.length > 0) {
      await list.items.getById(existing[0].ID).update({ Value: value });
    } else {
      await list.items.add({ Title: key, Value: value });
    }
  }
}
