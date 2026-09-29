import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import {
  ILookupOption,
  ISharePointFieldInfo,
  ISharePointItem,
  ISharePointListInfo,
  ISharePointUser,
  RestMethod
} from './SharePointAdminModels';

export default class SharePointAdminService {
  private readonly context: WebPartContext;
  private readonly siteUrl: string;

  public constructor(context: WebPartContext) {
    this.context = context;
    this.siteUrl = context.pageContext.web.absoluteUrl;
  }

  public getSiteUrl(): string {
    return this.siteUrl;
  }

  private async request(
    url: string,
    method: RestMethod = 'GET',
    body?: unknown
  ): Promise<SPHttpClientResponse> {
    const headers: Record<string, string> = {
      Accept: 'application/json;odata=nometadata'
    };

    if (method !== 'GET') {
      headers['Content-Type'] = 'application/json;odata=nometadata';
    }

    if (method === 'MERGE') {
      headers['IF-MATCH'] = '*';
      headers['X-HTTP-Method'] = 'MERGE';
    }

    if (method === 'DELETE') {
      headers['IF-MATCH'] = '*';
      headers['X-HTTP-Method'] = 'DELETE';
    }

    let response: SPHttpClientResponse;

    if (method === 'GET') {
      response = await this.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        { headers }
      );
    } else {
      response = await this.context.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers,
          body: body === undefined ? undefined : JSON.stringify(body)
        }
      );
    }

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`HTTP ${response.status}: ${errorText}`);
    }

    return response;
  }

  private async getJson<T>(url: string): Promise<T> {
    const response = await this.request(url);
    return await response.json() as T;
  }

  private escapeTitle(value: string): string {
    return value.replace(/'/g, "''");
  }

  public async getLists(): Promise<ISharePointListInfo[]> {
    const url =
      `${this.siteUrl}/_api/web/lists` +
      `?$select=Id,Title,ItemCount,Hidden,BaseTemplate` +
      `&$filter=Hidden eq false` +
      `&$orderby=Title asc`;

    const data = await this.getJson<{ value: ISharePointListInfo[] }>(url);

    return data.value.filter(list =>
      list.BaseTemplate === 100 ||
      list.BaseTemplate === 101
    );
  }

  public async getFields(listTitle: string): Promise<ISharePointFieldInfo[]> {
    const title = this.escapeTitle(listTitle);

    const url =
      `${this.siteUrl}/_api/web/lists/getbytitle('${title}')/fields` +
      `?$select=Title,InternalName,TypeAsString,Hidden,ReadOnlyField,Required,Choices,LookupList,LookupField`;

    const data =
      await this.getJson<{ value: ISharePointFieldInfo[] }>(url);

    return data.value.filter(field =>
      !field.Hidden &&
      field.InternalName !== 'Attachments' &&
      field.InternalName !== 'ContentType'
    );
  }

  public async getItems(
  listTitle: string,
  fields: ISharePointFieldInfo[],
  top: number = 200
): Promise<ISharePointItem[]> {

  const title =
    this.escapeTitle(listTitle);

  const select: string[] = [
    'Id'
  ];

  const camposSistemaIgnorados: string[] = [
    '_ColorTag',
    '_UIVersionString',
    'Edit',
    'LinkTitle',
    'LinkTitleNoMenu',
    'LinkFilename',
    'LinkFilenameNoMenu',
    'DocIcon',
    'ItemChildCount',
    'FolderChildCount',
    'AppAuthor',
    'AppEditor',
    'ContentType',
    'Attachments'
  ];

  fields.forEach(field => {

    if (
      field.InternalName === 'ID' ||
      field.InternalName === 'Id' ||
      field.Hidden ||
      field.ReadOnlyField ||
      field.TypeAsString === 'Computed' ||
      camposSistemaIgnorados.indexOf(
        field.InternalName
      ) >= 0
    ) {
      return;
    }

    if (
      field.TypeAsString === 'User' ||
      field.TypeAsString === 'UserMulti' ||
      field.TypeAsString === 'Lookup' ||
      field.TypeAsString === 'LookupMulti'
    ) {

      select.push(
        `${field.InternalName}Id`
      );

      return;
    }

    select.push(
      field.InternalName
    );
  });

  const uniqueSelect =
    select.filter(
      (
        value,
        index,
        array
      ) => array.indexOf(value) === index
    );

  const url =
    `${this.siteUrl}` +
    `/_api/web/lists/getbytitle('${title}')/items` +
    `?$select=${uniqueSelect.join(',')}` +
    `&$orderby=Id desc` +
    `&$top=${top}`;

  const data =
    await this.getJson<{
      value: ISharePointItem[];
    }>(url);

  return data.value;
}

  public async createItem(
    listTitle: string,
    body: Record<string, unknown>
  ): Promise<void> {
    const title = this.escapeTitle(listTitle);

    await this.request(
      `${this.siteUrl}/_api/web/lists/getbytitle('${title}')/items`,
      'POST',
      body
    );
  }

  public async updateItem(
    listTitle: string,
    itemId: number,
    body: Record<string, unknown>
  ): Promise<void> {
    const title = this.escapeTitle(listTitle);

    await this.request(
      `${this.siteUrl}/_api/web/lists/getbytitle('${title}')/items(${itemId})`,
      'MERGE',
      body
    );
  }

  public async deleteItem(
    listTitle: string,
    itemId: number
  ): Promise<void> {
    const title = this.escapeTitle(listTitle);

    await this.request(
      `${this.siteUrl}/_api/web/lists/getbytitle('${title}')/items(${itemId})`,
      'DELETE'
    );
  }

  public async getUsers(): Promise<ISharePointUser[]> {
    const url =
      `${this.siteUrl}/_api/web/siteusers` +
      `?$select=Id,Title,Email,LoginName` +
      `&$orderby=Title asc` +
      `&$top=5000`;

    const data =
      await this.getJson<{ value: ISharePointUser[] }>(url);

    return data.value.filter(user => !!user.Title);
  }

  public async getLookupOptions(
    field: ISharePointFieldInfo
  ): Promise<ILookupOption[]> {
    if (!field.LookupList || !field.LookupField) {
      return [];
    }

    const lookupList = field.LookupList.replace(/[{}]/g, '');
    const lookupField = field.LookupField;

    const url =
      `${this.siteUrl}/_api/web/lists(guid'${lookupList}')/items` +
      `?$select=Id,${encodeURIComponent(lookupField)}` +
      `&$orderby=${encodeURIComponent(lookupField)} asc` +
      `&$top=5000`;

    const data =
      await this.getJson<{ value: ISharePointItem[] }>(url);

    return data.value.map(item => ({
      Id: item.Id,
      Label: String(item[lookupField] || `ID ${item.Id}`)
    }));
  }

  public async executeRest(
    endpoint: string,
    method: RestMethod,
    bodyText: string
  ): Promise<{ status: number; data: unknown }> {
    const endpointNormalizado =
      endpoint.indexOf('http') === 0
        ? endpoint
        : `${this.siteUrl}${endpoint.indexOf('/') === 0 ? '' : '/'}${endpoint}`;

    let body: unknown = undefined;

    if (bodyText.trim()) {
      body = JSON.parse(bodyText);
    }

    const response = await this.request(
      endpointNormalizado,
      method,
      body
    );

    const text = await response.text();

    let data: unknown = text;

    if (text) {
      try {
        data = JSON.parse(text);
      } catch {
        data = text;
      }
    }

    return {
      status: response.status,
      data
    };
  }
}
