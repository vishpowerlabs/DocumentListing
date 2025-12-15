import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';

export interface IDocumentItem {
  Title: string;
  Description?: string;
  Category: string;
  SubCategory: string;
  FileRef: string;
  Modified: string;
  Id: number;
  [key: string]: any; // Allow dynamic fields
}

export interface IListInfo {
  Id: string;
  Title: string;
}

export interface IFieldInfo {
  InternalName: string;
  Title: string;
}

export default class DocumentService {
  constructor(
    private spHttpClient: SPHttpClient,
    private siteUrl: string
  ) { }

  public async getLists(baseTemplate: number = 101): Promise<IListInfo[]> {
    const url = `${this.siteUrl}/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq ${baseTemplate}&$select=Id,Title`;

    const response: SPHttpClientResponse = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const json = await response.json();

    return json.value as IListInfo[];
  }

  public async getColumns(listId: string): Promise<IFieldInfo[]> {
    const isGuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(listId);
    let url = '';

    if (isGuid) {
      url = `${this.siteUrl}/_api/web/lists(guid'${listId}')/Fields?$filter=Hidden eq false or CanBeDeleted eq true&$select=Title,InternalName,ReadOnlyField`;
    } else {
      url = `${this.siteUrl}/_api/web/lists/getbytitle('${listId}')/Fields?$filter=Hidden eq false or CanBeDeleted eq true&$select=Title,InternalName,ReadOnlyField`;
    }

    const response: SPHttpClientResponse = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const json = await response.json();

    return json.value as IFieldInfo[];
  }

  public async getFieldChoices(listId: string, fieldInternalName: string): Promise<string[]> {
    if (!fieldInternalName) return [];

    const isGuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(listId);
    let url = '';

    if (isGuid) {
      url = `${this.siteUrl}/_api/web/lists(guid'${listId}')/Fields?$filter=InternalName eq '${fieldInternalName}'`;
    } else {
      url = `${this.siteUrl}/_api/web/lists/getbytitle('${listId}')/Fields?$filter=InternalName eq '${fieldInternalName}'`;
    }

    try {
      const response: SPHttpClientResponse = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const json = await response.json();

      if (json.value && json.value.length > 0) {
        return json.value[0].Choices || [];
      }
      return [];
    } catch (e) {
      console.error(`Error fetching choices for ${fieldInternalName}`, e);
      return [];
    }
  }

  public async createRequest(listId: string, itemData: any): Promise<void> {
    const url = `${this.siteUrl}/_api/web/lists(guid'${listId}')/items`;

    const body = JSON.stringify(itemData);

    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'OData-Version': ''
      },
      body: body
    };

    const response: SPHttpClientResponse = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, options);

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error ? error.error.message.value : response.statusText);
    }
  }

  public async getDocuments(
    library: string,
    categoryColumn: string,
    subCategoryColumn: string,
    columnsToSelect: string[] // List of internal names to fetch
  ): Promise<IDocumentItem[]> {
    // Check if library is a GUID (simple check)
    const isGuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(library);

    // Build select string
    let selects = ['Title', 'FileRef', 'Modified', 'Id'];
    if (categoryColumn) selects.push(categoryColumn);
    if (subCategoryColumn) selects.push(subCategoryColumn);
    if (columnsToSelect && columnsToSelect.length > 0) {
      selects = selects.concat(columnsToSelect);
    }
    // Remove duplicates
    selects = selects.filter((item, pos) => selects.indexOf(item) === pos);

    // Build select string
    const selectStr = selects.join(',');

    let url = '';
    if (isGuid) {
      url = `${this.siteUrl}/_api/web/lists(guid'${library}')/items?$select=${selectStr}`;
    } else {
      // Fallback for existing configuration (Title)
      url = `${this.siteUrl}/_api/web/lists/getbytitle('${library}')/items?$select=${selectStr}`;
    }

    const response: SPHttpClientResponse =
      await this.spHttpClient.get(url, SPHttpClient.configurations.v1);

    const json = await response.json();

    return json.value.map((i: any) => {
      const item: IDocumentItem = {
        Title: i.Title,
        Category: i[categoryColumn],
        SubCategory: i[subCategoryColumn],
        FileRef: i.FileRef,
        Modified: i.Modified,
        Id: i.Id
      };

      // Map dynamic fields
      columnsToSelect.forEach(f => {
        item[f] = i[f];
      });

      return item;
    });
  }
}