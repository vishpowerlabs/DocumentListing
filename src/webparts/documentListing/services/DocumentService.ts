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
  [key: string]: unknown; // Allow dynamic fields
}

export interface IListInfo {
  Id: string;
  Title: string;
}

export interface IFieldInfo {
  InternalName: string;
  Title: string;
  ReadOnlyField: boolean;
  TypeAsString: string;
}

export default class DocumentService {
  constructor(
    private readonly spHttpClient: SPHttpClient,
    private readonly siteUrl: string
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
      url = `${this.siteUrl}/_api/web/lists(guid'${listId}')/Fields?$filter=Hidden eq false or CanBeDeleted eq true&$select=Title,InternalName,ReadOnlyField,TypeAsString`;
    } else {
      url = `${this.siteUrl}/_api/web/lists/getbytitle('${listId}')/Fields?$filter=Hidden eq false or CanBeDeleted eq true&$select=Title,InternalName,ReadOnlyField,TypeAsString`;
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

  public async createRequest(listId: string, itemData: Record<string, unknown>): Promise<void> {
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

  public async getExistingRequest(
    listId: string,
    fileIdCol: string,
    fileIdVal: string,
    emailCol: string,
    emailVal: string,
    selectCols: string[]
  ): Promise<Record<string, unknown> | undefined> {
    const isGuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(listId);

    // Filter by File ID and Email
    // Note: Assuming File ID is Text. If Number, remove quotes around fileIdVal.
    // Usually standard practice is to use Text for "Ref ID" columns to be safe, but let's assume Text for now as per previous implementation.
    const filter = `${fileIdCol} eq '${fileIdVal}' and ${emailCol} eq '${emailVal}'`;
    const select = selectCols.join(',');

    let url = '';
    if (isGuid) {
      url = `${this.siteUrl}/_api/web/lists(guid'${listId}')/items?$filter=${filter}&$select=Id,${select}`;
    } else {
      url = `${this.siteUrl}/_api/web/lists/getbytitle('${listId}')/items?$filter=${filter}&$select=Id,${select}`;
    }

    const response: SPHttpClientResponse = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const json = await response.json();

    return (json.value && json.value.length > 0) ? json.value[0] : undefined;
  }

  public async updateRequest(listId: string, itemId: number, itemData: Record<string, unknown>): Promise<void> {
    const isGuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(listId);
    let url = '';

    if (isGuid) {
      url = `${this.siteUrl}/_api/web/lists(guid'${listId}')/items(${itemId})`;
    } else {
      url = `${this.siteUrl}/_api/web/lists/getbytitle('${listId}')/items(${itemId})`;
    }

    const body = JSON.stringify(itemData);

    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'OData-Version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: body
    };

    const response: SPHttpClientResponse = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, options);

    if (!response.ok) {
      // Try to parse error if possible, or throw status text
      try {
        const error = await response.json();
        throw new Error(error.error ? error.error.message.value : response.statusText);
      } catch {
        throw new Error(response.statusText);
      }
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

    return json.value.map((i: Record<string, unknown>) => {
      const item: IDocumentItem = {
        Title: i.Title as string,
        Category: i[categoryColumn] as string,
        SubCategory: i[subCategoryColumn] as string,
        FileRef: i.FileRef as string,
        Modified: i.Modified as string,
        Id: i.Id as number
      };

      // Map dynamic fields
      columnsToSelect.forEach(f => {
        item[f] = i[f];
      });

      return item;
    });
  }
}