import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';

import {
  ThemeProvider,
  IReadonlyTheme,
  ThemeChangedEventArgs
} from '@microsoft/sp-component-base';

import styles from './DocumentListingWebPart.module.scss';
import { IDocumentListingWebPartProps } from './IDocumentListingWebPartProps';
import DocumentService, { IDocumentItem } from './services/DocumentService';

export default class DocumentListingWebPart
  extends BaseClientSideWebPart<IDocumentListingWebPartProps> {

  private service!: DocumentService;
  private themeProvider!: ThemeProvider;
  private themeVariant: IReadonlyTheme | undefined;
  private items: IDocumentItem[] = [];
  private currentPage: number = 1;

  private lists: IPropertyPaneDropdownOption[] = [];
  private requestLists: IPropertyPaneDropdownOption[] = []; // For generic lists (requests)
  private columns: IPropertyPaneDropdownOption[] = [];
  private requestColumns: IPropertyPaneDropdownOption[] = []; // Columns for the separate request list
  private listsDropdownDisabled: boolean = true;
  private requestListsDropdownDisabled: boolean = true;
  private columnsDropdownDisabled: boolean = true;
  private requestColumnsDropdownDisabled: boolean = true;

  public async onInit(): Promise<void> {
    this.themeProvider = this.context.serviceScope.consume(
      ThemeProvider.serviceKey
    );

    this.themeVariant = this.themeProvider.tryGetTheme();
    this.themeProvider.themeChangedEvent.add(this, this._handleThemeChanged);

    this.service = new DocumentService(
      this.context.spHttpClient,
      this.context.pageContext.web.absoluteUrl
    );
  }

  private _handleThemeChanged(args: ThemeChangedEventArgs): void {
    this.themeVariant = args.theme;
    this.render();
  }

  public async render(): Promise<void> {
    if (!this.properties.library) {
      this.domElement.innerHTML = `<p>Please configure the web part.</p>`;
      return;
    }

    this.applyTheme();

    try {
      // Collect all columns we need to fetch
      const columnsToFetch = [
        this.properties.titleColumn,
        this.properties.descriptionColumn
      ].filter(f => f); // Filter out empty

      this.items = await this.service.getDocuments(
        this.properties.library,
        this.properties.categoryColumn,
        this.properties.subCategoryColumn,
        columnsToFetch
      );

      // Fetch choices for Categories if a column is selected
      let categoryChoices: string[] = [];
      if (this.properties.categoryColumn) {
        categoryChoices = await this.service.getFieldChoices(
          this.properties.library,
          this.properties.categoryColumn
        );
      }

      const categories = categoryChoices.length > 0 ? categoryChoices : [];

      if (categoryChoices.length === 0) {
        this.items.forEach(i => {
          if (categories.indexOf(i.Category) === -1 && i.Category) {
            categories.push(i.Category);
          }
        });
        categories.sort();
      }

      this.domElement.innerHTML = `
        <div class="${styles.container}">
          <div class="${styles.leftNav}">
            ${categories.map(c => `
              <div class="${styles.categoryItem}" data-category="${c}">
                ${c}
              </div>
            `).join('')}
          </div>

          <div class="${styles.content}">
            <div class="${styles.subCategoryTabs}" id="subTabs"></div>

            <div class="${styles.tableWrapper}">
              <table class="${styles.table}">
                <thead>
                  <tr>
                    <th>Title</th>
                    <th>Description</th>
                    <th class="${styles.centerAlign}">Date Modified</th>
                    <th class="${styles.centerAlign}">Request Access</th>
                  </tr>
                </thead>
                <tbody id="docRows"></tbody>
              </table>
            </div>
          </div>
        </div>
      `;

      this.bindCategoryEvents();

      // Auto-select first category if available
      if (categories.length > 0) {
        // Highlight first nav item
        const firstCatEl = this.domElement.querySelector(`.${styles.categoryItem}[data-category="${categories[0]}"]`);
        if (firstCatEl) {
          firstCatEl.classList.add(styles.categoryActive);
          this.loadSubCategories(categories[0]);
        }
      }

    } catch (error) {
      this.domElement.innerHTML = `<p>Error loading documents: ${error}</p>`;
    }
  }

  private applyTheme(): void {
    const semantic = this.themeVariant?.semanticColors;
    const palette = this.themeVariant?.palette;

    this.domElement.style.setProperty('--bodyText', semantic?.bodyText || '#323130');
    this.domElement.style.setProperty('--bodyBackground', semantic?.bodyBackground || '#ffffff');
    this.domElement.style.setProperty('--neutralLight', palette?.neutralLight || '#edebe9');
    this.domElement.style.setProperty('--neutralLighter', palette?.neutralLighter || '#f3f2f1');
    this.domElement.style.setProperty('--themePrimary', palette?.themePrimary || '#0078d4');
  }

  private bindCategoryEvents(): void {
    this.domElement.querySelectorAll(`.${styles.categoryItem}`)
      .forEach(el => {
        el.addEventListener('click', e => {
          // Remove active from all
          this.domElement.querySelectorAll(`.${styles.categoryItem}`).forEach(c => c.classList.remove(styles.categoryActive));
          // Add active to current
          (e.currentTarget as HTMLElement).classList.add(styles.categoryActive);

          const category = (e.currentTarget as HTMLElement).dataset.category!;
          this.loadSubCategories(category);
        });
      });
  }

  private async loadSubCategories(category: string): Promise<void> {
    // Reset page logic when switching main category
    // (Actual reset happens in call to renderTable, but good to be explicit if shared state)
    this.currentPage = 1;
    // Fetch choices for SubCategories if a column is selected
    let subChoices: string[] = [];
    if (this.properties.subCategoryColumn) {
      subChoices = await this.service.getFieldChoices(
        this.properties.library,
        this.properties.subCategoryColumn
      );
    }

    const subs = subChoices.length > 0 ? subChoices : [];

    if (subChoices.length === 0) {
      this.items.forEach(i => {
        if (i.Category === category && subs.indexOf(i.SubCategory) === -1 && i.SubCategory) {
          subs.push(i.SubCategory);
        }
      });
      subs.sort();
    }

    // If fetching choices, we might list subcategories that don't have items in this category.
    // The requirement is "category and sub category show based on field choice insted of checking based on the records"
    // However, if we show ALL subcategories for every category, that might be n*m combinations.
    // Usually SubCategory is filtered by what's available?
    // BUT the user asked specifically "instead of checking based on records".
    // If they are independent columns, then yes, we show all choices.
    // If they are related lookup, `getFieldChoices` might not be enough.
    // Assuming independent choice columns for now as per standard SP setup.

    const tabs = subs.map((s, i) => `
      <div class="${styles.subTab} ${i === 0 ? styles.active : ''}"
           data-sub="${s}">
        ${s}
      </div>
    `).join('');

    this.domElement.querySelector('#subTabs')!.innerHTML = tabs;

    this.bindSubCategoryEvents(category);

    // Auto-select first subcategory (already marked active in HTML above)
    if (subs.length > 0) {
      this.currentPage = 1;
      this.renderTable(category, subs[0]);
    } else {
      this.currentPage = 1;
      this.renderTable(category, '');
    }
  }

  private bindSubCategoryEvents(category: string): void {
    this.domElement.querySelectorAll(`.${styles.subTab}`)
      .forEach(tab => {
        tab.addEventListener('click', e => {
          this.domElement.querySelectorAll(`.${styles.subTab}`)
            .forEach(t => t.classList.remove(styles.active));

          const el = e.currentTarget as HTMLElement;
          el.classList.add(styles.active);
          this.currentPage = 1; // Reset to page 1 on sub change
          this.renderTable(category, el.dataset.sub!);
        });
      });
  }

  private renderTable(category: string, sub: string): void {
    const allFilteredItems = this.items
      .filter(i => i.Category === category && i.SubCategory === sub);

    // Pagination Logic
    const pageSize = this.properties.pageSize || 10; // default 10
    const totalPages = Math.ceil(allFilteredItems.length / pageSize);

    // Ensure current page is valid
    if (this.currentPage > totalPages) this.currentPage = 1;
    if (this.currentPage < 1) this.currentPage = 1;

    const startIndex = (this.currentPage - 1) * pageSize;
    const endIndex = startIndex + pageSize;
    const pagedItems = allFilteredItems.slice(startIndex, endIndex);

    const rows = pagedItems.map(i => {
      // Map fields
      const title = this.properties.titleColumn ? i[this.properties.titleColumn] : i.Title;
      const desc = this.properties.descriptionColumn ? i[this.properties.descriptionColumn] : (i.Description || '');

      // Hardcoded Date Modified
      let dateStr = '';
      if (i.Modified) {
        dateStr = new Date(i.Modified).toLocaleDateString();
      }

      // Hardcoded Request Access Icon logic with click handler
      // We render a button-like span
      return `
        <tr>
          <td>${title || ''}</td>
          <td>${desc}</td>
          <td class="${styles.centerAlign}">${dateStr}</td>
          <td class="${styles.centerAlign}">
            <span class="${styles.mailIcon} request-access-btn" data-id="${i.Id}" title="Request Access" style="cursor: pointer;">✉</span>
          </td>
        </tr>
      `}).join('');

    // Render Table Body
    const tbody = this.domElement.querySelector('#docRows');
    if (tbody) tbody.innerHTML = rows;

    // Bind Request Access Buttons
    this.domElement.querySelectorAll('.request-access-btn').forEach(btn => {
      btn.addEventListener('click', (e) => this.handleRequestAccess(e));
    });

    // Render Pagination Controls
    // We need to inject pagination HTML if it doesn't exist, or update it
    let pagContainer = this.domElement.querySelector(`.${styles.paginationContainer}`);
    if (!pagContainer) {
      // Create it after table
      pagContainer = document.createElement('div');
      pagContainer.className = styles.paginationContainer;
      this.domElement.querySelector(`.${styles.content}`)?.appendChild(pagContainer);
    }

    if (totalPages > 1) {
      let pagHtml = `
          <button class="${styles.pageButton}" id="btnPrev" ${this.currentPage === 1 ? 'disabled' : ''}>Prev</button>
          <span>Page ${this.currentPage} of ${totalPages}</span>
          <button class="${styles.pageButton}" id="btnNext" ${this.currentPage === totalPages ? 'disabled' : ''}>Next</button>
        `;
      pagContainer.innerHTML = pagHtml;

      // Bind click events
      const btnPrev = pagContainer.querySelector('#btnPrev');
      const btnNext = pagContainer.querySelector('#btnNext');

      if (btnPrev) {
        btnPrev.addEventListener('click', () => {
          if (this.currentPage > 1) {
            this.currentPage--;
            this.renderTable(category, sub);
          }
        });
      }
      if (btnNext) {
        btnNext.addEventListener('click', () => {
          if (this.currentPage < totalPages) {
            this.currentPage++;
            this.renderTable(category, sub);
          }
        });
      }

      // Ensure visible
      (pagContainer as HTMLElement).style.display = 'flex';
    } else {
      (pagContainer as HTMLElement).style.display = 'none';
      pagContainer.innerHTML = '';
    }
  }

  private async handleRequestAccess(e: Event): Promise<void> {
    const fileId = (e.currentTarget as HTMLElement).dataset.id;
    // Note: Download Count field is optional for now, but if configured we use it.
    if (!this.properties.inputListId || !this.properties.inputFieldFileId || !this.properties.inputFieldEmail) {
      alert('Please configure the Request Access settings in the Web Part properties.');
      return;
    }

    try {
      const userEmail = this.context.pageContext.user.email;
      const listId = this.properties.inputListId;
      const fileIdField = this.properties.inputFieldFileId;
      const emailField = this.properties.inputFieldEmail;
      const countField = this.properties.inputFieldDownloadCount;

      let createdNew = false;
      let newCount = 1;

      if (countField) {
        // Check for existing request
        const existingItem = await this.service.getExistingRequest(
          listId,
          fileIdField,
          fileId,
          emailField,
          userEmail,
          [countField]
        );

        if (existingItem) {
          // Update existing
          const currentCount = existingItem[countField] ? parseInt(existingItem[countField]) : 0;
          newCount = (isNaN(currentCount) ? 0 : currentCount) + 1;

          await this.service.updateRequest(listId, existingItem.Id, {
            [countField]: newCount
          });
        } else {
          // Create new with count 1
          createdNew = true;
          const payload: any = {};
          payload[fileIdField] = fileId;
          payload[emailField] = userEmail;
          payload[countField] = 1; // Start at 1

          await this.service.createRequest(listId, payload);
        }
      } else {
        // Legacy behavior: Always create new if count field not configured
        createdNew = true;
        const payload: any = {};
        payload[fileIdField] = fileId;
        payload[emailField] = userEmail;

        await this.service.createRequest(listId, payload);
      }

      if (countField && !createdNew) {
        alert(`Request updated. Total requests: ${newCount}`);
      } else {
        alert('Access request submitted successfully!');
      }
    } catch (err: any) {
      alert(`Failed to submit request: ${err.message || err}`);
    }
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    this.listsDropdownDisabled = !this.lists || this.lists.length === 0;
    this.requestListsDropdownDisabled = !this.requestLists || this.requestLists.length === 0;
    this.columnsDropdownDisabled = !this.properties.library || !this.columns || this.columns.length === 0;
    this.requestColumnsDropdownDisabled = !this.properties.inputListId || !this.requestColumns || this.requestColumns.length === 0;

    if (this.lists.length === 0) {
      await this.loadLists();
      this.listsDropdownDisabled = false;
      this.context.propertyPane.refresh();

      // Load request lists separately
      await this.loadRequestLists();
      this.requestListsDropdownDisabled = false;
      this.context.propertyPane.refresh();

      // Attempt to load columns if a library is already selected
      if (this.properties.library) {
        await this.loadColumns(this.properties.library);
        this.columnsDropdownDisabled = false;
        this.context.propertyPane.refresh();
      }

      // Attempt to load request columns if selected
      if (this.properties.inputListId) {
        await this.loadRequestColumns(this.properties.inputListId);
        this.requestColumnsDropdownDisabled = false;
        this.context.propertyPane.refresh();
      }
    }
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    if (propertyPath === 'library' && newValue) {
      this.properties.library = newValue;
      this.columnsDropdownDisabled = true;
      this.context.propertyPane.refresh();

      // clear previous column selections
      this.properties.categoryColumn = '';
      this.properties.subCategoryColumn = '';
      this.properties.titleColumn = '';
      this.properties.descriptionColumn = '';

      await this.loadColumns(newValue);
      this.columnsDropdownDisabled = false;
      this.context.propertyPane.refresh();
    }

    if (propertyPath === 'inputListId' && newValue) {
      this.properties.inputListId = newValue;
      this.requestColumnsDropdownDisabled = true;
      this.context.propertyPane.refresh();

      this.properties.inputFieldFileId = '';
      this.properties.inputFieldEmail = '';
      this.properties.inputFieldDownloadCount = '';

      await this.loadRequestColumns(newValue);
      this.requestColumnsDropdownDisabled = false;
      this.context.propertyPane.refresh();
    }
  }

  private async loadLists(): Promise<void> {
    // 101 = Document Library
    const listInfos = await this.service.getLists(101);
    this.lists = listInfos.map(l => ({ key: l.Id, text: l.Title }));
  }

  private async loadRequestLists(): Promise<void> {
    // 100 = Generic List
    const listInfos = await this.service.getLists(100);
    this.requestLists = listInfos.map(l => ({ key: l.Id, text: l.Title }));
  }

  private async loadColumns(listId: string): Promise<void> {
    const fieldInfos = await this.service.getColumns(listId);
    console.log('Loaded Main Library Columns:', fieldInfos);
    // Sort by Title
    fieldInfos.sort((a, b) => (a.Title || a.InternalName).localeCompare(b.Title || b.InternalName));
    this.columns = fieldInfos.map(f => ({ key: f.InternalName, text: f.Title || f.InternalName }));
  }

  private async loadRequestColumns(listId: string): Promise<void> {
    const fieldInfos = await this.service.getColumns(listId);
    console.log('Loaded Request List Columns:', fieldInfos);
    // Sort by Title
    fieldInfos.sort((a, b) => (a.Title || a.InternalName).localeCompare(b.Title || b.InternalName));

    // We might want to filter out ReadOnly for writing, but let's show all just in case
    // to resolve the "missing column" issue.
    // Note: We fetched ReadOnlyField property in service now. We could use it:
    // const editableFields = fieldInfos.filter(f => !f.ReadOnlyField);
    // But let's stick to showing all for maximum flexibility unless user complains of clutter.

    this.requestColumns = fieldInfos.map(f => ({ key: f.InternalName, text: f.Title || f.InternalName }));
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        groups: [{
          groupName: 'Configuration',
          groupFields: [
            PropertyPaneDropdown('library', {
              label: 'Document Library Name',
              options: this.lists,
              disabled: this.listsDropdownDisabled
            }),
            PropertyPaneDropdown('categoryColumn', {
              label: 'Category Column',
              options: this.columns,
              disabled: this.columnsDropdownDisabled
            }),
            PropertyPaneDropdown('subCategoryColumn', {
              label: 'Sub Category Column',
              options: this.columns,
              disabled: this.columnsDropdownDisabled
            }),
            PropertyPaneDropdown('titleColumn', {
              label: 'Title Field',
              options: this.columns,
              disabled: this.columnsDropdownDisabled
            }),
            PropertyPaneDropdown('descriptionColumn', {
              label: 'Description Field',
              options: this.columns,
              disabled: this.columnsDropdownDisabled
            }),
            PropertyPaneSlider('pageSize', {
              label: 'Max Rows per Page',
              min: 5,
              max: 100,
              step: 5,
              value: this.properties.pageSize || 10
            })
          ]
        },
        {
          groupName: 'Request Access Configuration',
          groupFields: [
            PropertyPaneDropdown('inputListId', {
              label: 'Requests List (Generic Lists)',
              options: this.requestLists,
              disabled: this.requestListsDropdownDisabled
            }),
            PropertyPaneDropdown('inputFieldFileId', {
              label: 'Column for File ID',
              options: this.requestColumns,
              disabled: this.requestColumnsDropdownDisabled
            }),
            PropertyPaneDropdown('inputFieldEmail', {
              label: 'Column for User Email',
              options: this.requestColumns,
              disabled: this.requestColumnsDropdownDisabled
            }),
            PropertyPaneDropdown('inputFieldDownloadCount', {
              label: 'Column for Download Count',
              options: this.requestColumns,
              disabled: this.requestColumnsDropdownDisabled
            })
          ]
        }]
      }]
    };
  }
}