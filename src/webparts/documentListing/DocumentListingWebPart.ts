import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';

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

  // New State for Search & Sort
  private searchQuery: string = '';
  private sortConfig: { column: string; direction: 'asc' | 'desc' } = { column: 'Modified', direction: 'desc' };

  private lists: IPropertyPaneDropdownOption[] = [];
  private requestLists: IPropertyPaneDropdownOption[] = []; // For generic lists (requests)
  private columns: IPropertyPaneDropdownOption[] = [];
  private requestColumns: IPropertyPaneDropdownOption[] = []; // Columns for the separate request list
  private listsDropdownDisabled: boolean = true;
  private requestListsDropdownDisabled: boolean = true;
  // Filtered dropdown options
  private choiceColumns: IPropertyPaneDropdownOption[] = [];
  private textColumns: IPropertyPaneDropdownOption[] = [];
  private requestSimpleColumns: IPropertyPaneDropdownOption[] = [];

  private columnsDropdownDisabled: boolean = true;
  private requestColumnsDropdownDisabled: boolean = true;

  public async onInit(): Promise<void> {
    await super.onInit();

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
    // Trigger render
    this.render();
  }

  public render(): void {
    this._renderAsync().catch(err => console.error(err));
  }

  private async _renderAsync(): Promise<void> {
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
      ].filter(Boolean); // Filter out empty

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
          if (!categories.includes(i.Category) && i.Category) {
            categories.push(i.Category);
          }
        });
        categories.sort((a, b) => a.localeCompare(b));
      }

      this.domElement.innerHTML = `
        ${this.properties.webPartTitle ? `<div class="${styles.webPartHeader}">${this.properties.webPartTitle}</div>` : ''}
        
        <!-- Toast Notification Container -->
        <div id="toast" class="${styles.toast}">
          <i class="ms-Icon ms-Icon--Completed" aria-hidden="true"></i>
          <span id="toastMsg"></span>
        </div>

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

            <!-- Search Bar -->
            <div class="${styles.searchContainer}">
              <i class="ms-Icon ms-Icon--Search ${styles.searchIcon}" aria-hidden="true"></i>
              <input type="text" id="searchInput" class="${styles.searchInput}" placeholder="Search documents..." value="${this.searchQuery}">
            </div>

            <div class="${styles.tableWrapper}">
              <div class="${styles.tableContainer}">
                <div class="${styles.tableHeader}" id="tableHeader">
                   <!-- Headers injected by renderTable to support dynamic sorting visual -->
                </div>
                <div id="docRows"></div>
              </div>
            </div>
            
            <div class="${styles.paginationContainer}"></div>
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
          const cat = firstCatEl.getAttribute('data-category');
          if (cat) {
            this.loadSubCategories(cat).catch(console.error);
          }
        }
      }

    } catch (error) {
      const errorMessage = this.getErrorMessage(error);
      this.domElement.innerHTML = `<p>Error loading documents: ${errorMessage}</p>`;
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

          const category = (e.currentTarget as HTMLElement).dataset.category;
          if (category) {
            this.loadSubCategories(category).catch(console.error);
          }
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
        if (i.Category === category && !subs.includes(i.SubCategory) && i.SubCategory) {
          subs.push(i.SubCategory);
        }
      });
      subs.sort((a, b) => a.localeCompare(b));
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

    const subTabsContainer = this.domElement.querySelector('#subTabs');
    if (subTabsContainer) {
      subTabsContainer.innerHTML = tabs;
    }

    this.bindSubCategoryEvents(category);

    // Bind Search Event one time for the lifecycle of this render
    this.bindSearchEvent(category);

    // Auto-select first subcategory (already marked active in HTML above)
    if (subs.length > 0) {
      this.currentPage = 1;
      this.renderTable(category, subs[0]);
    } else {
      this.currentPage = 1;
      this.renderTable(category, '');
    }
  }

  private bindSearchEvent(category: string): void {
    const searchInput = this.domElement.querySelector('#searchInput') as HTMLInputElement;
    if (searchInput) {
      // Remove old listeners by cloning or just assume re-render clears strictly? 
      // Actually _renderAsync overwrites innerHTML so listeners are gone. 
      // But loadSubCategories DOES NOT overwrite main container, so search input PERSISTS.
      // We must be careful not to double bind.
      // Best approach: cloneNode to wipe listeners or just use a flag?
      // Simple approach: removeEventListener before adding? Hard without ref to exact function.
      // Or, since we only call loadSubCategories when Category changes, we can just assume we want to update the "current" category context for search?
      // ACTUALLY: The search input is outside the subTabs container. It was rendered in _renderAsync.
      // So loadSubCategories is just manipulating the table. 
      // We need to make sure the listener knows the CURRENT active subcategory.
      // Better: Store currentCategory and currentSubCategory in state so the event handler can read them.

      // Let's attach a fresh listener. To avoid duplicates, we can recreate the element or use 'oninput' property.
      searchInput.oninput = (e) => {
        this.searchQuery = (e.target as HTMLInputElement).value;
        this.currentPage = 1; // Reset page on search

        // Find active sub tab
        const activeSub = this.domElement.querySelector(`.${styles.subTab}.${styles.active}`) as HTMLElement;
        const currentSub = activeSub ? activeSub.dataset.sub || '' : '';

        this.renderTable(category, currentSub);
      };
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
          const sub = el.dataset.sub;
          if (sub) {
            this.renderTable(category, sub);
          }
        });
      });
  }

  private renderTable(category: string, sub: string): void {
    // 1. Filter by Category & SubCategory
    let items = this.items.filter(i => i.Category === category && i.SubCategory === sub);

    // 2. Filter by Search Query
    if (this.searchQuery) {
      const q = this.searchQuery.toLowerCase();
      items = items.filter(i => {
        const title = i.Title || '';
        const desc = i.Description || '';
        // Handle object fields if necessary, currently assumed strings or addressed below
        // Simplified check:
        return String(title).toLowerCase().indexOf(q) !== -1 ||
          String(desc).toLowerCase().indexOf(q) !== -1;
      });
    }

    // 3. Sort
    items.sort((a, b) => {
      const col = this.sortConfig.column;
      const dir = this.sortConfig.direction === 'asc' ? 1 : -1;

      let valA: any = a[col as keyof IDocumentItem];
      let valB: any = b[col as keyof IDocumentItem];

      // Handle simple object projections if needed (like Title/Description logic)
      // This is basic sorting. For robust sorting on complex objects, we might need more logic.
      if (typeof valA === 'string') valA = valA.toLowerCase();
      if (typeof valB === 'string') valB = valB.toLowerCase();

      if (valA < valB) return -1 * dir;
      if (valA > valB) return 1 * dir;
      return 0;
    });

    // Pagination Logic
    const pageSize = this.properties.pageSize || 10;
    const totalPages = Math.ceil(items.length / pageSize);

    // Ensure current page is valid
    if (this.currentPage > totalPages) this.currentPage = 1;
    if (this.currentPage < 1) this.currentPage = 1;

    const startIndex = (this.currentPage - 1) * pageSize;
    const endIndex = startIndex + pageSize;
    const pagedItems = items.slice(startIndex, endIndex);

    // Render Headers with Sort Indicators
    const renderHeader = (colKey: string, label: string, className: string): string => {
      const isSorted = this.sortConfig.column === colKey;

      const iconClass = isSorted
        ? (this.sortConfig.direction === 'asc' ? 'ms-Icon--SortUp' : 'ms-Icon--SortDown')
        : '';

      return `
        <div class="${styles.headerCell} ${className} ${styles.sortableHeader}" data-col="${colKey}">
          ${label}
          ${isSorted ? `<i class="ms-Icon ${iconClass} ${styles.sortIcon}" aria-hidden="true"></i>` : ''}
        </div>
      `;
    };

    const headerHtml = `
      ${renderHeader(this.properties.titleColumn || 'Title', 'Title', styles.colTitle)}
      ${renderHeader(this.properties.descriptionColumn || 'Description', 'Description', styles.colDesc)}
      ${renderHeader('Modified', 'Date', styles.colDate)}
      <div class="${styles.headerCell} ${styles.colAction}">Request Access</div>
    `;

    const headerContainer = this.domElement.querySelector('#tableHeader');
    if (headerContainer) {
      headerContainer.innerHTML = headerHtml;
      // Bind Sort Events
      headerContainer.querySelectorAll(`.${styles.sortableHeader}`).forEach(h => {
        h.addEventListener('click', (e) => {
          const col = (e.currentTarget as HTMLElement).dataset.col;
          if (col) {
            this.handleSort(col, category, sub);
          }
        });
      });
    }

    const rows = pagedItems.map(i => {
      const titleVal = this.properties.titleColumn ? i[this.properties.titleColumn] : i.Title;
      let displayTitle = '';

      if (typeof titleVal === 'object' && titleVal !== null) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        displayTitle = (titleVal as any).Title || (titleVal as any).title || '';
      } else {
        const primitiveVal = titleVal as string | number | boolean | null | undefined;
        displayTitle = (primitiveVal === null || primitiveVal === undefined) ? '' : String(primitiveVal);
      }

      const descVal = this.properties.descriptionColumn ? i[this.properties.descriptionColumn] : (i.Description || '');
      let displayDesc = '';

      if (typeof descVal === 'object' && descVal !== null) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const obj = descVal as any;
        displayDesc = obj.Description || obj.Title || obj.title || '';
      } else {
        const primitiveDesc = descVal as string | number | boolean | null | undefined;
        displayDesc = (primitiveDesc === null || primitiveDesc === undefined) ? '' : String(primitiveDesc);
      }

      return `
        <div class="${styles.tableRow}">
          <div class="${styles.tableCell} ${styles.colTitle}">${displayTitle}</div>
          <div class="${styles.tableCell} ${styles.colDesc}">${displayDesc}</div>
          <div class="${styles.tableCell} ${styles.colDate}">${i.Modified ? new Date(i.Modified).toLocaleDateString() : ''}</div>
          <div class="${styles.tableCell} ${styles.colAction}">
             <button class="${styles.mailIcon} request-access-btn" 
                     title="Request Access"
                     data-id="${i.Id}">
               <i class="ms-Icon ms-Icon--Mail" aria-hidden="true"></i>
             </button>
          </div>
        </div>
      `;
    }).join('');

    const docRowsContainer = this.domElement.querySelector('#docRows');
    if (docRowsContainer) {
      docRowsContainer.innerHTML = rows || `<div class="${styles.tableRow}"><div class="${styles.tableCell}">No documents found.</div></div>`;

      // Re-bind Request Access Buttons
      this.domElement.querySelectorAll('.request-access-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
          this.handleRequestAccess(e).catch(err => console.error(err));
        });
      });
    }

    // Render Pagination Controls
    let pagContainer = this.domElement.querySelector(`.${styles.paginationContainer}`);
    if (!pagContainer) {
      pagContainer = document.createElement('div');
      pagContainer.className = styles.paginationContainer;
      this.domElement.querySelector(`.${styles.content}`)?.appendChild(pagContainer);
    }

    if (totalPages > 1) {
      const pagHtml = `
          <button class="${styles.pageButton}" id="btnPrev" ${this.currentPage === 1 ? 'disabled' : ''}>Prev</button>
          <span>Page ${this.currentPage} of ${totalPages}</span>
          <button class="${styles.pageButton}" id="btnNext" ${this.currentPage === totalPages ? 'disabled' : ''}>Next</button>
        `;
      pagContainer.innerHTML = pagHtml;

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

      (pagContainer as HTMLElement).style.display = 'flex';
    } else {
      (pagContainer as HTMLElement).style.display = 'none';
      pagContainer.innerHTML = '';
    }
  }

  private handleSort(column: string, category: string, sub: string): void {
    if (this.sortConfig.column === column) {
      // Toggle direction
      this.sortConfig.direction = this.sortConfig.direction === 'asc' ? 'desc' : 'asc';
    } else {
      // New column, default desc for Date, asc for others? Defaulting to asc for now.
      this.sortConfig = { column, direction: 'asc' };
    }
    this.renderTable(category, sub);
  }

  private showToast(message: string): void {
    const toast = this.domElement.querySelector('#toast');
    const toastMsg = this.domElement.querySelector('#toastMsg');

    if (toast && toastMsg) {
      toastMsg.textContent = message;
      toast.classList.add(styles.toastVisible);
      // WAIT: styles.show only works if .show is defined in scss AND imported. 
      // In my scss update I did `.toast { ... &.show { ... } }`. 
      // So `styles.show` might NOT be generated if it's nested or depending on css modules config.
      // Usually standard CSS modules: yes, nested classes are exported. 
      // However, to be safe, I often prefer toggling a global class or ensure it's top level.
      // BUT for now, let's assume `styles.show` is unavailable if it's nested (Sass nesting doesn't expose nested class names as top-level exports in all loaders).
      // Actually, standard css-modules w/ sass: `.toast.show` -> the class is hashed. You can't just add a string 'show'.
      // You need to add `styles.show` IF `show` was top level.
      // FIX: The SCSS had `&.show`. This means the class is `.toast.show` (hashed together? No).
      // CSS Modules: `.toast` is `document_toast_hash`. `.show` inside is likely NOT exported if nested with `&`.
      // Better approach: Define `.toastShow` as a persistent helper or just rely on global class if convenient.
      // CORRECT FIX: In SCSS I should have defined `.toastShow` separately or use `:global(.show)`.
      // Let's assume I need to fix the SCSS or usage. 
      // PROPOSED LOGIC CHANGE: I'll use `styles.toastShow` and update SCSS in next step? 
      // OR simpler: Just manually add style `opacity: 1; transform: ...` in this method.
      // No, let's try to use `styles.toast` and dynamically style it.
      // actually, just adding 'show' string class works IF I used `:global(.show)` in SCSS.
      // I didn't. 
      // Re-reading SCSS: `&.show`. This parses to `.toast.show`.
      // The CSS file will have `.toast_hash.show_hash`? No, SASS `&` usually combines selector.
      // If it's `&.show`, it expects the element to have BOTH classes. 
      // BUT `show` isn't in `styles` object if it's not top level.
      // Workaround: I'll manually set style.opacity = '1'. 
    }

    // Let's do manual style manipulation for safety in this strict environment without checking `styles` object at runtime.
    if (toast) {
      const t = toast as HTMLElement;
      t.style.opacity = '1';
      t.style.transform = 'translateY(0)';
      t.style.pointerEvents = 'auto';

      if (toastMsg) toastMsg.textContent = message;

      setTimeout(() => {
        t.style.opacity = '0';
        t.style.transform = 'translateY(-10px)';
        t.style.pointerEvents = 'none';
      }, 3000);
    }
  }

  private async handleRequestAccess(e: Event): Promise<void> {
    const fileId = (e.currentTarget as HTMLElement).dataset.id;
    // Note: Download Count field is optional for now, but if configured we use it.
    // Note: Download Count field is optional for now, but if configured we use it.
    if (!this.properties.inputListId || !this.properties.inputFieldFileId || !this.properties.inputFieldEmail) {
      this.showToast('Please configure the Request Access settings in the Web Part properties.');
      return;
    }

    await this.processAccessRequest(fileId, this.context.pageContext.user.email);
  }

  private async _handleCountRequest(
    listId: string,
    fileIdField: string,
    fileId: string,
    emailField: string,
    userEmail: string,
    countField: string
  ): Promise<{ createdNew: boolean; newCount: number }> {
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
      const existingVal = existingItem[countField] as string | undefined;
      const currentCount = existingVal ? Number.parseInt(existingVal) : 0;
      const newCount = (Number.isNaN(currentCount) ? 0 : currentCount) + 1;

      await this.service.updateRequest(listId, existingItem.Id as number, {
        [countField]: newCount
      });
      return { createdNew: false, newCount };
    } else {
      // Create new with count 1
      const payload: Record<string, unknown> = {};
      payload[fileIdField] = fileId;
      payload[emailField] = userEmail;
      payload[countField] = 1; // Start at 1

      await this.service.createRequest(listId, payload);
      return { createdNew: true, newCount: 1 };
    }
  }

  private async processAccessRequest(fileId: string | undefined, userEmail: string): Promise<void> {
    if (!fileId) {
      this.showToast('File ID not found for request.');
      return;
    }

    try {
      const listId = this.properties.inputListId;
      const fileIdField = this.properties.inputFieldFileId;
      const emailField = this.properties.inputFieldEmail;
      const countField = this.properties.inputFieldDownloadCount;

      let createdNew = true;
      let newCount = 1;

      if (countField) {
        const result = await this._handleCountRequest(
          listId,
          fileIdField,
          fileId,
          emailField,
          userEmail,
          countField
        );
        createdNew = result.createdNew;
        newCount = result.newCount;
      } else {
        // Legacy behavior: Always create new if count field not configured
        const payload: Record<string, unknown> = {};
        payload[fileIdField] = fileId;
        payload[emailField] = userEmail;
        await this.service.createRequest(listId, payload);
      }

      if (countField && !createdNew) {
        this.showToast(`Request updated. Total requests: ${newCount}`);
      } else {
        this.showToast('Access request submitted successfully!');
      }
    } catch (err: unknown) {
      console.error('Error submitting request:', err);
      const errorMessage = this.getErrorMessage(err);
      this.showToast(`Failed to submit request: ${errorMessage}`);
    }
  }


  private getErrorMessage(err: unknown): string {
    if (err instanceof Error) {
      return err.message;
    } else if (typeof err === 'string') {
      return err;
    } else if (typeof err === 'object' && err !== null) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return (err as any).message || JSON.stringify(err);
    } else {
      // Primitive types or unknown that is not an object/null
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return String(err as any);
    }
  }

  protected onPropertyPaneConfigurationStart(): void {
    this._onPropertyPaneConfigurationStartAsync().catch(err => console.error(err));
  }

  private async _onPropertyPaneConfigurationStartAsync(): Promise<void> {
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

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    this._onPropertyPaneFieldChangedAsync(propertyPath, oldValue, newValue).catch(err => console.error(err));
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private async _onPropertyPaneFieldChangedAsync(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
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

    // All columns (fallback)
    this.columns = fieldInfos.map(f => ({
      key: f.InternalName,
      text: f.Title || f.InternalName
    }));

    // Sort
    this.columns.sort((a, b) => a.text.localeCompare(b.text));

    // Filter: Choice
    this.choiceColumns = fieldInfos
      .filter(f => f.TypeAsString === 'Choice' || f.TypeAsString === 'MultiChoice')
      .map(f => ({ key: f.InternalName, text: f.Title || f.InternalName }));
    this.choiceColumns.sort((a, b) => a.text.localeCompare(b.text));

    // Filter: Text or Note (MultiLine)
    this.textColumns = fieldInfos
      .filter(f => f.TypeAsString === 'Text' || f.TypeAsString === 'Note')
      .map(f => ({ key: f.InternalName, text: f.Title || f.InternalName }));
    this.textColumns.sort((a, b) => a.text.localeCompare(b.text));
  }

  private async loadRequestColumns(listId: string): Promise<void> {
    const fieldInfos = await this.service.getColumns(listId);
    console.log('Loaded Request List Columns:', fieldInfos);

    // All columns (fallback)
    this.requestColumns = fieldInfos.map(f => ({
      key: f.InternalName,
      text: f.Title || f.InternalName
    }));

    this.requestColumns.sort((a, b) => a.text.localeCompare(b.text));

    // Filter: Text or Number (simple types for ID/Email/Count)
    this.requestSimpleColumns = fieldInfos
      .filter(f => f.TypeAsString === 'Text' || f.TypeAsString === 'Number')
      .map(f => ({ key: f.InternalName, text: f.Title || f.InternalName }));
    this.requestSimpleColumns.sort((a, b) => a.text.localeCompare(b.text));
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        groups: [{
          groupName: 'Configuration',
          groupFields: [
            PropertyPaneTextField('webPartTitle', {
              label: 'Web Part Title (Header)'
            }),
            PropertyPaneDropdown('library', {
              label: 'Document Library Name',
              options: this.lists,
              disabled: this.listsDropdownDisabled
            }),
            PropertyPaneDropdown('categoryColumn', {
              label: 'Category Column',
              options: this.choiceColumns, // Filtered
              disabled: this.columnsDropdownDisabled
            }),
            PropertyPaneDropdown('subCategoryColumn', {
              label: 'Sub Category Column',
              options: this.choiceColumns, // Filtered
              disabled: this.columnsDropdownDisabled
            }),
            PropertyPaneDropdown('titleColumn', {
              label: 'Title Field (Optional override)',
              options: this.textColumns, // Filtered
              disabled: this.columnsDropdownDisabled
            }),
            PropertyPaneDropdown('descriptionColumn', {
              label: 'Description Field',
              options: this.textColumns, // Filtered
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
          groupName: "Request Access Configuration",
          groupFields: [
            PropertyPaneDropdown('inputListId', {
              label: 'Requests List',
              options: this.requestLists,
              disabled: this.requestListsDropdownDisabled
            }),
            PropertyPaneDropdown('inputFieldFileId', {
              label: 'Column for File ID',
              options: this.requestSimpleColumns, // Filtered
              disabled: this.requestColumnsDropdownDisabled
            }),
            PropertyPaneDropdown('inputFieldEmail', {
              label: 'Column for User Email',
              options: this.requestSimpleColumns, // Filtered
              disabled: this.requestColumnsDropdownDisabled
            }),
            PropertyPaneDropdown('inputFieldDownloadCount', {
              label: 'Column for Download Count (Optional)',
              options: this.requestSimpleColumns, // Filtered
              disabled: this.requestColumnsDropdownDisabled
            })
          ]
        }
        ]
      }
      ]
    };
  }
}