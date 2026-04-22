import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import "@pnp/sp/fields";
import "@pnp/sp/search";
import { SearchResults } from "@pnp/sp/search";
import { IToDoItem, ILookupOption, IAttachment } from "../models/IToDoItem";
import { ITrainingInductionItem } from "../models/ITrainingInductionItem";



export class SPService {
    private _sp: SPFI;
    private _fieldMaps: Map<string, Map<string, string>> = new Map();
    private _fieldTypeMaps: Map<string, Map<string, string>> = new Map();
    private _listTitles: Map<string, string> = new Map();
    private _listIds: Map<string, string> = new Map();
    private _listUrls: Map<string, string> = new Map();
    private _isInitialized: Map<string, boolean> = new Map();
    private _listFieldInternalNames: Map<string, string[]> = new Map();
    private _lookupDisplayFields: Map<string, Map<string, string>> = new Map();
    private _lastError: string = "";

    constructor(private context: WebPartContext) {
        this._sp = spfi().using(SPFx(context));
    }

    public async init(urlName: string = "ToDo"): Promise<void> {
        if (this._isInitialized.get(urlName)) return;
        try {
            const lists = await this._sp.web.lists.select("Title", "RootFolder/Name").expand("RootFolder")();
            const match = lists.find(l =>
                l.Title.toLowerCase() === urlName.toLowerCase() ||
                l.RootFolder.Name.toLowerCase() === urlName.toLowerCase() ||
                l.Title.replace(/[\s&]+/g, '').toLowerCase() === urlName.toLowerCase() ||
                l.Title.replace(/[\s&]+/g, '').toLowerCase() === urlName.toLowerCase().replace(/s$/, '')
            );
            const listTitle = match ? match.Title : urlName;
            console.log(`[SPService] Matched list "${urlName}" to SharePoint title: "${listTitle}"`);
            this._listTitles.set(urlName, listTitle);

            const listData = await this._sp.web.lists.getByTitle(listTitle).select("Id", "Title", "RootFolder/ServerRelativeUrl").expand("RootFolder")();
            const listId = listData.Id;
            let listUrl = listData.RootFolder.ServerRelativeUrl;

            // Normalize URL for Search: ensure absolute path and no trailing slash for the prefix
            const webUrl = this.context.pageContext.web.absoluteUrl;
            const siteUrl = webUrl.substring(0, webUrl.indexOf(this.context.pageContext.web.serverRelativeUrl));
            listUrl = siteUrl + listUrl;
            if (listUrl.endsWith('/')) listUrl = listUrl.slice(0, -1);

            console.log(`[SPService] List Identified: "${listTitle}" (Path: ${listUrl}, ID: ${listId})`);
            
            const fields = await this._sp.web.lists.getByTitle(listTitle).fields
                .select("Title", "InternalName", "TypeAsString", "LookupField")();

            const fieldMap = new Map<string, string>();
            const fieldTypeMap = new Map<string, string>();
            const displayFieldMap = new Map<string, string>();
            const internalNames: string[] = [];

            fields.forEach(f => {
                const title = (f.Title || "").trim();
                const internal = (f.InternalName || "").trim();
                const type = (f.TypeAsString || "").trim();

                fieldMap.set(title, internal);
                fieldTypeMap.set(internal, type);
                internalNames.push(internal);

                if (type === "Lookup" || type === "LookupMulti") {
                    if ((f as any).LookupField) {
                        displayFieldMap.set(internal, (f as any).LookupField);
                    }
                }
            });

            this._fieldMaps.set(urlName, fieldMap);
            this._fieldTypeMaps.set(urlName, fieldTypeMap);
            this._lookupDisplayFields.set(urlName, displayFieldMap);
            this._listFieldInternalNames.set(urlName, internalNames);
            this._listIds.set(urlName, listId);
            this._listUrls.set(urlName, listUrl);
            this._isInitialized.set(urlName, true);
        } catch (error) {
            console.error(`[SPService] Initialization Failed for ${urlName}:`, error);
            this._listTitles.set(urlName, urlName);
            this._isInitialized.set(urlName, true);
        }
    }

    public getInternalName(listName: string, title: string, fallback: string): string {
        const fieldMap = this._fieldMaps.get(listName);
        if (!fieldMap) return fallback;

        if (fieldMap.has(title)) return fieldMap.get(title)!;
        const keys = Array.from(fieldMap.keys());
        const match = keys.find(k =>
            k.toLowerCase() === title.toLowerCase() ||
            k.replace(/\s/g, '').toLowerCase() === title.toLowerCase()
        );
        if (match) return fieldMap.get(match)!;
        return fallback;
    }

    public async checkMissingFields(listName: string = "ToDo"): Promise<{ title: string; internalName: string; exists: boolean }[]> {
        await this.init(listName);
        const required = [
            { title: "Subject", internal: "Subject" },
            { title: "Category", internal: "Category" },
            { title: "Status", internal: "Status" },
            { title: "Priority", internal: "Priority" },
            { title: "Due Date", internal: "DueDate" },
            { title: "Task Owner", internal: "TaskOwner" },
        ];
        const fieldMap = this._fieldMaps.get(listName);
        if (!fieldMap) return [];
        const fieldInternalNames = Array.from(fieldMap.values());
        const fieldTitles = Array.from(fieldMap.keys());
        return required.map(r => ({
            title: r.title,
            internalName: r.internal,
            exists: fieldTitles.some(t => t.toLowerCase() === r.title.toLowerCase()) ||
                fieldInternalNames.some(i => i.toLowerCase() === r.internal.toLowerCase())
        }));
    }

    private _applyClientSideFilter(items: IToDoItem[], query: string): IToDoItem[] {
        if (!query) return items;
        const q = query.toLowerCase().trim();
        const isNum = !isNaN(Number(q));
        const num = Number(q);

        return items.filter(item => {
            // 1. ID - exact match
            if (isNum && item.Id === num) return true;

            // 2. Text Search - case insensitive includes
            const check = (val: any): boolean => {
                if (val === null || val === undefined) return false;
                return val.toString().toLowerCase().indexOf(q) >= 0;
            };

            return (
                check(item.Title) ||
                check(item.Regarding) ||
                check(item.Status?.Title) ||
                check(item.Status?.Name) ||
                check(item.Category?.Title) ||
                check(item.Category?.Name) ||
                check(item.Priority?.Title) ||
                check(item.Priority?.Name) ||
                check(item.Classification?.Title) ||
                check(item.Classification?.Name) ||
                check(item.TaskOwner?.Title) ||
                check(item.AssigneeInternal?.Title) ||
                check(item.AssigneeExternal?.Title)
            );
        });
    }

    private _buildKQLQuery(searchQuery: string): string {
        if (!searchQuery) return "";
        // REMOVAL: No wildcards (*) as per requirements
        const terms = searchQuery.trim().split(/\s+/).filter(t => t.length > 0);
        if (terms.length === 0) return "";

        // Use clean terms joined by AND
        const cleanTerms = terms.map(t => t.replace(/["\\]/g, ""));
        return `"${cleanTerms.join(' ')}"`;
    }

    public async getToDoTotalCount(searchQuery?: string, extraFilter?: string): Promise<number> {
        await this.init("ToDo");
        const listTitle = this._listTitles.get("ToDo")!;
        const listUrl = this._listUrls.get("ToDo")!;
        
        try {
            if (searchQuery && !extraFilter) {
                // If we have a search query AND no extra complex filters, use Search API for scalability (50,000+ items)
                const kql = `${this._buildKQLQuery(searchQuery)} (path:"${listUrl}" OR path:"${listUrl}/*")`;
                const results: SearchResults = await this._sp.search({
                    Querytext: kql,
                    RowLimit: 0,
                    SelectProperties: ["ListItemID"]
                });
                return results.TotalRows || 0;
            }

            // Fallback to REST filter for complex filtering or browse-mode counting
            const names = {
                Subject: this.getInternalName("ToDo", "Subject", "Title"),
                Regarding: this.getInternalName("ToDo", "Regarding", "Regarding"),
                TaskOwner: this.getInternalName("ToDo", "Task Owner", "TaskOwner"),
            };

            let query = this._sp.web.lists.getByTitle(listTitle).items.select("Id").top(5000);
            const filters: string[] = [];
            if (searchQuery) filters.push(`(substringof('${searchQuery}', ${names.Subject}) or substringof('${searchQuery}', ${names.Regarding}) or substringof('${searchQuery}', ${names.TaskOwner}/Title))`);
            if (extraFilter) filters.push(extraFilter);

            if (filters.length > 0) query = query.filter(filters.join(" and "));

            const items = await query();
            return items.length;
        } catch (e) {
            console.warn("[SPService] getToDoTotalCount fallback:", e);
            return 0;
        }
    }

    public async getToDoItemsPaged(
        page: number,
        pageSize: number,
        searchQuery?: string,
        sortField: string = "Id",
        isAscending: boolean = true,
        extraFilter?: string
    ): Promise<IToDoItem[]> {
        await this.init("ToDo");
        const listTitle = this._listTitles.get("ToDo")!;
        const fieldTypeMap = this._fieldTypeMaps.get("ToDo")!;
        const fieldInternalNames = this._listFieldInternalNames.get("ToDo") || [];

        const names = {
            Subject: this.getInternalName("ToDo", "Subject", "Title"),
            TaskOwner: this.getInternalName("ToDo", "Task Owner", "TaskOwner"),
            AssigneeInternal: this.getInternalName("ToDo", "Assigne Internal", "AssigneInternal"),
            AssigneeExternal: this.getInternalName("ToDo", "Assigne External", "AssigneExternal"),
            Status: this.getInternalName("ToDo", "Status", "Status"),
            Category: this.getInternalName("ToDo", "Category", "Category"),
            Classification: this.getInternalName("ToDo", "Classification", "Classification"),
            Priority: this.getInternalName("ToDo", "Priority", "Priority"),
            CompletedPercent: this.getInternalName("ToDo", "Completed %", "CompletedPercent"),
            StartDate: this.getInternalName("ToDo", "Start Date", "StartDate"),
            CompletionDate: this.getInternalName("ToDo", "Completion Date", "CompletionDate"),
            CreatedByUser: this.getInternalName("ToDo", "Created By User", "CreatedByUser"),
            UpdatedByUser: this.getInternalName("ToDo", "Updated By User", "UpdatedByUser"),
            CreatedOn: this.getInternalName("ToDo", "Created On", "CreatedOn"),
            UpdatedOn: this.getInternalName("ToDo", "Updated On", "UpdatedOn"),
            Description: this.getInternalName("ToDo", "Description", "Description"),
            Regarding: this.getInternalName("ToDo", "Regarding", "Regarding"),
            DueDate: this.getInternalName("ToDo", "Due Date", "DueDate"),
            Resolution: this.getInternalName("ToDo", "Resolution", "Resolution"),
            EmailNotifications: this.getInternalName("ToDo", "Email Notification", "EmailNotifications"),
            TrainingInduction: this.getInternalName("ToDo", "Training & Induction", "TrainingInduction"),
        };

        const selects = ["*", "Id", "Title", "Author/Title", "Author/EMail", "Editor/Title", "Editor/EMail", "Created", "Modified", "Attachments"];
        const expands = ["Author", "Editor", "AttachmentFiles"];

        const safelyAddSelect = (internalName: string): void => {
            if (fieldInternalNames.indexOf(internalName) < 0) return;
            const fieldType = fieldTypeMap.get(internalName) || "";
            if (fieldType === "User" || fieldType === "UserMulti") {
                selects.push(`${internalName}/Id`, `${internalName}/Title`, `${internalName}/EMail`);
                expands.push(internalName);
            } else if (fieldType === "Lookup" || fieldType === "LookupMulti") {
                selects.push(`${internalName}/Id`, `${internalName}/Title`);
                expands.push(internalName);
            } else {
                selects.push(internalName);
            }
        };

        Object.values(names).forEach(safelyAddSelect);

        const skipCount = (page - 1) * pageSize;
        let realSortField = sortField;
        if (sortField === "TaskOwner") realSortField = `${names.TaskOwner}/Title`;
        else if (sortField === "Subject") realSortField = names.Subject;
        else if (sortField === "Regarding") realSortField = names.Regarding;

        try {
            let rawItems: any[] = [];

            // STRATEGY: Hybrid Search for 100% scalability
            if (searchQuery) {
                // Fetch a large-enough subset for client-side search (LVT safe due to indexed sort + no filter)
                let query = this._sp.web.lists.getByTitle(listTitle).items
                    .select(...selects)
                    .expand(...expands)
                    .orderBy(realSortField, isAscending)
                    .top(300); 

                rawItems = await query();
                rawItems = this._applyClientSideFilter(rawItems, searchQuery);
                rawItems = rawItems.slice(skipCount, skipCount + pageSize);
            } else {
                let query = this._sp.web.lists.getByTitle(listTitle).items
                    .select(...selects)
                    .expand(...expands)
                    .top(pageSize)
                    .orderBy(realSortField, isAscending);

                if (extraFilter) query = query.filter(extraFilter);

                rawItems = await (skipCount > 0 ? query.skip(skipCount)() : query());
            }

            return rawItems.map(item => {
                const getLookup = (n: string) => item[n] ? { Id: item[n].Id || 0, Title: item[n].Title || item[n].Name || "", Name: item[n].Name } : undefined;
                const getPerson = (n: string) => item[n] ? { Id: item[n].Id || 0, Title: item[n].Title || "", EMail: item[n].EMail || "" } : undefined;
                return {
                    ...item,
                    Id: item.Id,
                    Title: item[names.Subject] || item.Title || "",
                    Description: item[names.Description],
                    Status: getLookup(names.Status),
                    Category: getLookup(names.Category),
                    Classification: getLookup(names.Classification),
                    Priority: getLookup(names.Priority),
                    TaskOwner: getPerson(names.TaskOwner),
                    AssigneeInternal: getPerson(names.AssigneeInternal),
                    AssigneeExternal: getPerson(names.AssigneeExternal),
                    Regarding: item[names.Regarding],
                    TrainingInduction: getLookup(names.TrainingInduction),
                    DueDate: item[names.DueDate],
                    StartDate: item[names.StartDate],
                    CompletionDate: item[names.CompletionDate],
                    CompletedPercent: item[names.CompletedPercent],
                    EmailNotifications: item[names.EmailNotifications],
                    Author: getPerson(names.CreatedByUser) || item.Author,
                    Editor: getPerson(names.UpdatedByUser) || item.Editor,
                    Created: item[names.CreatedOn] || item.Created,
                    Modified: item[names.UpdatedOn] || item.Modified,
                    Resolution: item[names.Resolution],
                    AttachmentFiles: item.AttachmentFiles || [],
                };
            });
        } catch (error) {
            console.error(`[SPService] Paged query failed:`, error);
            return [];
        }
    }

    public async getToDoItemsFiltered(
        searchQuery?: string,
        sortField: string = "Id",
        isAscending: boolean = true
    ): Promise<IToDoItem[]> {
        await this.init("ToDo");
        const listTitle = this._listTitles.get("ToDo")!;

        const names = {
            Subject: this.getInternalName("ToDo", "Subject", "Title"),
            Status: this.getInternalName("ToDo", "Status", "Status"),
            Category: this.getInternalName("ToDo", "Category", "Category"),
            Priority: this.getInternalName("ToDo", "Priority", "Priority"),
            Classification: this.getInternalName("ToDo", "Classification", "Classification"),
            TaskOwner: this.getInternalName("ToDo", "Task Owner", "TaskOwner"),
            AssigneeInternal: this.getInternalName("ToDo", "Assigne Internal", "AssigneInternal"),
            AssigneeExternal: this.getInternalName("ToDo", "Assigne External", "AssigneExternal"),
            Regarding: this.getInternalName("ToDo", "Regarding", "Regarding"),
            TrainingInduction: this.getInternalName("ToDo", "Training & Induction", "TrainingInduction"),
        };

        const fieldTypeMap = this._fieldTypeMaps.get("ToDo")!;
        const fieldInternalNames = this._listFieldInternalNames.get("ToDo") || [];

        const selects = ["*", "Id", "Title", "Author/Title", "Author/EMail", "Editor/Title", "Editor/EMail", "Created", "Modified", "Attachments"];
        const expands = ["Author", "Editor", "AttachmentFiles"];

        const safelyAddSelect = (internalName: string): void => {
            if (fieldInternalNames.indexOf(internalName) < 0) return;
            const fieldType = fieldTypeMap.get(internalName) || "";
            if (fieldType === "User" || fieldType === "UserMulti") {
                selects.push(`${internalName}/Id`, `${internalName}/Title`, `${internalName}/EMail`);
                expands.push(internalName);
            } else if (fieldType === "Lookup" || fieldType === "LookupMulti") {
                selects.push(`${internalName}/Id`, `${internalName}/Title`);
                expands.push(internalName);
            } else {
                selects.push(internalName);
            }
        };

        Object.values(names).forEach(safelyAddSelect);

        try {
            let rawItems: any[] = [];
            let realSortField = sortField;
            if (sortField === "TaskOwner") realSortField = `${names.TaskOwner}/Title`;

            let query = this._sp.web.lists.getByTitle(listTitle).items
                .select(...selects).expand(...expands).orderBy(realSortField, isAscending);

            if (searchQuery) {
                // Fetch latest 1000 items and filter locally for export performance
                const rawFetched = await query.top(1000)();
                rawItems = this._applyClientSideFilter(rawFetched, searchQuery);
            } else {
                // Batch-fetch for mass export
                const itemIterator = query.top(2000);
                for await (const items of itemIterator) {
                    rawItems.push(...items);
                    if (rawItems.length >= 10000) break;
                }
            }

            return rawItems.map(item => {
                const getLookup = (n: string) => item[n] ? { Id: item[n].Id || 0, Title: item[n].Title || item[n].Name || "", Name: item[n].Name } : undefined;
                const getPerson = (n: string) => item[n] ? { Id: item[n].Id || 0, Title: item[n].Title || "", EMail: item[n].EMail || "" } : undefined;
                return {
                    ...item,
                    Id: item.Id,
                    Title: item[names.Subject] || item.Title || "",
                    Status: getLookup(names.Status),
                    Category: getLookup(names.Category),
                    Priority: getLookup(names.Priority),
                    Classification: getLookup(names.Classification),
                    TaskOwner: getPerson(names.TaskOwner),
                    AssigneeInternal: getPerson(names.AssigneeInternal),
                    AssigneeExternal: getPerson(names.AssigneeExternal),
                    Regarding: item[names.Regarding],
                    TrainingInduction: getLookup(names.TrainingInduction),
                    Author: getPerson("Author") || item.Author,
                    Editor: getPerson("Editor") || item.Editor,
                    Created: item.Created,
                    Modified: item.Modified,
                    AttachmentFiles: item.AttachmentFiles || [],
                } as any;
            });
        } catch (error) {
            console.error(`[SPService] Filtered query failed:`, error);
            return [];
        }
    }

    public async getToDoItems(): Promise<IToDoItem[]> {
        await this.init("ToDo");
        const fieldInternalNames = this._listFieldInternalNames.get("ToDo") || [];
        const fieldTypeMap = this._fieldTypeMaps.get("ToDo") || new Map();
        const listTitle = this._listTitles.get("ToDo")!;

        const names = {
            Subject: this.getInternalName("ToDo", "Subject", "Title"),
            TaskOwner: this.getInternalName("ToDo", "Task Owner", "TaskOwner"),
            AssigneeInternal: this.getInternalName("ToDo", "Assigne Internal", "AssigneInternal"),
            AssigneeExternal: this.getInternalName("ToDo", "Assigne External", "AssigneExternal"),
            Status: this.getInternalName("ToDo", "Status", "Status"),
            Category: this.getInternalName("ToDo", "Category", "Category"),
            Classification: this.getInternalName("ToDo", "Classification", "Classification"),
            Priority: this.getInternalName("ToDo", "Priority", "Priority"),
            CompletedPercent: this.getInternalName("ToDo", "Completed %", "CompletedPercent"),
            StartDate: this.getInternalName("ToDo", "Start Date", "StartDate"),
            CompletionDate: this.getInternalName("ToDo", "Completion Date", "CompletionDate"),
            CreatedByUser: this.getInternalName("ToDo", "Created By User", "CreatedByUser"),
            UpdatedByUser: this.getInternalName("ToDo", "Updated By User", "UpdatedByUser"),
            CreatedOn: this.getInternalName("ToDo", "Created On", "CreatedOn"),
            UpdatedOn: this.getInternalName("ToDo", "Updated On", "UpdatedOn"),
            Description: this.getInternalName("ToDo", "Description", "Description"),
            Regarding: this.getInternalName("ToDo", "Regarding", "Regarding"),
            DueDate: this.getInternalName("ToDo", "Due Date", "DueDate"),
            Resolution: this.getInternalName("ToDo", "Resolution", "Resolution"),
            EmailNotifications: this.getInternalName("ToDo", "Email Notification", "EmailNotifications"),
            TrainingInduction: this.getInternalName("ToDo", "Training & Induction", "TrainingInduction"),
        };

        const selects = ["*", "Id", "Title", "Author/Title", "Author/EMail", "Editor/Title", "Editor/EMail", "Created", "Modified", "Attachments"];
        const expands = ["Author", "Editor", "AttachmentFiles"];

        const safelyAddSelect = (internalName: string): void => {
            if (fieldInternalNames.indexOf(internalName) < 0) return;
            const fieldType = fieldTypeMap.get(internalName) || "";

            if (fieldType === "User" || fieldType === "UserMulti") {
                selects.push(`${internalName}/Id`, `${internalName}/Title`, `${internalName}/EMail`);
                expands.push(internalName);
            } else if (fieldType === "Lookup" || fieldType === "LookupMulti") {
                selects.push(`${internalName}/Id`, `${internalName}/Title`);
                expands.push(internalName);
            } else {
                selects.push(internalName);
            }
        };

        Object.values(names).forEach(safelyAddSelect);

        try {
            const itemIterator = this._sp.web.lists.getByTitle(listTitle).items
                .select(...selects)
                .expand(...expands)
                .top(2000)
                .orderBy("Id", true);

            const rawItems: any[] = [];
            for await (const items of itemIterator) {
                rawItems.push(...items);
                if (rawItems.length >= 5000) break;
            }

            return rawItems.map((item: any) => {
                const getLookup = (n: string) => item[n] ? { Id: item[n].Id || 0, Title: item[n].Title || item[n].Name || "", Name: item[n].Name } : undefined;
                const getPerson = (n: string) => item[n] ? { Id: item[n].Id || 0, Title: item[n].Title || "", EMail: item[n].EMail || "" } : undefined;

                return {
                    ...item,
                    Id: item.Id,
                    Title: item[names.Subject] || item.Title || "",
                    Description: item[names.Description],
                    Status: getLookup(names.Status),
                    Category: getLookup(names.Category),
                    Classification: getLookup(names.Classification),
                    Priority: getLookup(names.Priority),
                    TaskOwner: getPerson(names.TaskOwner),
                    AssigneeInternal: getPerson(names.AssigneeInternal),
                    AssigneeExternal: getPerson(names.AssigneeExternal),
                    Regarding: item[names.Regarding],
                    TrainingInduction: getLookup(names.TrainingInduction),
                    DueDate: item[names.DueDate],
                    StartDate: item[names.StartDate],
                    CompletionDate: item[names.CompletionDate],
                    CompletedPercent: item[names.CompletedPercent],
                    EmailNotifications: item[names.EmailNotifications],
                    Author: getPerson(names.CreatedByUser) || item.Author,
                    Editor: getPerson(names.UpdatedByUser) || item.Editor,
                    Created: item[names.CreatedOn] || item.Created,
                    Modified: item[names.UpdatedOn] || item.Modified,
                    Resolution: item[names.Resolution],
                    AttachmentFiles: item.AttachmentFiles || [],
                };
            });
        } catch (error) {
            console.error(`Query failed on ${listTitle}:`, error);
            this._lastError = `Fetch failed: ${error.message || JSON.stringify(error)}`;
            return [];
        }
    }

    public async addToDoItem(item: any): Promise<any> {
        await this.init("ToDo");
        const listTitle = this._listTitles.get("ToDo")!;
        const fieldMap = this._fieldMaps.get("ToDo")!;
        const fieldTypeMap = this._fieldTypeMaps.get("ToDo")!;
        const listFieldInternalNames = this._listFieldInternalNames.get("ToDo") || [];

        const cleaned: any = {};
        const currentUser = await this._sp.web.currentUser();
        const now = new Date().toISOString();

        const names = {
            CreatedOn: this.getInternalName("ToDo", "Created On", "CreatedOn"),
            CreatedBy: this.getInternalName("ToDo", "Created By User", "CreatedByUser"),
            UpdatedOn: this.getInternalName("ToDo", "Updated On", "UpdatedOn"),
            UpdatedBy: this.getInternalName("ToDo", "Updated By User", "UpdatedByUser"),
        };

        // Helper to find internal name for a payload key
        const resolveInternal = (key: string): { name: string, isId: boolean } | null => {
            if (key === "Title") return { name: "Title", isId: false };

            const isId = key.endsWith("Id");
            const baseKey = isId ? key.slice(0, -2) : key;

            // Try direct mapping
            let internal = fieldMap.get(key) || fieldMap.get(baseKey);

            // Try fuzzy match
            if (!internal) {
                const search = baseKey.toLowerCase().replace(/[\s&_]+/g, '');
                internal = listFieldInternalNames.find(n => n.toLowerCase().replace(/[\s_]+/g, '') === search);
            }

            if (internal) return { name: internal, isId };
            return null;
        };

        Object.keys(item).forEach(key => {
            const res = resolveInternal(key);
            if (!res) return;

            const fieldType = fieldTypeMap.get(res.name);
            const saveKey = (res.isId && (fieldType === "Lookup" || fieldType === "User" || fieldType === "LookupMulti" || fieldType === "UserMulti"))
                ? `${res.name}Id`
                : res.name;

            if (item[key] !== undefined && item[key] !== null) {
                cleaned[saveKey] = item[key];
            }
        });

        // Add metadata
        if (listFieldInternalNames.indexOf(names.CreatedOn) > -1) cleaned[names.CreatedOn] = now;
        if (listFieldInternalNames.indexOf(names.CreatedBy) > -1) cleaned[`${names.CreatedBy}Id`] = currentUser.Id;
        if (listFieldInternalNames.indexOf(names.UpdatedOn) > -1) cleaned[names.UpdatedOn] = now;
        if (listFieldInternalNames.indexOf(names.UpdatedBy) > -1) cleaned[`${names.UpdatedBy}Id`] = currentUser.Id;

        console.log(`[SPService] addToDoItem Payload for ${listTitle}:`, cleaned);
        try {
            return await this._sp.web.lists.getByTitle(listTitle).items.add(cleaned);
        } catch (e) {
            console.error(`[SPService] addToDoItem failed for ${listTitle}. Payload:`, cleaned, e);
            throw e;
        }
    }

    public async updateToDoItem(id: number, item: any): Promise<any> {
        await this.init("ToDo");
        const listTitle = this._listTitles.get("ToDo")!;
        const fieldMap = this._fieldMaps.get("ToDo")!;
        const fieldTypeMap = this._fieldTypeMaps.get("ToDo")!;
        const listFieldInternalNames = this._listFieldInternalNames.get("ToDo") || [];

        const cleaned: any = {};
        const currentUser = await this._sp.web.currentUser();
        const now = new Date().toISOString();

        const names = {
            UpdatedOn: this.getInternalName("ToDo", "Updated On", "UpdatedOn"),
            UpdatedBy: this.getInternalName("ToDo", "Updated By User", "UpdatedByUser"),
        };

        const resolveInternal = (key: string): { name: string, isId: boolean } | null => {
            if (key === "Title") return { name: "Title", isId: false };
            const isId = key.endsWith("Id");
            const baseKey = isId ? key.slice(0, -2) : key;
            let internal = fieldMap.get(key) || fieldMap.get(baseKey);
            if (!internal) {
                const search = baseKey.toLowerCase().replace(/[\s&_]+/g, '');
                internal = listFieldInternalNames.find(n => n.toLowerCase().replace(/[\s_]+/g, '') === search);
            }
            if (internal) return { name: internal, isId };
            return null;
        };

        Object.keys(item).forEach(key => {
            const res = resolveInternal(key);
            if (!res) return;
            const fieldType = fieldTypeMap.get(res.name);
            const saveKey = (res.isId && (fieldType === "Lookup" || fieldType === "User" || fieldType === "LookupMulti" || fieldType === "UserMulti"))
                ? `${res.name}Id`
                : res.name;
            if (item[key] !== undefined && item[key] !== null) cleaned[saveKey] = item[key];
        });

        if (listFieldInternalNames.indexOf(names.UpdatedOn) > -1) cleaned[names.UpdatedOn] = now;
        if (listFieldInternalNames.indexOf(names.UpdatedBy) > -1) cleaned[`${names.UpdatedBy}Id`] = currentUser.Id;

        console.log(`[SPService] updateToDoItem Payload for ${listTitle}:`, cleaned);
        try {
            return await this._sp.web.lists.getByTitle(listTitle).items.getById(id).update(cleaned);
        } catch (e) {
            console.error(`[SPService] updateToDoItem failed for ${listTitle}. Payload:`, cleaned, e);
            throw e;
        }
    }

    public async deleteToDoItem(id: number): Promise<void> {
        await this.init("ToDo");
        const listTitle = this._listTitles.get("ToDo")!;
        await this._sp.web.lists.getByTitle(listTitle).items.getById(id).delete();
    }

    // ─── Training & Inductions Methods ────────────────────────────────────────

    public async getTrainingInductionTotalCount(searchQuery?: string): Promise<number> {
        const listName = "TrainingInductions";
        await this.init(listName);
        const listTitle = this._listTitles.get(listName)!;
        try {
            const names = {
                Title: this.getInternalName(listName, "Title", "Title"),
                Type: this.getInternalName(listName, "Type", "Type"),
                TrainingType: this.getInternalName(listName, "Trainign Type", "TrainingType"),
                Participant: this.getInternalName(listName, "Participants", "Participants"),
            };
            let filterString = "";
            if (searchQuery) {
                // Prioritize Type and Title search
                const q = searchQuery.replace(/'/g, "''");
                filterString = `(substringof('${q}', '${names.Title}') or substringof('${q}', '${names.Type}') or substringof('${q}', '${names.TrainingType}'))`;
                if (!isNaN(Number(searchQuery))) filterString = `(Id eq ${searchQuery} or ${filterString})`;
            }
            let query = this._sp.web.lists.getByTitle(listTitle).items.select("Id").top(5000);
            if (filterString) query = query.filter(filterString);
            const items = await query();
            return items.length;
        } catch (e) {
            console.error("[SPService] Count failed:", e);
            return 0;
        }
    }

    public async getTrainingInductionItemsPaged(
        page: number,
        pageSize: number,
        searchQuery?: string,
        sortField: string = "Id",
        isAscending: boolean = true
    ): Promise<ITrainingInductionItem[]> {
        const listName = "TrainingInductions";
        await this.init(listName);
        const listTitle = this._listTitles.get(listName)!;
        const fieldTypeMap = this._fieldTypeMaps.get(listName)!;
        const fieldInternalNames = this._listFieldInternalNames.get(listName) || [];

        // Comprehensive name resolution
        const names: any = {
            Title: this.getInternalName(listName, "Title", "Title"),
            Type: this.getInternalName(listName, "Type", "Type"),
            Status: this.getInternalName(listName, "Status", "Status"),
            TrainingFor: this.getInternalName(listName, "Training For", "TrainingFor"),
            TrainingType: this.getInternalName(listName, "Trainign Type", "TrainingType"),
            BusinessProfile: this.getInternalName(listName, "Business Profile", "BusinessProfile"),
            ScheduledDate: this.getInternalName(listName, "Scheduled Date", "ScheduledDate"),
            InductionLink: this.getInternalName(listName, "Induction Link", "InductionLink"),
            InvitationSection: this.getInternalName(listName, "Invitation Section", "InvitationSection"),
            InvitationStatus: this.getInternalName(listName, "Invitation Status", "InvitationStatus"),
            Participant: this.getInternalName(listName, "Participants", "Participants") || this.getInternalName(listName, "Participant", "Participant"),
            ParticipantsStatus: this.getInternalName(listName, "Participants Status", "ParticipantsStatus"),
            Manager: this.getInternalName(listName, "Manager", "Manager"),
            Supervisors: this.getInternalName(listName, "Supervisors", "Supervisors"),
            Coordinator: this.getInternalName(listName, "Coordinator", "Coordinator"),
            Project: this.getInternalName(listName, "Project", "Project"),
            CompletionDate: this.getInternalName(listName, "Completion Date", "CompletionDate"),
            SendInvitation: this.getInternalName(listName, "Send Invitation", "SendInvitation"),
            Company: this.getInternalName(listName, "Company", "Company"),
        };

        const selects = ["Id", "Attachments", "Created", "Modified", "Author/Title", "Editor/Title"];
        const expands = ["Author", "Editor"];

        Object.keys(names).forEach(key => {
            const internalName = (names as any)[key];
            if (!internalName || fieldInternalNames.indexOf(internalName) < 0) return;

            const fieldType = fieldTypeMap.get(internalName) || "";
            if (fieldType === "User" || fieldType === "UserMulti" || fieldType === "Lookup" || fieldType === "LookupMulti") {
                // Standard properties
                selects.push(`${internalName}/Id`, `${internalName}/Title`);

                // Add known good sub-fields if we've discovered them
                const displayField = (this._lookupDisplayFields.get(listName) || new Map()).get(internalName);
                if (displayField && displayField !== "Title" && displayField !== "ID") {
                    selects.push(`${internalName}/${displayField}`);
                }

                expands.push(internalName);
                if (fieldType === "User" || fieldType === "UserMulti") selects.push(`${internalName}/EMail`);
            } else {
                selects.push(internalName);
            }
        });

        // Add common lookup variants to ensure we catch them (safe subset only)
        ["Participants", "Participant", "Business_x0020_Profile", "BusinessProfile", "Projects", "Project", "Companies", "Company"].forEach(v => {
            if (fieldInternalNames.indexOf(v) > -1 && expands.indexOf(v) < 0) {
                const ft = fieldTypeMap.get(v);
                if (ft === "Lookup" || ft === "LookupMulti") {
                    selects.push(`${v}/Id`, `${v}/Title`);
                    const displayField = (this._lookupDisplayFields.get(listName) || new Map()).get(v);
                    if (displayField && displayField !== "Title" && displayField !== "ID") {
                        selects.push(`${v}/${displayField}`);
                    }
                    expands.push(v);
                }
            }
        });

        const skipCount = (page - 1) * pageSize;
        try {
            let baseQuery = this._sp.web.lists.getByTitle(listTitle).items
                .select(...selects).expand(...expands).top(pageSize).orderBy(sortField, isAscending);

            if (searchQuery) {
                const q = searchQuery.replace(/'/g, "''");
                let filter = `(substringof('${q}', '${names.Title}') or substringof('${q}', '${names.Type}'))`;
                baseQuery = baseQuery.filter(filter);
            }

            const rawItems = await (skipCount > 0 ? baseQuery.skip(skipCount)() : baseQuery());
            if (rawItems.length > 0) {
                console.log(`[SPService] Raw item keys for first row:`, Object.keys(rawItems[0]));
                console.log(`[SPService] Raw data for first Row:`, rawItems[0]);
            }

            return rawItems.map((item, idx) => {
                try {
                    // Hyper-robust lookup finder
                    const findVal = (possibleNames: string[]) => {
                        for (const n of possibleNames) {
                            if (item[n] !== undefined) return item[n];
                        }
                        // Try case-insensitive search in keys
                        const itemKeys = Object.keys(item);
                        for (const p of possibleNames) {
                            const match = itemKeys.find(k => k.toLowerCase() === p.toLowerCase());
                            if (match) return item[match];
                        }
                        return undefined;
                    };

                    const getLookup = (possibleNames: string[]) => {
                        const val = findVal(possibleNames);
                        if (!val) return undefined;

                        // Find any discovered display fields for these internal names
                        const suggestedFields: string[] = [];
                        possibleNames.forEach(n => {
                            const df = (this._lookupDisplayFields.get(listName) || new Map()).get(n);
                            if (df) suggestedFields.push(df);
                        });

                        const extract = (obj: any) => {
                            if (!obj) return undefined;
                            // Search for best display name: discovered field first, then Title, then known fallbacks
                            let bestTitle = "";
                            for (const fld of suggestedFields) {
                                if (obj[fld]) { bestTitle = obj[fld]; break; }
                            }

                            if (!bestTitle) {
                                bestTitle = obj.Title || obj.Name || obj.EmployeeName || obj.Employee_x0020_Name || obj.FullName || obj.Full_x0020_Name || obj.BusinessProfile || obj.Business_x0020_Profile || obj.CompanyName || obj.Company_x0020_Name || obj.ProjectName || obj.Project_x0020_Name || obj.Company;
                            }

                            return {
                                Id: obj.Id || 0,
                                Title: bestTitle || (typeof obj === 'string' ? obj : ''),
                                Name: obj.Name || bestTitle
                            };
                        };

                        if (Array.isArray(val)) return extract(val[0]);
                        return extract(val);
                    };

                    const getPerson = (possibleNames: string[]) => {
                        const val = findVal(possibleNames);
                        if (!val) return undefined;
                        if (Array.isArray(val)) {
                            const first = val[0];
                            return first ? { Id: first.Id || 0, Title: first.Title || "", EMail: first.EMail || "" } : undefined;
                        }
                        return { Id: val.Id || 0, Title: val.Title || "", EMail: val.EMail || "" };
                    };

                    const getVal = (possibleNames: string[]) => findVal(possibleNames);

                    return {
                        Id: item.Id,
                        Title: getVal(["Title", names.Title]) || `Item ${item.Id}`,
                        Type: getVal(["Type", "TrainignType", "TrainingType", names.Type, names.TrainingType]),
                        Status: getVal(["Status", names.Status]),
                        TrainingFor: getVal(["TrainingFor", names.TrainingFor]),
                        TrainingType: getVal(["TrainingType", names.TrainingType, "TrainignType"]),
                        BusinessProfile: getLookup(["BusinessProfile", "Business_x0020_Profile", names.BusinessProfile]),
                        ScheduledDate: getVal(["ScheduledDate", "Scheduled_x0020_Date", names.ScheduledDate]),
                        InductionLink: getVal(["InductionLink", "Induction_x0020_Link", names.InductionLink]),
                        InvitationSection: getVal(["InvitationSection", names.InvitationSection]),
                        InvitationStatus: getVal(["InvitationStatus", names.InvitationStatus]),
                        Participant: getLookup(["Participant", "Participants", names.Participant]),
                        Manager: getLookup(["Manager", names.Manager]),
                        Supervisors: getLookup(["Supervisors", names.Supervisors]),
                        Coordinator: getLookup(["Coordinator", names.Coordinator]),
                        ParticipantsStatus: getVal(["ParticipantsStatus", "Participants_x0020_Status", names.ParticipantsStatus]),
                        Project: getLookup(["Project", "Projects", names.Project]),
                        CompletionDate: getVal(["CompletionDate", "Completion_x0020_Date", names.CompletionDate]),
                        SendInvitation: getVal(["SendInvitation", names.SendInvitation]),
                        Company: getLookup(["Company", "Companies", names.Company]),
                        Author: getPerson(["Author"]) || item.Author,
                        Editor: getPerson(["Editor"]) || item.Editor,
                        Created: item.Created,
                        Modified: item.Modified,
                        AttachmentFiles: item.AttachmentFiles || [],
                    } as any;
                } catch (e) {
                    console.error(`[SPService] Robust mapping failed for item ${item.Id}:`, e);
                    return { Id: item.Id } as any;
                }
            });
        } catch (error) {
            return [];
        }
    }

    /**
     * Fetches all Training & Induction items, bypassing the 5000-item threshold 
     * by using PnPjs getAll() or manual batching.
     */
    public async getAllTrainingInductionItems(searchQuery?: string): Promise<ITrainingInductionItem[]> {
        const listName = "TrainingInductions";
        await this.init(listName);
        const listTitle = this._listTitles.get(listName)!;
        const fieldTypeMap = this._fieldTypeMaps.get(listName)!;
        const fieldInternalNames = this._listFieldInternalNames.get(listName) || [];

        const names: any = {
            Title: this.getInternalName(listName, "Title", "Title"),
            Type: this.getInternalName(listName, "Type", "Type"),
            Status: this.getInternalName(listName, "Status", "Status"),
            TrainingFor: this.getInternalName(listName, "Training For", "TrainingFor"),
            TrainingType: this.getInternalName(listName, "Trainign Type", "TrainingType"),
            BusinessProfile: this.getInternalName(listName, "Business Profile", "BusinessProfile"),
            ScheduledDate: this.getInternalName(listName, "Scheduled Date", "ScheduledDate"),
            InductionLink: this.getInternalName(listName, "Induction Link", "InductionLink"),
            InvitationSection: this.getInternalName(listName, "Invitation Section", "InvitationSection"),
            InvitationStatus: this.getInternalName(listName, "Invitation Status", "InvitationStatus"),
            Participant: this.getInternalName(listName, "Participants", "Participants") || this.getInternalName(listName, "Participant", "Participant"),
            ParticipantsStatus: this.getInternalName(listName, "Participants Status", "ParticipantsStatus"),
            Manager: this.getInternalName(listName, "Manager", "Manager"),
            Supervisors: this.getInternalName(listName, "Supervisors", "Supervisors"),
            Coordinator: this.getInternalName(listName, "Coordinator", "Coordinator"),
            Project: this.getInternalName(listName, "Project", "Project"),
            CompletionDate: this.getInternalName(listName, "Completion Date", "CompletionDate"),
            SendInvitation: this.getInternalName(listName, "Send Invitation", "SendInvitation"),
            Company: this.getInternalName(listName, "Company", "Company"),
        };

        const selects = ["Id", "Attachments", "Created", "Modified", "Author/Title", "Editor/Title"];
        const expands = ["Author", "Editor"];

        Object.keys(names).forEach(key => {
            const internalName = (names as any)[key];
            if (!internalName || fieldInternalNames.indexOf(internalName) < 0) return;
            const fieldType = fieldTypeMap.get(internalName) || "";
            if (fieldType === "User" || fieldType === "UserMulti" || fieldType === "Lookup" || fieldType === "LookupMulti") {
                selects.push(`${internalName}/Id`, `${internalName}/Title`);
                expands.push(internalName);
                if (fieldType === "User" || fieldType === "UserMulti") selects.push(`${internalName}/EMail`);
            } else {
                selects.push(internalName);
            }
        });

        try {
            let query = this._sp.web.lists.getByTitle(listTitle).items.select(...selects).expand(...expands);

            if (searchQuery) {
                const q = searchQuery.replace(/'/g, "''");
                let filter = `(substringof('${q}', '${names.Title}') or substringof('${q}', '${names.Type}'))`;
                query = query.filter(filter);
            }

            const rawItems: any[] = [];
            // Use async iterator (for-await) which is built-in to PnPjs v3/v4 items collection
            // and automatically handles pagination to bypass the threshold.
            for await (const batch of query) {
                rawItems.push(...batch);
            }

            return rawItems.map((item: any) => {
                const findVal = (possibleNames: string[]) => {
                    for (const n of possibleNames) { if (item[n] !== undefined) return item[n]; }
                    const itemKeys = Object.keys(item);
                    for (const p of possibleNames) {
                        const match = itemKeys.find(k => k.toLowerCase() === p.toLowerCase());
                        if (match) return item[match];
                    }
                    return undefined;
                };

                const getLookup = (possibleNames: string[]) => {
                    const val = findVal(possibleNames);
                    if (!val) return undefined;
                    const extract = (obj: any) => {
                        if (!obj) return undefined;
                        const t = obj.Title || obj.BusinessProfile || obj.CompanyName || obj.ProjectName || obj.Name || (typeof obj === 'string' ? obj : '');
                        return t ? { Id: obj.Id || 0, Title: t } : undefined;
                    };
                    return Array.isArray(val) ? extract(val[0]) : extract(val);
                };

                const getVal = (possibleNames: string[]) => findVal(possibleNames);

                return {
                    Id: item.Id,
                    Title: getVal(["Title", names.Title]) || `Item ${item.Id}`,
                    Type: getVal(["Type", "TrainingType", names.Type, names.TrainingType]),
                    Status: getVal(["Status", names.Status]),
                    Participant: getLookup(["Participant", "Participants", names.Participant]),
                    ParticipantsStatus: getVal(["ParticipantsStatus", names.ParticipantsStatus]),
                    Project: getLookup(["Project", "Projects", names.Project]),
                    Company: getLookup(["Company", "Companies", names.Company]),
                    BusinessProfile: getLookup(["BusinessProfile", names.BusinessProfile]),
                    Manager: getLookup(["Manager", names.Manager]),
                    Supervisors: getLookup(["Supervisors", names.Supervisors]),
                    Coordinator: getLookup(["Coordinator", names.Coordinator]),
                    ScheduledDate: getVal(["ScheduledDate", names.ScheduledDate]),
                    CompletionDate: getVal(["CompletionDate", names.CompletionDate]),
                    SendInvitation: getVal(["SendInvitation", "SendInvitations", "Send_x0020_Invitation", "Send_x0020_Invitations"]),
                    TrainingFor: getVal(["TrainingFor", "InductionFor", "Induction_x0020_For"]),
                } as any;
            });
        } catch (error) {
            console.error(`[SPService] getAll failed:`, error);
            return [];
        }
    }

    public async addTrainingInductionItem(item: any): Promise<any> {
        const listName = "TrainingInductions";
        await this.init(listName);
        const listTitle = this._listTitles.get(listName)!;
        const fieldMap = this._fieldMaps.get(listName)!;
        const fieldTypeMap = this._fieldTypeMaps.get(listName)!;
        const cleaned: any = {};

        Object.keys(item).forEach(key => {
            const isIdField = key.endsWith("Id");
            const baseKey = isIdField ? key.slice(0, -2) : key;
            const iNames = this._listFieldInternalNames.get(listName) || [];
            let internalName = fieldMap.get(key) ||
                fieldMap.get(baseKey) ||
                fieldMap.get(baseKey + "s") ||
                (iNames.indexOf(key) > -1 ? key : "") ||
                (iNames.indexOf(baseKey) > -1 ? baseKey : "") ||
                this.getInternalName(listName, baseKey, "");

            // Robust fallback for the Type field
            if (!internalName && (baseKey === "Type" || baseKey === "Title" || baseKey === "SystemForm")) {
                const iNames = this._listFieldInternalNames.get(listName) || [];
                if (iNames.indexOf("Type") > -1) internalName = "Type";
                else if (iNames.indexOf("TrainingType") > -1) internalName = "TrainingType";
                else if (iNames.indexOf("SystemForm") > -1) internalName = "SystemForm";
                else if (iNames.indexOf("System_x0020_Form") > -1) internalName = "System_x0020_Form";
            }

            if (internalName) {
                const fieldType = fieldTypeMap.get(internalName);
                const saveKey = (isIdField && (fieldType === "Lookup" || fieldType === "User" || fieldType === "LookupMulti" || fieldType === "UserMulti"))
                    ? `${internalName}Id`
                    : internalName;

                let val = item[key];

                // Explicitly cast to string if SharePoint expects a string but we have a primitive (Number/Boolean)
                const isTextField = fieldType === 'Text' || fieldType === 'Note' || fieldType === 'Choice';
                if (isTextField && val !== null && val !== undefined && typeof val !== 'string') {
                    val = String(val);
                }

                // Special handling for booleans (Yes/No fields in SharePoint)
                if (typeof val === 'boolean') {
                    const isBoolField = fieldType === 'Boolean' || fieldType === 'Boolean (yes/no)';
                    if (!isBoolField) {
                        val = val ? "Yes" : "No";
                    }
                }

                // Ensure dates are strings for REST
                if (val instanceof Date) {
                    val = val.toISOString();
                }

                cleaned[saveKey] = val;
            } else if (key === "Title") {
                cleaned.Title = item[key];
            }
        });

        console.log(`[SPService] Saving to ${listTitle}:`, cleaned);
        return await this._sp.web.lists.getByTitle(listTitle).items.add(cleaned);
    }

    public async updateTrainingInductionItem(id: number, item: Record<string, any>): Promise<any> {
        const listName = "TrainingInductions";
        await this.init(listName);
        const listTitle = this._listTitles.get(listName)!;
        const fieldMap = this._fieldMaps.get(listName)!;
        const fieldTypeMap = this._fieldTypeMaps.get(listName)!;
        const cleaned: Record<string, any> = {};

        Object.keys(item).forEach(key => {
            const isIdField = key.endsWith("Id");
            const baseKey = isIdField ? key.slice(0, -2) : key;
            let internalName = fieldMap.get(key) ||
                fieldMap.get(baseKey) ||
                fieldMap.get(baseKey + "s") ||
                this.getInternalName(listName, baseKey, "");

            // Robust fallback for the Type field
            if (!internalName && (baseKey === "Type" || baseKey === "Title" || baseKey === "SystemForm")) {
                const iNames = this._listFieldInternalNames.get(listName) || [];
                if (iNames.indexOf("Type") > -1) internalName = "Type";
                else if (iNames.indexOf("TrainingType") > -1) internalName = "TrainingType";
                else if (iNames.indexOf("SystemForm") > -1) internalName = "SystemForm";
                else if (iNames.indexOf("System_x0020_Form") > -1) internalName = "System_x0020_Form";
            }

            if (internalName) {
                const fieldType = fieldTypeMap.get(internalName);
                const saveKey = (isIdField && (fieldType === "Lookup" || fieldType === "User" || fieldType === "LookupMulti" || fieldType === "UserMulti"))
                    ? `${internalName}Id`
                    : internalName;

                let val = item[key];

                // Explicitly cast to string if SharePoint expects a string but we have a primitive (Number/Boolean)
                const isTextField = fieldType === 'Text' || fieldType === 'Note' || fieldType === 'Choice';
                if (isTextField && val !== null && val !== undefined && typeof val !== 'string') {
                    val = String(val);
                }

                // Special handling for booleans (Yes/No fields in SharePoint)
                if (typeof val === 'boolean') {
                    const isBoolField = fieldType === 'Boolean' || fieldType === 'Boolean (yes/no)';
                    if (!isBoolField) {
                        val = val ? "Yes" : "No";
                    }
                }

                if (val instanceof Date) {
                    val = val.toISOString();
                }

                cleaned[saveKey] = val;
            } else if (key === "Title") {
                cleaned.Title = item[key];
            }
        });
        console.log(`[SPService] Updating record ${id} in ${listTitle}:`, cleaned);
        return await this._sp.web.lists.getByTitle(listTitle).items.getById(id).update(cleaned);
    }

    public async deleteTrainingInductionItem(id: number): Promise<void> {
        const listName = "TrainingInductions";
        await this.init(listName);
        const listTitle = this._listTitles.get(listName)!;
        await this._sp.web.lists.getByTitle(listTitle).items.getById(id).delete();
    }

    // ─── Shared Lookup & Attachment Methods ───────────────────────────────────

    public async getLookupOptions(listUrlName: string, displayField: string = "Title"): Promise<ILookupOption[]> {
        try {
            const listTitle = await this.findListTitle(listUrlName);
            const list = this._sp.web.lists.getByTitle(listTitle);

            // Fetch fields to verify displayField
            const fields = await list.fields.select("InternalName")();
            let realField = "Title";
            const fieldNames = fields.map(f => f.InternalName);
            const check = (name: string): boolean => fieldNames.indexOf(name) > -1;

            if (check(displayField)) {
                realField = displayField;
            } else if (displayField.includes(" ") && check(displayField.replace(/\s+/g, '_x0020_'))) {
                realField = displayField.replace(/\s+/g, '_x0020_');
            } else if (check("EmployeeName")) {
                realField = "EmployeeName";
            } else if (check("Employee_x0020_Name")) {
                realField = "Employee_x0020_Name";
            } else if (check("Employee_Name")) {
                realField = "Employee_Name";
            } else if (check("BusinessProfile")) {
                realField = "BusinessProfile";
            } else if (check("FullName")) {
                realField = "FullName";
            } else if (check("Name")) {
                realField = "Name";
            }

            const items = await list.items.select("Id", realField, "Title").top(5000)();
            return items.map(item => ({ Id: item.Id, Title: item[realField] || item.Title || `Record ${item.Id}` }));
        } catch (error) {
            console.error(`[SPService] getLookupOptions failed for ${listUrlName}:`, error);
            return [];
        }
    }

    public async getAttachments(listName: string, itemId: number): Promise<IAttachment[]> {
        await this.init(listName);
        const listTitle = this._listTitles.get(listName)!;
        try { return await this._sp.web.lists.getByTitle(listTitle).items.getById(itemId).attachmentFiles(); }
        catch { return []; }
    }

    public async uploadAttachment(listName: string, itemId: number, file: File): Promise<void> {
        await this.init(listName);
        const listTitle = this._listTitles.get(listName)!;
        await this._sp.web.lists.getByTitle(listTitle).items.getById(itemId).attachmentFiles.add(file.name, file);
    }

    public async deleteAttachment(listName: string, itemId: number, fileName: string): Promise<void> {
        await this.init(listName);
        const listTitle = this._listTitles.get(listName)!;
        await this._sp.web.lists.getByTitle(listTitle).items.getById(itemId).attachmentFiles.getByName(fileName).delete();
    }

    // ─── Custom Document Library Methods ──────────────────────────────────────

    public async getLibraryDocuments(libraryName: string, recordId: number): Promise<any[]> {
        try {
            const listTitle = await this.findListTitle(libraryName);
            const list = this._sp.web.lists.getByTitle(listTitle);

            // 1. Resolve internal names with heavy logging
            const fields = await list.fields.select("Title", "InternalName", "TypeAsString").top(5000)();
            console.log(`[SPService] Found ${fields.length} fields in total.`);

            // Normalize: strip whitespace for loose field-name matching
            const normField = (s: string) => s.toLowerCase().replace(/\s+/g, '');
            const findField = (name: string) => {
                const normName = normField(name);
                const match = fields.find(f =>
                    f.InternalName.toLowerCase() === name.toLowerCase() ||
                    f.Title.toLowerCase() === name.toLowerCase() ||
                    normField(f.InternalName) === normName ||
                    normField(f.Title) === normName
                );
                if (match) console.log(`[SPService] Resolved "${name}" → "${match.InternalName}" (Type: ${match.TypeAsString})`);
                else console.warn(`[SPService] Could not resolve field "${name}"`);
                return match;
            };

            const recordIdField = fields.find(f =>
                f.InternalName === "RecordID" ||
                f.Title === "RecordID" ||
                f.InternalName === "Record_x0020_ID" ||
                normField(f.Title) === "recordid" ||
                normField(f.InternalName) === "recordid"
            );
            const recordIdInternalName = recordIdField ? recordIdField.InternalName : "RecordID";
            const isRecordIdLookup = recordIdField?.TypeAsString?.includes("Lookup") ?? false;
            console.log(`[SPService] RecordID: "${recordIdInternalName}", isLookup: ${isRecordIdLookup}`);

            const docTypeField = findField("DocumentType");
            const docTypeInternalName = docTypeField ? docTypeField.InternalName : "DocumentType";
            const isDocTypeLookup = docTypeField?.TypeAsString?.includes("Lookup") ?? false;

            // User confirmed: _ExtendedDescription is the internal name
            const descriptionField = fields.find(f => f.InternalName === "_ExtendedDescription" || f.Title === "Description");
            const descriptionInternalName = descriptionField ? descriptionField.InternalName : "_ExtendedDescription";

            // 2. Build Select and Expand arrays dynamically based on resolved fields
            const selectFields = ["Id", "FileLeafRef", "FileRef", "Title", "Created", "Author/Title"];
            const expandFields = ["Author"];

            // Only expand DocType if it is actually a Lookup field
            if (docTypeField && isDocTypeLookup) {
                selectFields.push(`${docTypeInternalName}/Id`, `${docTypeInternalName}/Name`, `${docTypeInternalName}/Title`);
                expandFields.push(docTypeInternalName);
            }
            if (descriptionField) {
                selectFields.push(descriptionInternalName);
            }
            // Only expand RecordID if it is a Lookup; otherwise select the plain column
            if (recordIdField) {
                if (isRecordIdLookup) {
                    selectFields.push(`${recordIdInternalName}/Id`);
                    expandFields.push(recordIdInternalName);
                } else {
                    selectFields.push(recordIdInternalName);
                }
            }

            // 3. Fetch items (OData filter syntax differs for Lookup vs plain Number)
            let items: any[] = [];
            const filterExpr = isRecordIdLookup
                ? `${recordIdInternalName}Id eq ${recordId}`
                : `${recordIdInternalName} eq ${recordId}`;
            try {
                items = await list.items
                    .select(...selectFields)
                    .expand(...expandFields)
                    .filter(filterExpr)
                    .orderBy("Created", false)();

                console.log(`[SPService] OData filter "${filterExpr}" returned ${items.length} items.`);
            } catch (odataErr) {
                console.warn(`[SPService] OData filter failed:`, odataErr);
            }

            // 4. Fallback: memory filter if OData returns zero or fails
            if (items.length === 0) {
                console.log(`[SPService] Falling back to memory filter...`);
                try {
                    const allItems = await list.items
                        .select(...selectFields)
                        .expand(...expandFields)
                        .top(5000)();

                    console.log(`[SPService] Memory Filter: Fetched ${allItems.length} total items.`);
                    if (allItems.length > 0) {
                        console.log(`[SPService] Sample Item [0] keys:`, Object.keys(allItems[0]));
                        console.log(`[SPService] Sample Item [0] ${recordIdInternalName}:`, allItems[0][recordIdInternalName]);
                    }

                    items = allItems.filter(item => {
                        const val = item[recordIdInternalName];
                        if (!val) return false;

                        if (isRecordIdLookup) {
                            // Handle both single lookup and multi-lookup
                            const ids: number[] = [];
                            if (Array.isArray(val)) {
                                val.forEach(v => { if (v && v.Id) ids.push(Number(v.Id)); });
                            } else if (typeof val === 'object') {
                                if (val.Id) ids.push(Number(val.Id));
                            } else {
                                ids.push(Number(val));
                            }
                            return ids.indexOf(Number(recordId)) > -1;
                        }

                        return String(val) === String(recordId);
                    });
                    console.log(`[SPService] Memory filter narrowed to ${items.length} units.`);
                } catch (memErr) {
                    console.error("[SPService] Memory filter fetch failed:", memErr);
                    items = []; // Never show unrelated documents
                }
            }

            return items.map(item => {
                const dt = item[docTypeInternalName];
                return {
                    Id: item.Id,
                    Title: item.Title || item.FileLeafRef,
                    FileName: item.FileLeafRef,
                    ServerRelativeUrl: item.FileRef,
                    DocumentType: dt ? { Id: dt.Id, Title: dt.Name || dt.Title || "" } : { Id: 0, Title: "" },
                    Description: item[descriptionInternalName] || "",
                    Created: item.Created,
                    Author: { Title: item.Author?.Title || "Unknown" }
                };
            });
        } catch (e) {
            console.error(`[SPService] getLibraryDocuments failed for ${libraryName}:`, e);
            return [];
        }
    }

    public async uploadLibraryDocument(libraryName: string, recordId: number, file: File): Promise<void> {
        try {
            const listTitle = await this.findListTitle(libraryName);
            const list = this._sp.web.lists.getByTitle(listTitle);

            const fields = await list.fields.select("Title", "InternalName").top(5000)();
            const recordIdField = fields.find(f => f.Title === "RecordID" || f.InternalName === "RecordID");
            const recordIdInternalName = recordIdField ? recordIdField.InternalName : "RecordID";

            const result = await list.rootFolder.files.addUsingPath(file.name, file, { Overwrite: true });
            const item = await this._sp.web.getFileByServerRelativePath(result.ServerRelativeUrl).getItem();

            const updatePayload: any = {
                Title: file.name.split('.').slice(0, -1).join('.')
            };
            updatePayload[`${recordIdInternalName}Id`] = recordId;

            await item.update(updatePayload);
        } catch (e) {
            console.error(`[SPService] uploadLibraryDocument failed:`, e);
            throw e;
        }
    }

    public async deleteLibraryDocument(libraryName: string, documentId: number): Promise<void> {
        const listTitle = await this.findListTitle(libraryName);
        await this._sp.web.lists.getByTitle(listTitle).items.getById(documentId).delete();
    }

    public async updateLibraryDocumentMetadata(libraryName: string, docId: number, metadata: { Title: string, DocumentType: number, Description: string }): Promise<void> {
        try {
            const listTitle = await this.findListTitle(libraryName);
            const list = this._sp.web.lists.getByTitle(listTitle);

            const fields = await list.fields.select("Title", "InternalName", "TypeAsString").top(5000)();
            const normF = (s: string) => s.toLowerCase().replace(/\s+/g, '');
            const findFieldMeta = (name: string) => {
                const n = normF(name);
                return fields.find(f =>
                    f.InternalName.toLowerCase() === name.toLowerCase() ||
                    f.Title.toLowerCase() === name.toLowerCase() ||
                    normF(f.InternalName) === n ||
                    normF(f.Title) === n
                );
            };

            const docTypeField = findFieldMeta("DocumentType");
            const docTypeInternalName = docTypeField ? docTypeField.InternalName : "DocumentType";
            const isDocTypeLookup = docTypeField?.TypeAsString?.includes("Lookup") ?? false;

            const descField = fields.find(f => f.InternalName === "_ExtendedDescription" || f.Title === "Description");
            const descriptionInternalName = descField ? descField.InternalName : "_ExtendedDescription";

            const updatePayload: any = {
                Title: metadata.Title,
                [descriptionInternalName]: metadata.Description
            };

            // Only set lookup Id when the field is actually a Lookup type
            if (metadata.DocumentType > 0 && isDocTypeLookup) {
                updatePayload[`${docTypeInternalName}Id`] = metadata.DocumentType;
            }

            await list.items.getById(docId).update(updatePayload);
        } catch (e) {
            console.error(`[SPService] updateLibraryDocumentMetadata failed:`, e);
            throw e;
        }
    }

    private async findListTitle(urlName: string): Promise<string> {
        try {
            const lists = await this._sp.web.lists.select("Title", "RootFolder/Name").expand("RootFolder")();
            const clean = (s: string): string => s.toLowerCase().replace(/[^a-z0-9]/g, '');
            const target = clean(urlName);

            const match = lists.find(l =>
                clean(l.Title) === target ||
                clean(l.RootFolder.Name) === target ||
                l.Title.toLowerCase() === urlName.toLowerCase()
            );

            if (match) {
                console.log(`[SPService] Resolved "${urlName}" to "${match.Title}"`);
                return match.Title;
            }
            return urlName;
        } catch (e) {
            console.error(`[SPService] findListTitle failed for ${urlName}:`, e);
            return urlName;
        }
    }
}
