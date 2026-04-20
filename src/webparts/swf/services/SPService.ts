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

export class SPService {
    private _sp: SPFI;
    private _fieldMap: Map<string, string> = new Map();
    // ─── NEW: stores TypeAsString for each internal name ──────────────────────
    private _fieldTypeMap: Map<string, string> = new Map();
    private _listTitle: string = "";
    private _listId: string = "";
    private _listUrl: string = "";
    private _isInitialized: boolean = false;
    private _lastError: string = "";

    constructor(private context: WebPartContext) {
        this._sp = spfi().using(SPFx(context));
    }

    public async init(): Promise<void> {
        if (this._isInitialized) return;
        try {
            const urlName = "ToDo";
            const lists = await this._sp.web.lists.select("Title", "RootFolder/Name").expand("RootFolder")();
            const match = lists.find(l =>
                l.Title.toLowerCase() === urlName.toLowerCase() ||
                l.RootFolder.Name.toLowerCase() === urlName.toLowerCase() ||
                l.Title.replace(/\s+/g, '').toLowerCase() === urlName.toLowerCase()
            );
            this._listTitle = match ? match.Title : "ToDo";

            const listData = await this._sp.web.lists.getByTitle(this._listTitle).select("Id", "Title", "RootFolder/ServerRelativeUrl").expand("RootFolder")();
            this._listId = listData.Id;
            this._listUrl = listData.RootFolder.ServerRelativeUrl;
            
            // Normalize URL for Search: ensure absolute path and no trailing slash for the prefix
            const webUrl = this.context.pageContext.web.absoluteUrl;
            const siteUrl = webUrl.substring(0, webUrl.indexOf(this.context.pageContext.web.serverRelativeUrl));
            this._listUrl = siteUrl + this._listUrl;
            if (this._listUrl.endsWith('/')) this._listUrl = this._listUrl.slice(0, -1);

            console.log(`[SPService] List Identified: "${this._listTitle}" (Path: ${this._listUrl}, ID: ${this._listId})`);

            const fields = await this._sp.web.lists.getByTitle(this._listTitle).fields
                .select("Title", "InternalName", "TypeAsString")();

            // Store BOTH display→internal name map AND internal→type map with trimmed keys
            fields.forEach(f => {
                const title = (f.Title || "").trim();
                const internal = (f.InternalName || "").trim();
                const type = (f.TypeAsString || "").trim();

                this._fieldMap.set(title, internal);
                this._fieldTypeMap.set(internal, type);
            });

            console.log("[SPService] Discovered Fields Mapping:", Array.from(this._fieldMap.entries()));
            console.log("[SPService] Field Types Mapping:", Array.from(this._fieldTypeMap.entries()));

            this._isInitialized = true;
        } catch (error) {
            console.error("[SPService] Initialization Failed:", error);
            this._listTitle = "ToDo";
            this._isInitialized = true;
        }
    }

    public getDiscoveredFields(): any {
        const obj: any = { ListTitle: this._listTitle };
        this._fieldMap.forEach((v, k) => { obj[k] = v; });
        return obj;
    }

    public getLastError(): string { return this._lastError; }

    public getInternalName(title: string, fallback: string): string {
        if (this._fieldMap.has(title)) return this._fieldMap.get(title)!;
        const keys = Array.from(this._fieldMap.keys());
        const match = keys.find(k =>
            k.toLowerCase() === title.toLowerCase() ||
            k.replace(/\s/g, '').toLowerCase() === title.toLowerCase()
        );
        if (match) return this._fieldMap.get(match)!;
        return fallback;
    }

    public async checkMissingFields(): Promise<{ title: string; internalName: string; exists: boolean }[]> {
        await this.init();
        const required = [
            { title: "Subject", internal: "Subject" },
            { title: "Category", internal: "Category" },
            { title: "Status", internal: "Status" },
            { title: "Priority", internal: "Priority" },
            { title: "Due Date", internal: "DueDate" },
            { title: "Task Owner", internal: "TaskOwner" },
        ];
        const fieldInternalNames = Array.from(this._fieldMap.values());
        const fieldTitles = Array.from(this._fieldMap.keys());
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

    public async getToDoTotalCount(searchQuery?: string): Promise<number> {
        await this.init();
        try {
            const listData = await this._sp.web.lists.getByTitle(this._listTitle).select("ItemCount")();
            const totalItems = listData.ItemCount || 0;

            if (searchQuery) {
                // Scaling strategy: Use Search API with RowLimit 0 to get total matching results
                // Avoiding substringof on 50,000+ items
                const kql = `${this._buildKQLQuery(searchQuery)} (path:"${this._listUrl}" OR path:"${this._listUrl}/*")`;
                const results: SearchResults = await this._sp.search({
                    Querytext: kql,
                    RowLimit: 0,
                    SelectProperties: ["ListItemID"]
                });
                return results.TotalRows || 0;
            } else {
                return totalItems;
            }
        } catch (e) {
            console.warn("[SPService] Count logic fallback:", e);
            return 0;
        }
    }

    /** Fetches a single page of items using server-side skip/top with optional search and sort */
    public async getToDoItemsPaged(
        page: number, 
        pageSize: number, 
        searchQuery?: string,
        sortField: string = "Id",
        isAscending: boolean = true
    ): Promise<IToDoItem[]> {
        await this.init();

        const fieldInternalNames = Array.from(this._fieldMap.values());

        const names = {
            Subject:          this.getInternalName("Subject",          "Title"),
            TaskOwner:        this.getInternalName("Task Owner",       "TaskOwner"),
            AssigneeInternal: this.getInternalName("Assigne Internal", "AssigneInternal"),
            AssigneeExternal: this.getInternalName("Assigne External", "AssigneExternal"),
            Status:           this.getInternalName("Status",           "Status"),
            Category:         this.getInternalName("Category",         "Category"),
            Classification:   this.getInternalName("Classification",   "Classification"),
            Priority:         this.getInternalName("Priority",         "Priority"),
            CompletedPercent: this.getInternalName("Completed %",      "CompletedPercent"),
            StartDate:        this.getInternalName("Start Date",       "StartDate"),
            CompletionDate:   this.getInternalName("Completion Date",  "CompletionDate"),
            CreatedByUser:    this.getInternalName("Created By User",  "CreatedByUser"),
            UpdatedByUser:    this.getInternalName("Updated By User",  "UpdatedByUser"),
            CreatedOn:        this.getInternalName("Created On",       "CreatedOn"),
            UpdatedOn:        this.getInternalName("Updated On",       "UpdatedOn"),
            Description:      this.getInternalName("Description",      "Description"),
            Regarding:        this.getInternalName("Regarding",        "Regarding"),
            DueDate:          this.getInternalName("Due Date",         "DueDate"),
            Resolution:       this.getInternalName("Resolution",       "Resolution"),
            EmailNotifications: this.getInternalName("Email Notification", "EmailNotifications"),
        };

        const selects = ["*", "Id", "Title", "Author/Title", "Author/EMail", "Editor/Title", "Editor/EMail", "Created", "Modified", "Attachments", `${names.Status}/Title`, `${names.Status}/Name`, `${names.Category}/Title`, `${names.Category}/Name`, `${names.Classification}/Title`, `${names.Classification}/Name`, `${names.Priority}/Title`, `${names.Priority}/Name`];
        const expands = ["Author", "Editor", "AttachmentFiles", names.Status, names.Category, names.Classification, names.Priority];

        const safelyAddSelect = (internalName: string): void => {
            if (fieldInternalNames.indexOf(internalName) < 0) return;
            const fieldType = this._fieldTypeMap.get(internalName) || "";
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

        safelyAddSelect(names.Subject);
        safelyAddSelect(names.Status);
        safelyAddSelect(names.Category);
        safelyAddSelect(names.Classification);
        safelyAddSelect(names.Priority);
        safelyAddSelect(names.TaskOwner);
        safelyAddSelect(names.AssigneeInternal);
        safelyAddSelect(names.AssigneeExternal);
        safelyAddSelect(names.StartDate);
        safelyAddSelect(names.CompletionDate);
        safelyAddSelect(names.CompletedPercent);
        safelyAddSelect(names.CreatedByUser);
        safelyAddSelect(names.UpdatedByUser);
        safelyAddSelect(names.CreatedOn);
        safelyAddSelect(names.UpdatedOn);
        safelyAddSelect(names.Description);
        safelyAddSelect(names.Regarding);
        safelyAddSelect(names.DueDate);
        safelyAddSelect(names.Resolution);
        safelyAddSelect(names.EmailNotifications);

        const skipCount = (page - 1) * pageSize;

        try {
            let rawItems: any[] = [];
            // Mapping sort field to internal name if needed
            let realSortField = sortField;
            if (sortField === "TaskOwner") realSortField = `${names.TaskOwner}/Title`;
            else if (sortField === "Subject")  realSortField = names.Subject;
            else if (sortField === "Regarding") realSortField = names.Regarding;


            // STRATEGY: Hybrid Search for 100% scalability
            // 1. Browsing (No search): Use standard REST pagination (Total list size compatible)
            // 2. Searching: Fetch 300 latest-sorted items using indexed sort, then filter locally.
            
            if (searchQuery) {
                // Fetch a large-enough subset for client-side search (LVT safe due to indexed sort + no filter)
                let query = this._sp.web.lists.getByTitle(this._listTitle).items
                    .select(...selects)
                    .expand(...expands)
                    .orderBy(realSortField, isAscending)
                    .top(300); // Scalable cap as per requirement 2.A

                rawItems = await query();
                // Apply comprehensive client-side matching (Requirement 2.B)
                rawItems = this._applyClientSideFilter(rawItems, searchQuery);
                
                // Manual pagination for search results
                rawItems = rawItems.slice(skipCount, skipCount + pageSize);
            } else {
                // Regular browsing remains unchanged to support deep pagination
                let query = this._sp.web.lists
                    .getByTitle(this._listTitle).items
                    .select(...selects)
                    .expand(...expands)
                    .top(pageSize)
                    .orderBy(realSortField, isAscending);

                rawItems = await (skipCount > 0 ? query.skip(skipCount)() : query());
            }

            return rawItems.map(item => {
                const getLookup = (n: string): { Id: number; Title: string; Name?: string } | undefined => {
                    const v = item[n];
                    if (v === null || v === undefined || v === "") return undefined;
                    if (typeof v === "object") return { Id: v.Id || 0, Title: v.Title || v.Name || "", Name: v.Name };
                    if (typeof v === "string") return { Id: 0, Title: v };
                    return undefined;
                };
                const getPerson = (n: string): { Id: number; Title: string; EMail: string } | undefined => {
                    const v = item[n];
                    if (!v || typeof v !== "object") return undefined;
                    return { Id: v.Id || 0, Title: v.Title || "", EMail: v.EMail || "" };
                };
                return {
                    ...item,
                    Id: item.Id,
                    Title: item[names.Subject] || item.Title || "",
                    Description: item[names.Description],
                    Status:         getLookup(names.Status),
                    Category:       getLookup(names.Category),
                    Classification: getLookup(names.Classification),
                    Priority:       getLookup(names.Priority),
                    TaskOwner:        getPerson(names.TaskOwner),
                    AssigneeInternal: getPerson(names.AssigneeInternal),
                    AssigneeExternal: getPerson(names.AssigneeExternal),
                    Regarding:        item[names.Regarding],
                    DueDate:          item[names.DueDate],
                    StartDate:        item[names.StartDate],
                    CompletionDate:   item[names.CompletionDate],
                    CompletedPercent: item[names.CompletedPercent],
                    EmailNotifications: item[names.EmailNotifications],
                    Author:  getPerson(names.CreatedByUser)  || item.Author,
                    Editor:  getPerson(names.UpdatedByUser)  || item.Editor,
                    Created: item[names.CreatedOn]  || item.Created,
                    Modified: item[names.UpdatedOn] || item.Modified,
                    Resolution: item[names.Resolution],
                    AttachmentFiles: item.AttachmentFiles || [],
                };
            });
        } catch (error) {
            console.error(`[SPService] Paged query failed:`, error);
            this._lastError = `Paged fetch failed: ${error.message || JSON.stringify(error)}`;
            return [];
        }
    }

    /** Fetches all matching items for export purposes (no pagination) */
    public async getToDoItemsFiltered(
        searchQuery?: string,
        sortField: string = "Id",
        isAscending: boolean = true
    ): Promise<IToDoItem[]> {
        await this.init();

        const fieldInternalNames = Array.from(this._fieldMap.values());

        const names = {
            Subject:          this.getInternalName("Subject",          "Title"),
            TaskOwner:        this.getInternalName("Task Owner",       "TaskOwner"),
            AssigneeInternal: this.getInternalName("Assigne Internal", "AssigneInternal"),
            AssigneeExternal: this.getInternalName("Assigne External", "AssigneExternal"),
            Status:           this.getInternalName("Status",           "Status"),
            Category:         this.getInternalName("Category",         "Category"),
            Classification:   this.getInternalName("Classification",   "Classification"),
            Priority:         this.getInternalName("Priority",         "Priority"),
            CompletedPercent: this.getInternalName("Completed %",      "CompletedPercent"),
            StartDate:        this.getInternalName("Start Date",       "StartDate"),
            CompletionDate:   this.getInternalName("Completion Date",  "CompletionDate"),
            CreatedByUser:    this.getInternalName("Created By User",  "CreatedByUser"),
            UpdatedByUser:    this.getInternalName("Updated By User",  "UpdatedByUser"),
            CreatedOn:        this.getInternalName("Created On",       "CreatedOn"),
            UpdatedOn:        this.getInternalName("Updated On",       "UpdatedOn"),
            Description:      this.getInternalName("Description",      "Description"),
            Regarding:        this.getInternalName("Regarding",        "Regarding"),
            DueDate:          this.getInternalName("Due Date",         "DueDate"),
            Resolution:       this.getInternalName("Resolution",       "Resolution"),
            EmailNotifications: this.getInternalName("Email Notification", "EmailNotifications"),
        };

        const selects = ["*", "Id", "Title", "Author/Title", "Author/EMail", "Editor/Title", "Editor/EMail", "Created", "Modified", "Attachments", `${names.Status}/Title`, `${names.Status}/Name`, `${names.Category}/Title`, `${names.Category}/Name`, `${names.Classification}/Title`, `${names.Classification}/Name`, `${names.Priority}/Title`, `${names.Priority}/Name`];
        const expands = ["Author", "Editor", "AttachmentFiles", names.Status, names.Category, names.Classification, names.Priority];

        const safelyAddSelect = (internalName: string): void => {
            if (fieldInternalNames.indexOf(internalName) < 0) return;
            const fieldType = this._fieldTypeMap.get(internalName) || "";
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

        safelyAddSelect(names.Subject);
        safelyAddSelect(names.Status);
        safelyAddSelect(names.Category);
        safelyAddSelect(names.Classification);
        safelyAddSelect(names.Priority);
        safelyAddSelect(names.TaskOwner);
        safelyAddSelect(names.AssigneeInternal);
        safelyAddSelect(names.AssigneeExternal);
        safelyAddSelect(names.StartDate);
        safelyAddSelect(names.CompletionDate);
        safelyAddSelect(names.CompletedPercent);
        safelyAddSelect(names.CreatedByUser);
        safelyAddSelect(names.UpdatedByUser);
        safelyAddSelect(names.CreatedOn);
        safelyAddSelect(names.UpdatedOn);
        safelyAddSelect(names.Description);
        safelyAddSelect(names.Regarding);
        safelyAddSelect(names.DueDate);
        safelyAddSelect(names.Resolution);
        safelyAddSelect(names.EmailNotifications);

        try {
            let rawItems: any[] = [];

            if (searchQuery) {
                // Hybrid Search for exports: fetch 1000 items and filter on client side
                const query = this._sp.web.lists.getByTitle(this._listTitle).items
                    .select(...selects)
                    .expand(...expands)
                    .top(1000)
                    .orderBy("Id", false);

                const rawFetched = await query();
                const mappedFetched = rawFetched.map(item => {
                    const getLookup = (n: string): { Id: number; Title: string; Name?: string } | undefined => {
                        const v = item[n];
                        if (v === null || v === undefined || v === "") return undefined;
                        if (typeof v === "object") return { Id: v.Id || 0, Title: v.Title || v.Name || "", Name: v.Name };
                        if (typeof v === "string") return { Id: 0, Title: v };
                        return undefined;
                    };
                    const getPerson = (n: string): { Id: number; Title: string; EMail: string } | undefined => {
                        const v = item[n];
                        if (!v || typeof v !== "object") return undefined;
                        return { Id: v.Id || 0, Title: v.Title || "", EMail: v.EMail || "" };
                    };
                    return {
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
                        // ... mapping only what's needed for the filter check
                    } as IToDoItem;
                });

                const filteredMapped = this._applyClientSideFilter(mappedFetched, searchQuery);
                const filteredIds = filteredMapped.map(i => i.Id);
                rawItems = rawFetched.filter(i => filteredIds.indexOf(i.Id) >= 0);

            } else {
                // Batch-fetch for mass export
                const itemIterator = this._sp.web.lists.getByTitle(this._listTitle).items
                    .select(...selects)
                    .expand(...expands)
                    .top(2000)
                    .orderBy("Id", true);
                
                for await (const items of itemIterator) {
                    rawItems.push(...items);
                    if (rawItems.length >= 10000) break;
                }
            }

            return rawItems.map(item => {
                const getLookup = (n: string): { Id: number; Title: string; Name?: string } | undefined => {
                    const v = item[n];
                    if (v === null || v === undefined || v === "") return undefined;
                    if (typeof v === "object") return { Id: v.Id || 0, Title: v.Title || v.Name || "", Name: v.Name };
                    if (typeof v === "string") return { Id: 0, Title: v };
                    return undefined;
                };
                const getPerson = (n: string): { Id: number; Title: string; EMail: string } | undefined => {
                    const v = item[n];
                    if (!v || typeof v !== "object") return undefined;
                    return { Id: v.Id || 0, Title: v.Title || "", EMail: v.EMail || "" };
                };
                return {
                    ...item,
                    Id: item.Id,
                    Title: item[names.Subject] || item.Title || "",
                    Description: item[names.Description],
                    Status:         getLookup(names.Status),
                    Category:       getLookup(names.Category),
                    Classification: getLookup(names.Classification),
                    Priority:       getLookup(names.Priority),
                    TaskOwner:        getPerson(names.TaskOwner),
                    AssigneeInternal: getPerson(names.AssigneeInternal),
                    AssigneeExternal: getPerson(names.AssigneeExternal),
                    Regarding:        item[names.Regarding],
                    DueDate:          item[names.DueDate],
                    StartDate:        item[names.StartDate],
                    CompletionDate:   item[names.CompletionDate],
                    CompletedPercent: item[names.CompletedPercent],
                    EmailNotifications: item[names.EmailNotifications],
                    Author:  getPerson(names.CreatedByUser)  || item.Author,
                    Editor:  getPerson(names.UpdatedByUser)  || item.Editor,
                    Created: item[names.CreatedOn]  || item.Created,
                    Modified: item[names.UpdatedOn] || item.Modified,
                    Resolution: item[names.Resolution],
                    AttachmentFiles: item.AttachmentFiles || [],
                };
            });
        } catch (error) {
            console.error(`[SPService] Filtered query failed:`, error);
            return [];
        }
    }

    public async getToDoItems(): Promise<IToDoItem[]> {
        await this.init();

        const fieldInternalNames = Array.from(this._fieldMap.values());

        const names = {
            Subject:          this.getInternalName("Subject",          "Title"),
            TaskOwner:        this.getInternalName("Task Owner",       "TaskOwner"),
            AssigneeInternal: this.getInternalName("Assigne Internal", "AssigneInternal"),
            AssigneeExternal: this.getInternalName("Assigne External", "AssigneExternal"),
            Status:           this.getInternalName("Status",           "Status"),
            Category:         this.getInternalName("Category",         "Category"),
            Classification:   this.getInternalName("Classification",   "Classification"),
            Priority:         this.getInternalName("Priority",         "Priority"),
            CompletedPercent: this.getInternalName("Completed %",      "CompletedPercent"),
            StartDate:        this.getInternalName("Start Date",       "StartDate"),
            CompletionDate:   this.getInternalName("Completion Date",  "CompletionDate"),
            CreatedByUser:    this.getInternalName("Created By User",  "CreatedByUser"),
            UpdatedByUser:    this.getInternalName("Updated By User",  "UpdatedByUser"),
            CreatedOn:        this.getInternalName("Created On",       "CreatedOn"),
            UpdatedOn:        this.getInternalName("Updated On",       "UpdatedOn"),
            Description:      this.getInternalName("Description",      "Description"),
            Regarding:        this.getInternalName("Regarding",        "Regarding"),
            DueDate:          this.getInternalName("Due Date",         "DueDate"),
            Resolution:       this.getInternalName("Resolution",       "Resolution"),
            EmailNotifications: this.getInternalName("Email Notification", "EmailNotifications"),
        };

        const selects = ["*", "Id", "Title", "Author/Title", "Author/EMail", "Editor/Title", "Editor/EMail", "Created", "Modified", "Attachments", `${names.Status}/Title`, `${names.Status}/Name`, `${names.Category}/Title`, `${names.Category}/Name`, `${names.Classification}/Title`, `${names.Classification}/Name`, `${names.Priority}/Title`, `${names.Priority}/Name`];
        const expands = ["Author", "Editor", "AttachmentFiles", names.Status, names.Category, names.Classification, names.Priority];

        // ─── Use ACTUAL TypeAsString to decide expand strategy ─────────────────
        const safelyAddSelect = (internalName: string): void => {
            if (fieldInternalNames.indexOf(internalName) < 0) return;
            const fieldType = this._fieldTypeMap.get(internalName) || "";

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

        safelyAddSelect(names.Subject);
        safelyAddSelect(names.Status);
        safelyAddSelect(names.Category);
        safelyAddSelect(names.Classification);
        safelyAddSelect(names.Priority);
        safelyAddSelect(names.TaskOwner);
        safelyAddSelect(names.AssigneeInternal);
        safelyAddSelect(names.AssigneeExternal);
        safelyAddSelect(names.StartDate);
        safelyAddSelect(names.CompletionDate);
        safelyAddSelect(names.CompletedPercent);
        safelyAddSelect(names.CreatedByUser);
        safelyAddSelect(names.UpdatedByUser);
        safelyAddSelect(names.CreatedOn);
        safelyAddSelect(names.UpdatedOn);
        safelyAddSelect(names.Description);
        safelyAddSelect(names.Regarding);
        safelyAddSelect(names.DueDate);
        safelyAddSelect(names.Resolution);
        safelyAddSelect(names.EmailNotifications);

        try {
            const itemIterator = this._sp.web.lists.getByTitle(this._listTitle).items
                .select(...selects)
                .expand(...expands)
                .top(2000)
                .orderBy("Id", true);
            
            const rawItems: any[] = [];
            for await (const items of itemIterator) {
                rawItems.push(...items);
                if (rawItems.length >= 5000) break;
            }

            const items: any[] = rawItems;
            return items.map((item: any) => {

                // ── Handles Lookup (object) and Choice (string) ──────────────
                const getLookup = (n: string): { Id: number; Title: string; Name?: string } | undefined => {
                    const v = item[n];
                    if (v === null || v === undefined || v === "") return undefined;
                    if (typeof v === "object") return { Id: v.Id || 0, Title: v.Title || v.Name || "", Name: v.Name };
                    if (typeof v === "string") return { Id: 0, Title: v };
                    return undefined;
                };

                // ── Handles Person field ──────────────────────────────────────
                const getPerson = (n: string): { Id: number; Title: string; EMail: string } | undefined => {
                    const v = item[n];
                    if (!v || typeof v !== "object") return undefined;
                    return { Id: v.Id || 0, Title: v.Title || "", EMail: v.EMail || "" };
                };

                return {
                    ...item,
                    Id: item.Id,
                    Title: item[names.Subject] || item.Title || "",
                    Description: item[names.Description],
                    Status:         getLookup(names.Status),
                    Category:       getLookup(names.Category),
                    Classification: getLookup(names.Classification),
                    Priority:       getLookup(names.Priority),
                    TaskOwner:        getPerson(names.TaskOwner),
                    AssigneeInternal: getPerson(names.AssigneeInternal),
                    AssigneeExternal: getPerson(names.AssigneeExternal),
                    Regarding:        item[names.Regarding],
                    DueDate:          item[names.DueDate],
                    StartDate:        item[names.StartDate],
                    CompletionDate:   item[names.CompletionDate],
                    CompletedPercent: item[names.CompletedPercent],
                    EmailNotifications: item[names.EmailNotifications],
                    Author:  getPerson(names.CreatedByUser)  || item.Author,
                    Editor:  getPerson(names.UpdatedByUser)  || item.Editor,
                    Created: item[names.CreatedOn]  || item.Created,
                    Modified: item[names.UpdatedOn] || item.Modified,
                    Resolution: item[names.Resolution],
                    AttachmentFiles: item.AttachmentFiles || [],
                };
            });
        } catch (error) {
            console.error(`Query failed on ${this._listTitle}:`, error);
            this._lastError = `Fetch failed: ${error.message || JSON.stringify(error)}`;
            try {
                const fallback = await this._sp.web.lists
                    .getByTitle(this._listTitle).items.select("Id", "Title", "Created")();
                return fallback.map(i => ({ Id: i.Id, Title: i.Title, Created: i.Created }));
            } catch { return []; }
        }
    }

    public async addToDoItem(item: any): Promise<any> {
        await this.init();
        try {
            const fieldInternalNames = Array.from(this._fieldMap.values());
            const cleaned: any = {};

            const currentUser = await this._sp.web.currentUser();
            const now = new Date().toISOString();
            const createdOnInt    = this.getInternalName("Created On",       "CreatedOn");
            const createdByInt    = this.getInternalName("Created By User",  "CreatedByUser");
            const updatedOnInt    = this.getInternalName("Updated On",       "UpdatedOn");
            const updatedByInt    = this.getInternalName("Updated By User",  "UpdatedByUser");

            if (fieldInternalNames.indexOf(createdOnInt)  > -1) item[createdOnInt]          = now;
            if (fieldInternalNames.indexOf(createdByInt)  > -1) item[`${createdByInt}Id`]   = currentUser.Id;
            if (fieldInternalNames.indexOf(updatedOnInt)  > -1) item[updatedOnInt]           = now;
            if (fieldInternalNames.indexOf(updatedByInt)  > -1) item[`${updatedByInt}Id`]   = currentUser.Id;

            Object.keys(item).forEach(key => {
                const baseKey = key.endsWith("Id") ? key.slice(0, -2) : key;
                if (
                    fieldInternalNames.indexOf(key)     > -1 ||
                    fieldInternalNames.indexOf(baseKey) > -1 ||
                    key === "Title"
                ) {
                    if (item[key] !== undefined && item[key] !== null) cleaned[key] = item[key];
                }
            });

            return await this._sp.web.lists.getByTitle(this._listTitle).items.add(cleaned);
        } catch (error) {
            this._lastError = `Save failed: ${error.message || JSON.stringify(error)}`;
            throw error;
        }
    }

    public async updateToDoItem(id: number, item: any): Promise<any> {
        await this.init();
        try {
            const fieldInternalNames = Array.from(this._fieldMap.values());
            const cleaned: any = {};

            const currentUser  = await this._sp.web.currentUser();
            const now          = new Date().toISOString();
            const updatedOnInt = this.getInternalName("Updated On",       "UpdatedOn");
            const updatedByInt = this.getInternalName("Updated By User",  "UpdatedByUser");

            if (fieldInternalNames.indexOf(updatedOnInt) > -1) item[updatedOnInt]          = now;
            if (fieldInternalNames.indexOf(updatedByInt) > -1) item[`${updatedByInt}Id`]  = currentUser.Id;

            Object.keys(item).forEach(key => {
                const baseKey = key.endsWith("Id") ? key.slice(0, -2) : key;
                if (
                    fieldInternalNames.indexOf(key)     > -1 ||
                    fieldInternalNames.indexOf(baseKey) > -1 ||
                    key === "Title"
                ) {
                    if (item[key] !== undefined && item[key] !== null) cleaned[key] = item[key];
                }
            });

            return await this._sp.web.lists.getByTitle(this._listTitle).items.getById(id).update(cleaned);
        } catch (error) {
            if (error.data) {
                const data = await error.data.json();
                this._lastError = `Update failed: ${data.odata.error.message.value}`;
            }
            throw error;
        }
    }

    public async getLookupOptions(listUrlName: string, displayField: string = "Title"): Promise<ILookupOption[]> {
        try {
            const realTitle = await this.findListTitle(listUrlName);
            const items = await this._sp.web.lists.getByTitle(realTitle).items
                .select("Id", displayField, "Title")();
            return items.map(item => ({ Id: item.Id, Title: item[displayField] || item.Title }));
        } catch { return []; }
    }

    public async getAttachments(itemId: number): Promise<IAttachment[]> {
        await this.init();
        try { return await this._sp.web.lists.getByTitle(this._listTitle).items.getById(itemId).attachmentFiles(); }
        catch { return []; }
    }

    public async uploadAttachment(itemId: number, file: File): Promise<void> {
        await this.init();
        await this._sp.web.lists.getByTitle(this._listTitle).items.getById(itemId).attachmentFiles.add(file.name, file);
    }

    public async deleteAttachment(itemId: number, fileName: string): Promise<void> {
        await this.init();
        await this._sp.web.lists.getByTitle(this._listTitle).items.getById(itemId).attachmentFiles.getByName(fileName).delete();
    }

    public async deleteToDoItem(id: number): Promise<void> {
        await this.init();
        await this._sp.web.lists.getByTitle(this._listTitle).items.getById(id).delete();
    }

    private async findListTitle(urlName: string): Promise<string> {
        try {
            const lists = await this._sp.web.lists.select("Title", "RootFolder/Name").expand("RootFolder")();
            const match = lists.find(l =>
                l.Title.toLowerCase() === urlName.toLowerCase() ||
                l.RootFolder.Name.toLowerCase() === urlName.toLowerCase() ||
                l.Title.replace(/\s+/g, '').toLowerCase() === urlName.toLowerCase()
            );
            return match ? match.Title : urlName;
        } catch { return urlName; }
    }
}
