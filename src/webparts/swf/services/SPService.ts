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
    private _isInitialized: boolean = false;
    private _lastError: string = "";

    constructor(context: WebPartContext) {
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

            const listData = await this._sp.web.lists.getByTitle(this._listTitle).select("Id", "Title")();
            this._listId = listData.Id;

            console.log(`[SPService] List Identified: "${this._listTitle}" (ID: ${this._listId})`);

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



    public async getToDoTotalCount(searchQuery?: string): Promise<number> {
        await this.init();
        try {
            if (searchQuery) {
                const results: SearchResults = await this._sp.search({
                    Querytext: `${searchQuery} ListId:${this._listId}`,
                    RowLimit: 0,
                    SelectProperties: ["ListItemID"]
                });
                return results.TotalRows;
            } else {
                // For the total count (no search), the list metadata is the most reliable scale-safe source.
                const listData = await this._sp.web.lists.getByTitle(this._listTitle).select("ItemCount")();
                return listData.ItemCount || 0;
            }
        } catch (e) {
            console.warn("[SPService] Could not fetch total count:", e);
            // Emergency fallback for large lists
            const searchFallback = await this._sp.search({
                Querytext: `ListId:${this._listId}`,
                RowLimit: 0,
                SelectProperties: ["ListItemID"]
            });
            return searchFallback.TotalRows || 0;
        }
        return 0;
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

            // Decide between Search (Threshold-safe) and REST (Real-time)
            const useSearch = !!searchQuery || skipCount >= 5000;

            if (useSearch) {
                // Use Search API for threshold-safe deep paging or keyword search
                const queryText = searchQuery 
                    ? `${searchQuery} ListId:${this._listId}` 
                    : `ListId:${this._listId}`;

                const searchResults: SearchResults = await this._sp.search({
                    Querytext: queryText,
                    StartRow: skipCount,
                    RowLimit: pageSize,
                    SelectProperties: ["ListItemID"],
                    // Simple sorting mapping for search
                    SortList: [{ Property: sortField === 'Id' ? 'ListItemID' : sortField, Direction: isAscending ? 0 : 1 }]
                });

                if (searchResults.TotalRows === 0) return [];
                const ids = searchResults.PrimarySearchResults.map((r: any) => r.ListItemID).filter((id: any) => id);

                if (ids.length === 0) return [];

                // Hydrate results via REST using IDs (LVT-safe)
                const idFilter = ids.map((id: any) => `Id eq ${id}`).join(' or ');
                rawItems = await this._sp.web.lists.getByTitle(this._listTitle).items
                    .filter(idFilter)
                    .select(...selects)
                    .expand(...expands)();
                
                // REST filter doesn't guarantee search order, so we re-sort to match search results
                rawItems.sort((a, b) => {
                    const idxA = ids.indexOf(a.Id.toString());
                    const idxB = ids.indexOf(b.Id.toString());
                    return idxA - idxB;
                });

            } else {
                // Use REST for real-time visibility on the first 5000 items (no crawl delay)
                let baseQuery = this._sp.web.lists
                    .getByTitle(this._listTitle).items
                    .select(...selects)
                    .expand(...expands)
                    .top(pageSize)
                    .orderBy(realSortField, isAscending);

                rawItems = await (skipCount > 0 ? baseQuery.skip(skipCount)() : baseQuery());
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
            
            // Mapping sort field to internal name if needed (Unused in these paths)
            /* let realSortField = sortField;
            if (sortField === "TaskOwner") realSortField = `${names.TaskOwner}/Title`;
            else if (sortField === "Subject")  realSortField = names.Subject;
            else if (sortField === "Regarding") realSortField = names.Regarding; */

            if (searchQuery) {
                // For export search, we might need more items than pageSize
                // Use Search to find all relevant IDs
                const searchResults: SearchResults = await this._sp.search({
                    Querytext: `${searchQuery} ListId:${this._listId}`,
                    RowLimit: 5000,
                    SelectProperties: ["ListItemID"]
                });

                if (searchResults.TotalRows === 0) return [];
                const ids = searchResults.PrimarySearchResults.map((r: any) => r.ListItemID).filter((id: any) => id);

                if (ids.length === 0) return [];

                // In a real large-scale scenario, we'd batch these IDs too.
                // For now, we fetch them in a single filtered call which is safe for 5000 IDs.
                const idFilter = ids.map((id: any) => `Id eq ${id}`).join(' or ');
                rawItems = await this._sp.web.lists.getByTitle(this._listTitle).items
                    .filter(idFilter)
                    .select(...selects)
                    .expand(...expands)();

            } else {
                // Fetch in batches for mass export to avoid LVT (PnP JS v4 Async Iterator)
                const itemIterator = this._sp.web.lists.getByTitle(this._listTitle).items
                    .select(...selects)
                    .expand(...expands)
                    .top(2000)
                    .orderBy("Id", true);
                
                for await (const items of itemIterator) {
                    rawItems.push(...items);
                    if (rawItems.length >= 10000) break; // Safety cap
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
