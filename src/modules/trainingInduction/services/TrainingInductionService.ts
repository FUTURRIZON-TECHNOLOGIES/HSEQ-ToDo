import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/attachments';
import '@pnp/sp/fields';
import '@pnp/sp/files';
import '@pnp/sp/files/web';
import '@pnp/sp/folders';
import { IAttachment, ILookupOption, ITrainingInductionItem } from '../models/ITrainingInductionItem';

export class TrainingInductionService {
    private _sp: SPFI;
    private _fieldMaps: Map<string, Map<string, string>> = new Map();
    private _fieldTypeMaps: Map<string, Map<string, string>> = new Map();
    private _listTitles: Map<string, string> = new Map();
    private _isInitialized: Map<string, boolean> = new Map();
    private _listFieldInternalNames: Map<string, string[]> = new Map();
    private _lookupDisplayFields: Map<string, Map<string, string>> = new Map();
    private _lastError: string = "";

    constructor(context: WebPartContext) {
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

                console.log(`[SPService] Field: ${title} -> ${internal} (${type})`);

                if (type === "Lookup" || type === "LookupMulti") {
                    if ((f as any).LookupField) {
                        displayFieldMap.set(internal, (f as any).LookupField);
                        console.log(`[SPService] Lookup field "${internal}" display field is "${(f as any).LookupField}"`);
                    }
                }
            });

            this._fieldMaps.set(urlName, fieldMap);
            this._fieldTypeMaps.set(urlName, fieldTypeMap);
            this._lookupDisplayFields.set(urlName, displayFieldMap);
            this._listFieldInternalNames.set(urlName, internalNames);
            this._isInitialized.set(urlName, true);
        } catch (error) {
            console.error(`[SPService] Initialization Failed for ${urlName}:`, error);
            this._listTitles.set(urlName, urlName);
            this._fieldMaps.set(urlName, new Map());
            this._fieldTypeMaps.set(urlName, new Map());
            this._isInitialized.set(urlName, true);
        }
    }

    public getLastError(): string { return this._lastError; }

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

    private normalizeFieldToken(value: string): string {
        return (value || "")
            .toLowerCase()
            .replace(/_x[0-9a-f]{4}_/gi, "")
            .replace(/[^a-z0-9]/g, "");
    }

    private findFieldByAliases(
        fields: Array<{ Title?: string; InternalName?: string; TypeAsString?: string }>,
        aliases: string[]
    ): { Title?: string; InternalName?: string; TypeAsString?: string } | undefined {
        const normalizedAliases = aliases.map(alias => this.normalizeFieldToken(alias));
        return fields.find(field => {
            const title = this.normalizeFieldToken(field.Title || "");
            const internalName = this.normalizeFieldToken(field.InternalName || "");
            return normalizedAliases.indexOf(title) > -1 || normalizedAliases.indexOf(internalName) > -1;
        });
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

            // Fetch fields to verify displayField with robust normalization
            const fields = await list.fields.select("InternalName", "Title")();
            const normField = (s: string) => (s || "").toLowerCase().replace(/[^a-z0-9]/g, '');
            const targetNorm = normField(displayField);

            let realField = "Title";
            const match = fields.find(f => 
                (f.InternalName || "").toLowerCase() === displayField.toLowerCase() ||
                (f.Title || "").toLowerCase() === displayField.toLowerCase() ||
                normField(f.InternalName) === targetNorm ||
                normField(f.Title) === targetNorm
            );

            if (match) {
                realField = match.InternalName;
            } else {
                // Fallback checks
                const nameMatch = fields.find(f => normField(f.InternalName) === "name" || normField(f.Title) === "name");
                if (nameMatch) realField = nameMatch.InternalName;
            }

            console.log(`[SPService] Resolved lookup field for ${listUrlName}: ${realField} (requested: ${displayField})`);

            const items = await list.items.select("Id", realField, "Title").top(5000)();
            return items.map(item => ({ 
                Id: item.Id, 
                Title: item[realField] || item.Title || `Record ${item.Id}` 
            }));
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

            // Normalize field names aggressively so SharePoint encodings like _x0020_ still match.
            const normField = (s: string) => this.normalizeFieldToken(s);
            const findField = (name: string) => {
                const normName = normField(name);
                const match = fields.find(f =>
                    (f.InternalName || "").toLowerCase() === name.toLowerCase() ||
                    (f.Title || "").toLowerCase() === name.toLowerCase() ||
                    normField(f.InternalName) === normName ||
                    normField(f.Title) === normName
                );
                if (match) console.log(`[SPService] Resolved "${name}" → "${match.InternalName}" (Type: ${match.TypeAsString})`);
                else console.warn(`[SPService] Could not resolve field "${name}"`);
                return match;
            };

            const recordIdField = this.findFieldByAliases(fields, ["RecordID", "Record ID", "RecordId"]);
            const recordIdInternalName = recordIdField?.InternalName || "RecordID";
            const isRecordIdLookup = recordIdField?.TypeAsString?.includes("Lookup") ?? false;
            console.log(`[SPService] RecordID: "${recordIdInternalName}", isLookup: ${isRecordIdLookup}`);

            const docTypeField = findField("DocumentType") || findField("Document Type");
            const docTypeInternalName = docTypeField?.InternalName || "DocumentType";
            const isDocTypeLookup = docTypeField?.TypeAsString?.includes("Lookup") ?? false;

            const findDescriptionField = () => {
                // Priority: _ExtendedDescription (internal), then "Description"
                return fields.find(f => f.InternalName === "_ExtendedDescription") || 
                       fields.find(f => f.InternalName === "Description") ||
                       fields.find(f => this.normalizeFieldToken(f.Title || "") === "description");
            };
            const descriptionField = findDescriptionField();
            const descriptionInternalName = descriptionField?.InternalName || "";

            // 2. Build Select and Expand arrays dynamically based on resolved fields
            const selectFields = ["Id", "FileLeafRef", "FileRef", "Title", "Created", "FSObjType", "Author/Title"];
            const expandFields = ["Author"];

            // Only expand DocType if it is actually a Lookup field
            if (docTypeField && isDocTypeLookup) {
                selectFields.push(`${docTypeInternalName}/Id`, `${docTypeInternalName}/Name`, `${docTypeInternalName}/Title`);
                expandFields.push(docTypeInternalName);
            } else if (docTypeField) {
                selectFields.push(docTypeInternalName);
            }

            const descriptionSelectKey = descriptionInternalName.startsWith('_')
                ? 'OData_' + descriptionInternalName
                : descriptionInternalName;
            if (descriptionSelectKey) {
                selectFields.push(descriptionSelectKey);
            }

            // For Lookup RecordID, select the hidden numeric field (RecordIDId) directly — no expand
            // needed, and it avoids the $expand+$filter conflict that silently returns 0 in SP Online.
            const recordIdSelectKey = isRecordIdLookup ? `${recordIdInternalName}Id` : recordIdInternalName;
            if (recordIdField) {
                selectFields.push(recordIdSelectKey);
            }

            // Ensure fallback arrays contain everything needed for metadata mapping
            const fallbackSelectFields = [...selectFields];
            const fallbackExpandFields = [...expandFields];

            // For memory filter fallback, we might want to also expand RecordID to be safe
            if (recordIdField && isRecordIdLookup) {
                fallbackSelectFields.push(`${recordIdInternalName}/Id`, `${recordIdInternalName}/Title`);
                fallbackExpandFields.push(recordIdInternalName);
            }

            const matchesCurrentRecord = (item: any): boolean => {
                // Exclude folders
                if (item.FSObjType !== undefined && item.FSObjType !== null && Number(item.FSObjType) !== 0) return false;

                // Extract value from all possible SharePoint property names/structures
                const val = item[recordIdSelectKey] ?? 
                           item[recordIdInternalName] ?? 
                           item[`${recordIdInternalName}Id`] ??
                           (typeof item[recordIdInternalName] === 'object' ? item[recordIdInternalName]?.Id : undefined);

                if (val === undefined || val === null) return false;

                // Robust comparison: handle both primitive values and potential array/object wrappers
                const compareVal = typeof val === 'object' ? (val.Id ?? val.value) : val;
                return String(compareVal).trim() === String(recordId).trim();
            };

            // 3. Fetch items — Prioritize OData filter for performance
            let items: any[] = [];
            const filterExpr = `${recordIdSelectKey} eq ${recordId}`; 
            try {
                items = await list.items
                    .select(...selectFields)
                    .expand(...expandFields)
                    .filter(filterExpr)
                    .orderBy("Created", false)();

                console.log(`[SPService] OData filter "${filterExpr}" returned ${items.length} items.`);
                items = items.filter(matchesCurrentRecord);
            } catch (odataErr) {
                console.warn(`[SPService] OData filter failed:`, odataErr);
                items = []; // Ensure we fall back
            }

            // 4. Fallback: memory filter if OData returned 0 (failed silently or genuinely empty)
            if (items.length === 0) {
                if (recordIdField && isRecordIdLookup) {
                    console.log(`[SPService] Falling back to CAML lookup-id filter for ${recordIdInternalName}=${recordId}...`);
                    try {
                        const viewFields = [
                            "<FieldRef Name='ID' />",
                            "<FieldRef Name='FileLeafRef' />",
                            "<FieldRef Name='FileRef' />",
                            "<FieldRef Name='Title' />",
                            "<FieldRef Name='Created' />",
                            "<FieldRef Name='FSObjType' />",
                            "<FieldRef Name='Author' />",
                            `<FieldRef Name='${recordIdInternalName}' LookupId='TRUE' />`
                        ];

                        if (docTypeField) {
                            viewFields.push(`<FieldRef Name='${docTypeInternalName}' />`);
                        }
                        if (descriptionInternalName) {
                            viewFields.push(`<FieldRef Name='${descriptionInternalName}' />`);
                        }

                        const camlItems = await list.getItemsByCAMLQuery(
                            {
                                ViewXml: `
                                    <View Scope="RecursiveAll">
                                        <Query>
                                            <Where>
                                                <And>
                                                    <Eq>
                                                        <FieldRef Name="FSObjType" />
                                                        <Value Type="Integer">0</Value>
                                                    </Eq>
                                                    <Eq>
                                                        <FieldRef Name="${recordIdInternalName}" LookupId="TRUE" />
                                                        <Value Type="Lookup">${recordId}</Value>
                                                    </Eq>
                                                </And>
                                            </Where>
                                            <OrderBy>
                                                <FieldRef Name="Created" Ascending="FALSE" />
                                            </OrderBy>
                                        </Query>
                                        <ViewFields>
                                            ${viewFields.join("")}
                                        </ViewFields>
                                        <RowLimit>5000</RowLimit>
                                    </View>`
                            },
                            ...fallbackExpandFields.filter((f: string) => f !== 'Author')
                        );

                        console.log(`[SPService] CAML lookup-id filter returned ${camlItems.length} items.`);
                        items = camlItems.filter(matchesCurrentRecord);
                    } catch (camlErr) {
                        console.warn("[SPService] CAML filter failed:", camlErr);
                    }
                }
            }

            if (items.length === 0) {
                console.log(`[SPService] Falling back to memory filter...`);
                try {
                    const allItems = await list.items
                        .select(...fallbackSelectFields)
                        .expand(...fallbackExpandFields)
                        .top(5000)();

                    console.log(`[SPService] Memory Filter: Fetched ${allItems.length} total items.`);
                    if (allItems.length > 0) {
                        console.log("[SPService] Sample RecordID payloads:", allItems.slice(0, 5).map(sample => ({
                            Id: sample.Id,
                            File: sample.FileLeafRef,
                            RecordIDId: sample[recordIdSelectKey],
                            RecordIDLookupId: sample[recordIdInternalName]?.Id,
                            RecordIDLookupTitle: sample[recordIdInternalName]?.Title,
                            FSObjType: sample.FSObjType
                        })));
                    }

                    items = allItems.filter(matchesCurrentRecord);
                    console.log(`[SPService] Memory filter narrowed to ${items.length} items.`);
                } catch (memErr) {
                    console.error("[SPService] Memory filter fetch failed:", memErr);
                    items = [];
                }
            }

            return items.map(item => {
                const dt = item[docTypeInternalName];
                return {
                    Id: item.Id,
                    Title: item.Title || item.FileLeafRef,
                    FileName: item.FileLeafRef,
                    ServerRelativeUrl: item.FileRef,
                    DocumentType: dt ? { Id: dt.Id, Name: dt.Name || dt.Title || "" } : { Id: 0, Name: "" },
                    Description: descriptionSelectKey ? (item[descriptionSelectKey] || "") : "",
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

            const fields = await list.fields.select("Title", "InternalName", "TypeAsString").top(5000)();
            const recordIdField = this.findFieldByAliases(fields, ["RecordID", "Record ID", "RecordId"]);

            const recordIdInternalName = recordIdField?.InternalName || "RecordID";
            const isRecordIdLookup = recordIdField?.TypeAsString?.includes("Lookup") ?? false;

            const result = await list.rootFolder.files.addUsingPath(file.name, file, { Overwrite: true });
            const item = await this._sp.web.getFileByServerRelativePath(result.ServerRelativeUrl).getItem();

            const updatePayload: any = {
                Title: file.name.split('.').slice(0, -1).join('.')
            };

            if (isRecordIdLookup) {
                updatePayload[`${recordIdInternalName}Id`] = recordId;
            } else {
                updatePayload[recordIdInternalName] = recordId;
            }

            console.log(`[SPService] Uploading document for RecordID ${recordId}. Field: ${recordIdInternalName}, isLookup: ${isRecordIdLookup}`);
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

            const findDescriptionField = () => {
                // User specified priority: _ExtendedDescription
                const userChoice = fields.find(f => f.InternalName === "_ExtendedDescription");
                if (userChoice) return userChoice;

                const exact = fields.find(f => f.InternalName === "Description");
                if (exact) return exact;

                const titleMatch = fields.find(f => f.Title === "Description");
                if (titleMatch) return titleMatch;

                return fields.find(f => 
                    this.normalizeFieldToken(f.InternalName || "") === "description" ||
                    this.normalizeFieldToken(f.Title || "") === "description"
                );
            };

            const descField = findDescriptionField();
            const descriptionInternalName = descField?.InternalName;

            const updatePayload: any = {
                Title: metadata.Title
            };

            if (descriptionInternalName) {
                const restKey = descriptionInternalName.startsWith('_')
                    ? 'OData_' + descriptionInternalName
                    : descriptionInternalName;
                updatePayload[restKey] = metadata.Description;
            }

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
