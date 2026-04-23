import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx, SPFI } from "@pnp/sp";
import { IActivityTaskItem, IActivityTaskHazard, ISPLookup } from "../models/IActivityTask";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";

/**
 * ActivityTaskService
 * Uses web.getList(serverRelativeUrl) for KNOWN lists so the list display title
 * doesn't matter — only the URL name (folder inside /Lists/) matters.
 * The URL name comes from the SharePoint URL the user provided:
 *   /sites/HSEQ/Lists/ActivityTaskRegister
 *   /sites/HSEQ/Lists/ActivityTaskHazardType
 */
export class ActivityTaskService {
    private _sp: SPFI;
    private _siteRelUrl: string;   // e.g. "/sites/HSEQ"

    // Internal URL names (from the /Lists/<name>/ segment of the SP URL)
    private readonly MAIN_URL    = "ActivityTaskRegister";
    private readonly HAZARD_URL  = "ActivityTaskHazardType";

    constructor(context: WebPartContext) {
        this._sp = spfi().using(SPFx(context));
        // "/sites/HSEQ" — no trailing slash
        this._siteRelUrl = context.pageContext.web.serverRelativeUrl.replace(/\/$/, "");
    }

    /** Build absolute server-relative URL for a /Lists/<urlName> list */
    private lUrl(urlName: string): string {
        return `${this._siteRelUrl}/Lists/${urlName}`;
    }

    private get mainList()   { return this._sp.web.getList(this.lUrl(this.MAIN_URL));   }
    private get hazardList() { return this._sp.web.getList(this.lUrl(this.HAZARD_URL)); }

    // ─────────────────────────────────────────────────────────────────────────────
    // MAIN LIST — read
    // ─────────────────────────────────────────────────────────────────────────────
    public async getActivityTasksPaged(
        page:        number,
        pageSize:    number,
        searchQuery?: string,
        sortField:   string  = "Id",
        isAscending: boolean = true
    ): Promise<IActivityTaskItem[]> {
        const skip = (page - 1) * pageSize;
        try {
            let q = this.mainList.items
                .select(
                    "Id", "Title", "Task",
                    "Activity/Id",       "Activity/Title",
                    "BusinessProfile/Id","BusinessProfile/Title",
                    "WorkZone/Id",       "WorkZone/Title",
                    "Hazard/Id",         "Hazard/Title",
                    "Consequence", "Likelihood", "RiskRanking",
                    "RevisedConsequence", "RevisedLikelihood", "RevisedRanking",
                    "HighRiskWork", "Active",
                    "ActivityId", "BusinessProfileId", "WorkZoneId", "HazardId"
                )
                .expand("Activity", "BusinessProfile", "WorkZone", "Hazard")
                .orderBy(sortField, isAscending)
                .top(pageSize);

            if (skip > 0) q = q.skip(skip);
            if (searchQuery) {
                q = q.filter(`substringof('${searchQuery}',Task) or substringof('${searchQuery}',Title)`);
            }

            const rows = await q();
            return rows.map(r => this.map(r));
        } catch (err) {
            console.error("[ATService] paged fetch failed, trying simple:", err);
            return this.simpleFetch(page, pageSize);
        }
    }

    private async simpleFetch(page: number, pageSize: number): Promise<IActivityTaskItem[]> {
        try {
            const skip = (page - 1) * pageSize;
            let q = this.mainList.items
                .select("Id","Title","Task","Consequence","Likelihood","RiskRanking","HighRiskWork","Active")
                .orderBy("Id", true).top(pageSize);
            if (skip > 0) q = q.skip(skip);
            const rows = await q();
            return rows.map(r => ({ ...r } as IActivityTaskItem));
        } catch (e) {
            console.error("[ATService] simple fetch also failed:", e);
            return [];
        }
    }

    public async getTotalCount(searchQuery?: string): Promise<number> {
        try {
            if (searchQuery) {
                const r = await this.mainList.items
                    .filter(`substringof('${searchQuery}',Task) or substringof('${searchQuery}',Title)`)
                    .select("Id")();
                return r.length;
            }
            const info = await this.mainList.select("ItemCount")();
            return (info as any).ItemCount ?? 0;
        } catch (e) {
            console.error("[ATService] getTotalCount failed:", e);
            return 0;
        }
    }

    // ─────────────────────────────────────────────────────────────────────────────
    // MAIN LIST — write
    // ─────────────────────────────────────────────────────────────────────────────
    public async addActivityTask(data: Partial<IActivityTaskItem>): Promise<any> {
        const p = this.payload(data);
        console.log("[ATService] ADD payload:", JSON.stringify(p));
        return this.mainList.items.add(p);
    }

    public async updateActivityTask(id: number, data: Partial<IActivityTaskItem>): Promise<any> {
        const p = this.payload(data);
        console.log("[ATService] UPDATE payload:", JSON.stringify(p));
        return this.mainList.items.getById(id).update(p);
    }

    public async deleteActivityTask(id: number): Promise<void> {
        await this.mainList.items.getById(id).delete();
    }

    /** Only send safe SP-writable fields; use *Id suffix for lookups */
    private payload(d: Partial<IActivityTaskItem>): Record<string, any> {
        const p: Record<string, any> = {};
        const set = (k: string, v: any) => { if (v !== undefined && v !== null) p[k] = v; };

        set("Task",              d.Task);
        set("Consequence",       d.Consequence);
        set("Likelihood",        d.Likelihood);
        set("RiskRanking",       d.RiskRanking);
        set("RevisedConsequence",d.RevisedConsequence);
        set("RevisedLikelihood", d.RevisedLikelihood);
        set("RevisedRanking",    d.RevisedRanking);
        set("HighRiskWork",      d.HighRiskWork);
        set("Active",            d.Active);

        // Lookup IDs — SP REST expects the "<FieldName>Id" suffix
        if (d.ActivityId)        set("ActivityId",        d.ActivityId);
        if (d.BusinessProfileId) set("BusinessProfileId", d.BusinessProfileId);
        if (d.WorkZoneId)        set("WorkZoneId",        d.WorkZoneId);
        if (d.HazardId)          set("HazardId",          d.HazardId);

        return p;
    }

    // ─────────────────────────────────────────────────────────────────────────────
    // HAZARD sub-list
    // ─────────────────────────────────────────────────────────────────────────────
    public async getHazards(taskId: number): Promise<IActivityTaskHazard[]> {
        try {
            const rows = await this.hazardList.items
                .filter(`ActivityTaskRegisterId eq ${taskId}`)
                .select("Id","Title","ActivityTaskRegisterId")();
            return rows.map(r => ({ Id: r.Id, Title: r.Title||"", ActivityTaskRegisterId: r.ActivityTaskRegisterId }));
        } catch (e) {
            console.error("[ATService] getHazards failed:", e);
            return [];
        }
    }

    public async addHazard(taskId: number, title: string): Promise<IActivityTaskHazard> {
        const r = await this.hazardList.items.add({ Title: title, ActivityTaskRegisterId: taskId });
        return { Id: r.data.Id, Title: title, ActivityTaskRegisterId: taskId };
    }

    public async deleteHazard(id: number): Promise<void> {
        await this.hazardList.items.getById(id).delete();
    }

    // ─────────────────────────────────────────────────────────────────────────────
    // LOOKUP reference lists
    // These use web.lists.getByTitle() because we know their DISPLAY TITLES from
    // the column definitions in SharePoint.
    // ─────────────────────────────────────────────────────────────────────────────
    public async getLookupOptions(
        listDisplayTitle: string,
        fieldDisplayName?: string
    ): Promise<ISPLookup[]> {
        try {
            const list = this._sp.web.lists.getByTitle(listDisplayTitle);

            // Discover the internal field name if a display name was given
            let internalName = "Title";
            if (fieldDisplayName && fieldDisplayName !== "Title") {
                const found = await list.fields
                    .filter(`Title eq '${fieldDisplayName}'`)
                    .select("InternalName")();
                if (found.length > 0) internalName = found[0].InternalName;
                console.log(`[ATService] "${listDisplayTitle}" / "${fieldDisplayName}" → internal: "${internalName}"`);
            }

            const sel = internalName === "Title" ? ["Id","Title"] : ["Id","Title",internalName];
            const rows = await list.items.select(...sel).top(500)();

            return rows.map(r => ({
                Id: r.Id,
                Title: String(
                    (r[internalName] !== undefined && r[internalName] !== null && r[internalName] !== "")
                        ? r[internalName]
                        : (r.Title || `Item ${r.Id}`)
                )
            })).filter(x => x.Title.trim() !== "");
        } catch (e) {
            console.error(`[ATService] getLookupOptions("${listDisplayTitle}"/"${fieldDisplayName}") failed:`, e);
            return [];
        }
    }

    // ─────────────────────────────────────────────────────────────────────────────
    // CHOICE columns — read from the MAIN list's field schema
    // ─────────────────────────────────────────────────────────────────────────────
    public async getChoiceOptions(fieldDisplayName: string): Promise<string[]> {
        try {
            // Try by display title first
            const byTitle = await this.mainList.fields
                .filter(`Title eq '${fieldDisplayName}'`)
                .select("Choices")();
            if (byTitle.length > 0 && (byTitle[0] as any).Choices?.length) {
                return (byTitle[0] as any).Choices as string[];
            }
            // Fallback: by internal name
            const field: any = await this.mainList.fields
                .getByInternalNameOrTitle(fieldDisplayName)
                .select("Choices")();
            return field.Choices || [];
        } catch (e) {
            console.error(`[ATService] getChoiceOptions("${fieldDisplayName}") failed:`, e);
            return [];
        }
    }

    // ─────────────────────────────────────────────────────────────────────────────
    // Mapping helper
    // ─────────────────────────────────────────────────────────────────────────────
    private map(item: any): IActivityTaskItem {
        return {
            ...item,
            ActivityValue:        item.Activity?.Title        || "",
            ActivityId:           item.ActivityId             ?? item.Activity?.Id,
            BusinessProfileValue: item.BusinessProfile?.Title || "",
            BusinessProfileId:    item.BusinessProfileId      ?? item.BusinessProfile?.Id,
            WorkZoneValue:        item.WorkZone?.Title        || "",
            WorkZoneId:           item.WorkZoneId             ?? item.WorkZone?.Id,
            HazardValue:          item.Hazard?.Title          || "",
            HazardId:             item.HazardId               ?? item.Hazard?.Id
        };
    }
}
