import { WebPartContext } from "@microsoft/sp-webpart-base";
import { BaseSPService } from "../../../common/services/BaseSPService";
import { IActivityTaskItem, IActivityTaskHazard, ISPLookup } from "../models/IActivityTask";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export class ActivityTaskService extends BaseSPService {
    private readonly MAIN_LIST = "ActivityTaskRegister";
    private readonly HAZARD_LIST = "ActivityTaskHazardType";

    constructor(context: WebPartContext) {
        super(context);
    }

    public async getActivityTasksPaged(
        page: number,
        pageSize: number,
        searchQuery?: string,
        sortField: string = "Id",
        isAscending: boolean = true
    ): Promise<IActivityTaskItem[]> {
        await this.init(this.MAIN_LIST);
        const skipCount = (page - 1) * pageSize;

        try {
            let query = this._sp.web.lists.getByTitle(this._listTitle).items
                .select(
                    "*",
                    "Id",
                    "Activity/Id", "Activity/Title",
                    "WorkZone/Id", "WorkZone/Title",
                    "ResponsiblePersons/Id", "ResponsiblePersons/Title",
                    "Author/Title", "Editor/Title"
                )
                .expand("Activity", "WorkZone", "ResponsiblePersons", "Author", "Editor")
                .orderBy(sortField, isAscending)
                .top(pageSize);

            if (skipCount > 0) {
                query = query.skip(skipCount);
            }

            if (searchQuery) {
                // Simplistic search for now, as per guidelines for performance on large lists
                // In a real scenario, this would use the search API or filtered query
                query = query.filter(`substringof('${searchQuery}', Task) or substringof('${searchQuery}', Title)`);
            }

            const items = await query();
            return items.map(item => this.mapToActivityTask(item));
        } catch (error) {
            console.error("[ActivityTaskService] getActivityTasksPaged failed", error);
            return [];
        }
    }

    public async getTotalCount(searchQuery?: string): Promise<number> {
        await this.init(this.MAIN_LIST);
        try {
            if (searchQuery) {
                const results = await this._sp.web.lists.getByTitle(this._listTitle).items
                    .filter(`substringof('${searchQuery}', Task)`)
                    .select("Id")();
                return results.length;
            }
            const list = await this._sp.web.lists.getByTitle(this._listTitle).select("ItemCount")();
            return list.ItemCount;
        } catch (error) {
            return 0;
        }
    }

    public async getHazards(taskId: number): Promise<IActivityTaskHazard[]> {
        await this.init(this.HAZARD_LIST);
        try {
            const items = await this._sp.web.lists.getByTitle(this._listTitle).items
                .filter(`ActivityTaskRegisterId eq ${taskId}`)
                .select("*", "Hazard/Id", "Hazard/Title")
                .expand("Hazard")();
            
            return items.map(item => ({
                Id: item.Id,
                Title: item.Title,
                ActivityTaskRegisterId: item.ActivityTaskRegisterId,
                Hazard: item.Hazard ? { Id: item.Hazard.Id, Title: item.Hazard.Title } : undefined
            }));
        } catch (error) {
            console.error("[ActivityTaskService] getHazards failed", error);
            return [];
        }
    }

    public async addActivityTask(item: any): Promise<any> {
        await this.init(this.MAIN_LIST);
        return this._sp.web.lists.getByTitle(this._listTitle).items.add(item);
    }

    public async updateActivityTask(id: number, item: any): Promise<any> {
        await this.init(this.MAIN_LIST);
        return this._sp.web.lists.getByTitle(this._listTitle).items.getById(id).update(item);
    }

    public async deleteActivityTask(id: number): Promise<void> {
        await this.init(this.MAIN_LIST);
        await this._sp.web.lists.getByTitle(this._listTitle).items.getById(id).delete();
    }

    public async getLookupOptions(listName: string): Promise<ISPLookup[]> {
        try {
            const items = await this._sp.web.lists.getByTitle(listName).items.select("Id", "Title").top(500)();
            return items.map(i => ({ Id: i.Id, Title: i.Title }));
        } catch (error) {
            console.error(`[ActivityTaskService] getLookupOptions failed for ${listName}`, error);
            return [];
        }
    }

    public async getChoiceOptions(fieldName: string): Promise<string[]> {
        await this.init(this.MAIN_LIST);
        try {
            const field: any = await this._sp.web.lists.getByTitle(this._listTitle).fields.getByInternalNameOrTitle(fieldName).select("Choices")();
            return field.Choices || [];
        } catch (error) {
            return [];
        }
    }

    private mapToActivityTask(item: any): IActivityTaskItem {
        return {
            ...item,
            ActivityValue: item.Activity?.Title || "",
            WorkZoneValue: item.WorkZone?.Title || "",
            ResponsiblePersonsValue: item.ResponsiblePersons?.Title || ""
        };
    }
}
