import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import "@pnp/sp/fields";
import "@pnp/sp/search";

export class BaseSPService {
    protected _sp: SPFI;
    protected _fieldMap: Map<string, string> = new Map();
    protected _fieldTypeMap: Map<string, string> = new Map();
    protected _listTitle: string = "";
    protected _listId: string = "";
    protected _listUrl: string = "";
    protected _isInitialized: boolean = false;
    protected _lastError: string = "";

    constructor(protected context: WebPartContext) {
        this._sp = spfi().using(SPFx(context));
    }

    /**
     * Initializes the service by finding the list and its fields.
     * @param listInternalName The internal name or title of the list to target.
     */
    public async init(listInternalName: string): Promise<void> {
        if (this._isInitialized && this._listTitle.toLowerCase() === listInternalName.toLowerCase()) return;
        
        try {
            const lists = await this._sp.web.lists.select("Title", "RootFolder/Name").expand("RootFolder")();
            const match = lists.find(l =>
                l.Title.toLowerCase() === listInternalName.toLowerCase() ||
                l.RootFolder.Name.toLowerCase() === listInternalName.toLowerCase() ||
                l.Title.replace(/\s+/g, '').toLowerCase() === listInternalName.toLowerCase()
            );
            
            this._listTitle = match ? match.Title : listInternalName;

            const listData = await this._sp.web.lists.getByTitle(this._listTitle)
                .select("Id", "Title", "RootFolder/ServerRelativeUrl")
                .expand("RootFolder")();
                
            this._listId = listData.Id;
            this._listUrl = listData.RootFolder.ServerRelativeUrl;
            
            // Normalize URL for Search
            const webUrl = this.context.pageContext.web.absoluteUrl;
            const siteUrl = webUrl.substring(0, webUrl.indexOf(this.context.pageContext.web.serverRelativeUrl));
            this._listUrl = siteUrl + this._listUrl;
            if (this._listUrl.endsWith('/')) this._listUrl = this._listUrl.slice(0, -1);

            const fields = await this._sp.web.lists.getByTitle(this._listTitle).fields
                .select("Title", "InternalName", "TypeAsString")();

            this._fieldMap.clear();
            this._fieldTypeMap.clear();
            
            fields.forEach(f => {
                const title = (f.Title || "").trim();
                const internal = (f.InternalName || "").trim();
                const type = (f.TypeAsString || "").trim();

                this._fieldMap.set(title, internal);
                this._fieldTypeMap.set(internal, type);
            });

            this._isInitialized = true;
        } catch (error) {
            console.error(`[BaseSPService] Initialization Failed for ${listInternalName}:`, error);
            this._lastError = error.message;
            throw error;
        }
    }

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

    public getLastError(): string { return this._lastError; }
}
