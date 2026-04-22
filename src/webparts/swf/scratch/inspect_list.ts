import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export async function inspectList(context: WebPartContext, listTitle: string) {
    const sp = spfi().using(SPFx(context));
    const fields = await sp.web.lists.getByTitle(listTitle).fields.select("Title", "InternalName", "TypeAsString")();
    console.log(`Fields for ${listTitle}:`, JSON.stringify(fields, null, 2));
}
